using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using OneDriveTidy.Core.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.Graph.Drives.Item.Items.Item.Delta;

namespace OneDriveTidy.Core.Services
{
    public class GraphService
    {
        private readonly string[] _scopes = new[] { "Files.ReadWrite.All", "User.Read" };

        private GraphServiceClient? _graphClient;
        private readonly DatabaseService _dbService;
        private readonly IConfiguration _configuration;
        private readonly string _authRecordPath;
        private AuthenticationRecord? _authRecord;

        public event Action<string>? ScanStatusChanged;
        public event Action<int>? ItemsProcessed;

        public GraphService(DatabaseService dbService, IConfiguration configuration)
        {
            _dbService = dbService;
            _configuration = configuration;
            
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            _authRecordPath = Path.Combine(appData, "OneDriveTidy", "auth_record.json");
        }

        public async Task InitializeAsync()
        {
            var clientId = _configuration["AzureAd:ClientId"];
            var tenantId = _configuration["AzureAd:TenantId"];

            if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(tenantId))
            {
                throw new InvalidOperationException("AzureAd:ClientId or AzureAd:TenantId is missing in configuration.");
            }

            var options = new InteractiveBrowserCredentialOptions
            {
                ClientId = clientId,
                TenantId = tenantId,
                // RedirectUri is usually http://localhost for desktop apps
                TokenCachePersistenceOptions = new TokenCachePersistenceOptions
                {
                    Name = "OneDriveTidyTokenCache"
                }
            };

            // Try to load persisted authentication record
            if (File.Exists(_authRecordPath))
            {
                try 
                {
                    using var stream = File.OpenRead(_authRecordPath);
                    _authRecord = await AuthenticationRecord.DeserializeAsync(stream);
                    options.AuthenticationRecord = _authRecord;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load auth record: {ex.Message}");
                    // Corrupt record, ignore and re-authenticate
                }
            }

            var credential = new InteractiveBrowserCredential(options);

            // If we don't have a record, authenticate interactively and save it
            if (_authRecord == null)
            {
                _authRecord = await credential.AuthenticateAsync();
                
                using var stream = new FileStream(_authRecordPath, FileMode.Create, FileAccess.Write);
                await _authRecord.SerializeAsync(stream);
            }

            _graphClient = new GraphServiceClient(credential, _scopes);

            // Test connection
            var user = await _graphClient.Me.GetAsync();
            ScanStatusChanged?.Invoke($"Connected as: {user?.DisplayName}");
        }

        public bool IsInitialized => _graphClient != null;
        public bool IsScanning { get; private set; } = false;

        public async Task ScanAllFilesAsync(CancellationToken cancellationToken = default)
        {
            if (_graphClient == null) throw new InvalidOperationException("Graph Client not initialized.");
            if (IsScanning) return; // Prevent concurrent scans

            IsScanning = true;
            ScanStatusChanged?.Invoke("Starting scan...");

            try 
            {
                // We use Delta Query to get all items and changes
                // https://learn.microsoft.com/en-us/graph/api/driveitem-delta?view=graph-rest-1.0
                
                var drive = await _graphClient.Me.Drive.GetAsync(cancellationToken: cancellationToken);
                if (drive?.Id == null) return;

                var deltaUrl = _dbService.GetDeltaLink(); // Load from DB for incremental
                
                DeltaGetResponse? deltaResponse = null;

                if (string.IsNullOrEmpty(deltaUrl))
                {
                    // Fresh scan
                    deltaResponse = await ExecuteFreshScanAsync(drive.Id, cancellationToken);
                }
                else
                {
                    // Incremental / Resume scan
                    try 
                    {
                        deltaResponse = await _graphClient.Drives[drive.Id].Items["root"]
                            .Delta
                            .WithUrl(deltaUrl)
                            .GetAsDeltaGetResponseAsync(cancellationToken: cancellationToken);
                    }
                    catch (Exception ex)
                    {
                        // If the link is expired or invalid (410 Gone, etc.), restart fresh
                        ScanStatusChanged?.Invoke($"Resume link invalid ({ex.Message}). Restarting fresh scan...");
                        deltaResponse = await ExecuteFreshScanAsync(drive.Id, cancellationToken);
                    }
                }

                int processedCount = 0;

                while (deltaResponse != null && !cancellationToken.IsCancellationRequested)
                {
                    var batch = new List<DriveItemModel>();

                    if (deltaResponse.Value != null)
                    {
                        foreach (var item in deltaResponse.Value)
                        {
                            if (item.Deleted != null)
                            {
                                // Item was deleted
                                if (item.Id != null) _dbService.DeleteItem(item.Id);
                            }
                            else if (item.File != null || item.Folder != null)
                            {
                                // Item is file or folder
                                var model = MapToModel(item);
                                batch.Add(model);
                            }
                        }
                    }

                    // Get next page links
                    var nextPageLink = deltaResponse.OdataNextLink;
                    var deltaLink = deltaResponse.OdataDeltaLink;

                    if (batch.Count > 0)
                    {
                        _dbService.UpsertItems(batch);
                        processedCount += batch.Count;
                        ItemsProcessed?.Invoke(processedCount);
                        ScanStatusChanged?.Invoke($"Processed {processedCount} items...");
                        
                        // Save checkpoint immediately so we can resume if stopped
                        if (!string.IsNullOrEmpty(nextPageLink))
                        {
                            _dbService.SaveDeltaLink(nextPageLink);
                        }
                    }

                    if (!string.IsNullOrEmpty(nextPageLink))
                    {
                        deltaResponse = await _graphClient.Drives[drive.Id].Items["root"]
                            .Delta
                            .WithUrl(nextPageLink)
                            .GetAsDeltaGetResponseAsync(cancellationToken: cancellationToken);
                    }
                    else
                    {
                        // End of current sync
                        if (!string.IsNullOrEmpty(deltaLink))
                        {
                            _dbService.SaveDeltaLink(deltaLink);
                        }
                        deltaResponse = null; 
                    }
                }

                ScanStatusChanged?.Invoke("Scan complete.");
            }
            finally
            {
                IsScanning = false;
            }
        }

        private async Task<DeltaGetResponse?> ExecuteFreshScanAsync(string driveId, CancellationToken cancellationToken)
        {
            if (_graphClient == null) return null;
            
            // Clear existing data on fresh scan to ensure consistency? 
            // Optional: _dbService.ClearAll(); 
            // For now, we just upsert/overwrite.

            return await _graphClient.Drives[driveId].Items["root"]
                    .Delta
                    .GetAsDeltaGetResponseAsync(config => 
                    {
                        // config.QueryParameters.Token = ... if we had a token
                    }, cancellationToken);
        }

        public async Task DeleteItemAsync(string itemId)
        {
            if (_graphClient == null) throw new InvalidOperationException("Graph Client not initialized.");
            
            // Correct way to access items via Drive
            var drive = await _graphClient.Me.Drive.GetAsync();
            if (drive?.Id == null) throw new InvalidOperationException("Could not get Drive ID");

            await _graphClient.Drives[drive.Id].Items[itemId].DeleteAsync();
        }

        private DriveItemModel MapToModel(DriveItem item)
        {
            return new DriveItemModel
            {
                Id = item.Id ?? string.Empty,
                Name = item.Name ?? string.Empty,
                ParentId = item.ParentReference?.Id,
                Path = item.ParentReference?.Path, // Note: This might need parsing to remove /drive/root:
                ContentHash = item.File?.Hashes?.Sha1Hash, // Using SHA1 for duplicates
                Size = item.Size,
                CreatedDateTime = item.CreatedDateTime,
                LastModifiedDateTime = item.LastModifiedDateTime,
                IsFolder = item.Folder != null,
                WebUrl = item.WebUrl
            };
        }
    }
}
