using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
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
        private readonly ILogger<GraphService> _logger;
        private readonly string _authRecordPath;
        private AuthenticationRecord? _authRecord;

        public event Action<string>? ScanStatusChanged;
        public event Action<int>? ItemsProcessed;

        public GraphService(DatabaseService dbService, IConfiguration configuration, ILogger<GraphService> logger)
        {
            _dbService = dbService;
            _configuration = configuration;
            _logger = logger;
            
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            _authRecordPath = Path.Combine(appData, "OneDriveTidy", "auth_record.json");
        }

        public async Task InitializeAsync()
        {
            _logger.LogInformation("Initializing GraphService...");
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
                    _logger.LogInformation("Loaded persisted auth record.");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to load auth record.");
                    // Corrupt record, ignore
                }
            }

            var credential = new InteractiveBrowserCredential(options);

            try 
            {
                _graphClient = new GraphServiceClient(credential, _scopes);
                // Test connection to verify token/record is valid
                var user = await _graphClient.Me.GetAsync();
                ScanStatusChanged?.Invoke($"Connected as: {user?.DisplayName}");
                _logger.LogInformation("Connected as: {DisplayName}", user?.DisplayName);
                return;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Initial auth check failed: {Message}. Retrying interactively...", ex.Message);
                // If the cached credential failed, we need to clear it and try fresh
                if (_authRecord != null)
                {
                    // Clear the record from options and retry
                    options.AuthenticationRecord = null;
                    credential = new InteractiveBrowserCredential(options);
                    
                    // Delete the invalid file
                    try { File.Delete(_authRecordPath); } catch { }
                }
            }

            // If we are here, we need to authenticate interactively
            _authRecord = await credential.AuthenticateAsync();
            
            using var writeStream = new FileStream(_authRecordPath, FileMode.Create, FileAccess.Write);
            await _authRecord.SerializeAsync(writeStream);

            _graphClient = new GraphServiceClient(credential, _scopes);

            // Test connection
            var finalUser = await _graphClient.Me.GetAsync();
            ScanStatusChanged?.Invoke($"Connected as: {finalUser?.DisplayName}");
            _logger.LogInformation($"Connected as: {finalUser?.DisplayName}");
        }

        public bool IsInitialized => _graphClient != null;
        public bool IsScanning { get; private set; } = false;

        public async Task ScanAllFilesAsync(CancellationToken cancellationToken = default)
        {
            if (_graphClient == null) throw new InvalidOperationException("Graph Client not initialized.");
            if (IsScanning) return; // Prevent concurrent scans

            IsScanning = true;
            ScanStatusChanged?.Invoke("Starting scan...");
            _logger.LogInformation("Starting scan...");

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
                    _logger.LogInformation("No delta link found. Starting fresh scan.");
                    deltaResponse = await ExecuteFreshScanAsync(drive.Id, cancellationToken);
                }
                else
                {
                    // Incremental / Resume scan
                    try 
                    {
                        _logger.LogInformation("Resuming scan from delta link.");
                        deltaResponse = await _graphClient.Drives[drive.Id].Items["root"]
                            .Delta
                            .WithUrl(deltaUrl)
                            .GetAsDeltaGetResponseAsync(cancellationToken: cancellationToken);
                    }
                    catch (Exception ex)
                    {
                        // If the link is expired or invalid (410 Gone, etc.), restart fresh
                        var msg = $"Resume link invalid ({ex.Message}). Restarting fresh scan...";
                        ScanStatusChanged?.Invoke(msg);
                        _logger.LogWarning(msg);
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
                        if (processedCount % 100 == 0)
                        {
                            ScanStatusChanged?.Invoke($"Processed {processedCount} items...");
                            _logger.LogInformation("Processed {ProcessedCount} items...", processedCount);
                        }
                        
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
                _logger.LogInformation("Scan complete.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Scan failed.");
                ScanStatusChanged?.Invoke($"Scan failed: {ex.Message}");
                throw;
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
            _logger.LogInformation("Deleted item {ItemId} from Graph.", itemId);
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
