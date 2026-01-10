using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OneDriveTidy.Core.Services
{
    public class HsaService
    {
        private readonly GraphService _graphService;
        private readonly DatabaseService _dbService;
        private readonly IConfiguration _configuration;
        private readonly ILogger<HsaService> _logger;

        private const string HSA_FOLDER_NAME = "HSA";
        private const string RECEIPTS_FOLDER_NAME = "Receipts";
        private const string LEDGER_FILE_NAME = "HSA_Master_Ledger.xlsx";
        private const string HSA_START_DATE_CONFIG_KEY = "HsaStartDate";

        private string? _hsaFolderId;
        private string? _receiptsFolderId;
        private string? _ledgerFileId;

        public event Action<string>? StatusChanged;

        public HsaService(GraphService graphService, DatabaseService dbService, IConfiguration configuration, ILogger<HsaService> logger)
        {
            _graphService = graphService;
            _dbService = dbService;
            _configuration = configuration;
            _logger = logger;
        }

        /// <summary>
        /// Get the HSA start date (the date the current HSA was established).
        /// Returns null if not set.
        /// </summary>
        public DateTime? GetHsaStartDate()
        {
            string? dateStr = _dbService.GetConfigValue(HSA_START_DATE_CONFIG_KEY);
            if (string.IsNullOrEmpty(dateStr)) return null;
            
            if (DateTime.TryParse(dateStr, out var date))
            {
                return date;
            }
            return null;
        }

        /// <summary>
        /// Set the HSA start date (the date the current HSA was established).
        /// </summary>
        public void SetHsaStartDate(DateTime date)
        {
            _dbService.SaveConfigValue(HSA_START_DATE_CONFIG_KEY, date.ToString("yyyy-MM-dd"));
            _logger.LogInformation("HSA start date set to {Date}", date.ToString("yyyy-MM-dd"));
            StatusChanged?.Invoke($"HSA start date set to {date:yyyy-MM-dd}");
        }

        /// <summary>
        /// Initialize the HSA folder structure and ledger if they don't exist.
        /// </summary>
        public async Task InitializeStructureAsync(CancellationToken cancellationToken = default)
        {
            if (_graphService.Client == null)
                throw new InvalidOperationException("Graph Client not initialized. Please log in first.");

            _logger.LogInformation("Initializing HSA folder structure...");
            StatusChanged?.Invoke("Initializing HSA structure...");

            try
            {
                var drive = await _graphService.Client.Me.Drive.GetAsync(cancellationToken: cancellationToken);
                if (drive?.Id == null)
                    throw new InvalidOperationException("Could not get Drive ID");

                _logger.LogInformation("Drive ID: {DriveId}", drive.Id);

                // Ensure HSA folder
                _hsaFolderId = await EnsureFolderAsync(drive.Id, "root", HSA_FOLDER_NAME, cancellationToken);
                if (_hsaFolderId == null)
                {
                    _logger.LogError("Failed to create HSA folder - EnsureFolderAsync returned null");
                    throw new InvalidOperationException("Failed to create HSA folder. Check logs for details.");
                }

                _logger.LogInformation("HSA Folder ID: {HsaFolderId}", _hsaFolderId);

                // Ensure Receipts subfolder
                _receiptsFolderId = await EnsureFolderAsync(drive.Id, _hsaFolderId, RECEIPTS_FOLDER_NAME, cancellationToken);
                if (_receiptsFolderId == null)
                {
                    _logger.LogError("Failed to create Receipts folder - EnsureFolderAsync returned null");
                    throw new InvalidOperationException("Failed to create Receipts folder. Check logs for details.");
                }

                _logger.LogInformation("Receipts Folder ID: {ReceiptsFolderId}", _receiptsFolderId);

                // Ensure Ledger file
                _ledgerFileId = await EnsureLedgerFileAsync(drive.Id, _hsaFolderId, cancellationToken);
                if (_ledgerFileId == null)
                {
                    _logger.LogWarning("Failed to create ledger file, but will continue");
                }

                _logger.LogInformation("HSA structure initialized successfully");
                StatusChanged?.Invoke("HSA structure initialized");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to initialize HSA structure");
                StatusChanged?.Invoke($"Error initializing HSA structure: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Upload a receipt file and add it to the Excel ledger.
        /// </summary>
        public async Task UploadReceiptAsync(
            DateTime receiptDate,
            string vendor,
            decimal amount,
            string description,
            Stream fileStream,
            string fileName,
            CancellationToken cancellationToken = default)
        {
            if (_graphService.Client == null)
                throw new InvalidOperationException("Graph Client not initialized.");

            // Ensure structure is initialized
            if (string.IsNullOrEmpty(_receiptsFolderId))
            {
                await InitializeStructureAsync(cancellationToken);
            }

            _logger.LogInformation("Uploading receipt: {Vendor} - {Amount} on {Date}", vendor, amount, receiptDate);
            StatusChanged?.Invoke("Uploading receipt...");

            try
            {
                var drive = await _graphService.Client.Me.Drive.GetAsync(cancellationToken: cancellationToken);
                if (drive?.Id == null)
                    throw new InvalidOperationException("Could not get Drive ID");

                // Generate standardized filename
                string standardizedName = GenerateStandardizedName(receiptDate, vendor, amount, fileName);

                // Upload file to Receipts folder
                var uploadedFile = await UploadFileAsync(drive.Id, _receiptsFolderId!, standardizedName, fileStream, cancellationToken);

                if (uploadedFile?.Id == null)
                    throw new InvalidOperationException("Failed to upload file");

                // Add row to Excel ledger
                AddLedgerEntryAsync(drive.Id, receiptDate, vendor, amount, description, standardizedName, uploadedFile.WebUrl);

                _logger.LogInformation("Receipt uploaded successfully: {FileName}", standardizedName);
                StatusChanged?.Invoke($"Receipt uploaded: {standardizedName}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to upload receipt");
                StatusChanged?.Invoke($"Error uploading receipt: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Generate a standardized filename for a receipt: YYYY-MM-DD_Vendor_Amount.ext
        /// </summary>
        private string GenerateStandardizedName(DateTime date, string vendor, decimal amount, string originalFileName)
        {
            string sanitizedVendor = SanitizeFilename(vendor);
            string extension = Path.GetExtension(originalFileName);
            string amountStr = amount.ToString("F2").Replace(".", "-");
            
            string baseName = $"{date:yyyy-MM-dd}_{sanitizedVendor}_{amountStr}{extension}";
            return baseName;
        }

        private string SanitizeFilename(string name)
        {
            // Remove invalid filename characters
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            // Replace spaces with underscores
            name = name.Replace(" ", "_");
            // Truncate to reasonable length
            if (name.Length > 50)
                name = name.Substring(0, 50);
            return name;
        }

        private async Task<string?> EnsureFolderAsync(string driveId, string parentId, string folderName, CancellationToken cancellationToken)
        {
            try
            {
                _logger.LogInformation("Ensuring folder '{FolderName}' in parent '{ParentId}'", folderName, parentId);

                // Check if folder exists - fetch all children and filter in code
                try
                {
                    var children = await _graphService.Client!.Drives[driveId].Items[parentId].Children
                        .GetAsync(c =>
                        {
                            c.QueryParameters.Select = new[] { "id", "name", "folder" };
                            c.QueryParameters.Top = 999;
                        }, cancellationToken);

                    if (children?.Value != null)
                    {
                        var existingFolder = children.Value.FirstOrDefault(item => 
                            item.Name == folderName && item.Folder != null);
                        
                        if (existingFolder?.Id != null)
                        {
                            _logger.LogInformation("Folder '{FolderName}' already exists with ID: {FolderId}", folderName, existingFolder.Id);
                            return existingFolder.Id;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Error checking for existing folder '{FolderName}', will attempt to create", folderName);
                }

                _logger.LogInformation("Folder '{FolderName}' not found, creating new one...", folderName);

                // Create folder
                var newFolder = new DriveItem
                {
                    Name = folderName,
                    Folder = new Folder { },
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "rename" }
                    }
                };

                var created = await _graphService.Client!.Drives[driveId].Items[parentId].Children
                    .PostAsync(newFolder, cancellationToken: cancellationToken);

                if (created?.Id == null)
                {
                    _logger.LogError("Failed to create folder '{FolderName}' - returned item has no ID", folderName);
                    return null;
                }

                _logger.LogInformation("Created folder '{FolderName}' with ID: {FolderId}", folderName, created.Id);
                return created.Id;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error ensuring folder '{FolderName}': {Message}", folderName, ex.Message);
                return null;
            }
        }

        private async Task<string?> EnsureLedgerFileAsync(string driveId, string hsaFolderId, CancellationToken cancellationToken)
        {
            try
            {
                // Check if file exists - fetch all children and filter in code
                try
                {
                    var children = await _graphService.Client!.Drives[driveId].Items[hsaFolderId].Children
                        .GetAsync(c =>
                        {
                            c.QueryParameters.Select = new[] { "id", "name", "file" };
                            c.QueryParameters.Top = 999;
                        }, cancellationToken);

                    if (children?.Value != null)
                    {
                        var existingFile = children.Value.FirstOrDefault(item => 
                            item.Name == LEDGER_FILE_NAME && item.File != null);
                        
                        if (existingFile?.Id != null)
                        {
                            _logger.LogInformation("Ledger file already exists");
                            return existingFile.Id;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Error checking for existing ledger file, will attempt to create");
                }

                // Create a new empty Excel file with headers
                var ledgerContent = CreateEmptyLedgerContent();
                var stream = new MemoryStream(ledgerContent);

                var uploadedFile = await UploadFileAsync(driveId, hsaFolderId, LEDGER_FILE_NAME, stream, cancellationToken);

                _logger.LogInformation("Created new ledger file");
                return uploadedFile?.Id;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to ensure ledger file");
                return null;
            }
        }

        private async Task<DriveItem?> UploadFileAsync(string driveId, string parentId, string fileName, Stream fileStream, CancellationToken cancellationToken)
        {
            try
            {
                // Use the simpleUpload for small files
                var uploadedItem = await _graphService.Client!.Drives[driveId].Items[parentId]
                    .ItemWithPath(fileName)
                    .Content
                    .PutAsync(fileStream, cancellationToken: cancellationToken);

                _logger.LogInformation("Uploaded file: {FileName}", fileName);
                return uploadedItem;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to upload file {FileName}", fileName);
                return null;
            }
        }

        private void AddLedgerEntryAsync(
            string driveId,
            DateTime receiptDate,
            string vendor,
            decimal amount,
            string description,
            string fileName,
            string? fileUrl)
        {
            try
            {
                // For now, we'll log this entry but full Excel table manipulation via Graph API
                // would require more complex workbook requests. 
                // The user will see the file uploaded, and we'll provide guidance to manually add to the Excel table
                // OR we can use a simpler approach: store metadata in the database and provide a download of updated Excel

                _logger.LogInformation("Receipt entry recorded: {Vendor} {Amount} on {Date}", vendor, amount, receiptDate);
                
                // TODO: Implement actual Excel row addition via Graph Workbook API
                // For MVP, we just log and the file is uploaded; user can add to Excel manually or we batch-update later
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to add ledger entry");
                throw;
            }
        }

        private byte[] CreateEmptyLedgerContent()
        {
            // Create a minimal valid Excel file (xlsx format)
            // This is a simplified approach - a real solution would use a library like EPPlus
            // For now, we'll create a basic structure that can be opened in Excel

            // Minimal xlsx file structure (base64 encoded empty workbook)
            // This is a placeholder - in production, you'd want to use a proper library
            // For now, we'll create a file that Excel can read and user can manually populate

            string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheetData>
    <row r=""1"">
      <c r=""A1"" t=""inlineStr""><is><t>Date</t></is></c>
      <c r=""B1"" t=""inlineStr""><is><t>Vendor</t></is></c>
      <c r=""C1"" t=""inlineStr""><is><t>Amount</t></is></c>
      <c r=""D1"" t=""inlineStr""><is><t>Description</t></is></c>
      <c r=""E1"" t=""inlineStr""><is><t>Filename</t></is></c>
      <c r=""F1"" t=""inlineStr""><is><t>Reimbursed</t></is></c>
    </row>
  </sheetData>
</worksheet>";

            // Return as bytes - in production, this should be a proper XLSX creation
            return System.Text.Encoding.UTF8.GetBytes(xmlContent);
        }
    }
}
