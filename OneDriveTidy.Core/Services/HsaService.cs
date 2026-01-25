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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Kiota.Abstractions.Serialization;
using Microsoft.Kiota.Abstractions;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

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
                else
                {
                    // Ensure schema is up to date (add Reimbursed? column if missing)
                    await EnsureReimbursedColumnAsync(drive.Id, _ledgerFileId, cancellationToken);
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
        /// Retrieves all rows from the master ledger.
        /// </summary>
        public async Task<List<LedgerEntry>> GetLedgerEntriesAsync(CancellationToken cancellationToken = default)
        {
            if (_graphService.Client == null) return new List<LedgerEntry>();

            try 
            {
                // Ensure IDs are available
                if (_ledgerFileId == null || _hsaFolderId == null)
                {
                    // This fetches IDs without creating if possible, or checks structure
                     var drive = await _graphService.Client.Me.Drive.GetAsync(cancellationToken: cancellationToken);
                     if (drive?.Id != null)
                     {
                         // Try to find the HSA folder if we don't have it
                         if (_hsaFolderId == null)
                         {
                            // This logic is duplicated from EnsureFolder somewhat but we want "Find" semantics
                            // Relying on EnsureStructure being called at least once is safer, but let's try to be robust
                            var hsaFolder = await GetFolderAsync(drive.Id, "root", HSA_FOLDER_NAME, cancellationToken);
                            _hsaFolderId = hsaFolder?.Id;
                         }

                         if (_hsaFolderId != null && _ledgerFileId == null)
                         {
                             // Try to find ledger
                             var ledger = await GetFileAsync(drive.Id, _hsaFolderId, LEDGER_FILE_NAME, cancellationToken);
                             _ledgerFileId = ledger?.Id;
                         }
                     }
                }

                if (_ledgerFileId == null) return new List<LedgerEntry>();

                var drive2 = await _graphService.Client.Me.Drive.GetAsync(cancellationToken: cancellationToken);
                if (drive2?.Id == null) return new List<LedgerEntry>();

                var response = await _graphService.Client.Drives[drive2.Id].Items[_ledgerFileId].Workbook.Tables["Receipts"].Rows.GetAsync(c => {
                    // We only need the values
                    c.QueryParameters.Select = new[] { "values" };
                }, cancellationToken);

                var entries = new List<LedgerEntry>();
                if (response?.Value != null)
                {
                    foreach (var row in response.Value)
                    {
                        if (row.Values is UntypedArray arr)
                        {
                            var cells = arr.GetValue().ToList();
                            // Expected columns: Date, Vendor, Amount, Description, Filename, Link
                            if (cells.Count >= 3)
                            {
                                var dateStr = (cells[0] as UntypedString)?.GetValue();
                                var vendor = (cells[1] as UntypedString)?.GetValue();
                                var amountVal = cells[2];
                                
                                decimal amount = 0;
                                if (amountVal is UntypedDouble d) amount = (decimal)d.GetValue();
                                else if (amountVal is UntypedInteger i) amount = (decimal)i.GetValue();
                                else if (amountVal is UntypedString s && decimal.TryParse(s.GetValue(), out var parsed)) amount = parsed;

                                if (dateStr != null && vendor != null)
                                {
                                    if (DateTime.TryParse(dateStr, out var date))
                                    {
                                         entries.Add(new LedgerEntry(date, vendor, amount));
                                    }
                                }
                            }
                        }
                    }
                }
                return entries;

            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to retrieve ledger entries.");
                return new List<LedgerEntry>();
            }
        }

        private async Task<DriveItem?> GetFolderAsync(string driveId, string parentId, string name, CancellationToken token)
        {
             try {
                var children = await _graphService.Client!.Drives[driveId].Items[parentId].Children.GetAsync(c => {
                    c.QueryParameters.Filter = $"name eq '{name}' and folder ne null";
                }, token);
                return children?.Value?.FirstOrDefault();
             } catch { return null; }
        }

        private async Task<DriveItem?> GetFileAsync(string driveId, string parentId, string name, CancellationToken token)
        {
             try {
                var children = await _graphService.Client!.Drives[driveId].Items[parentId].Children.GetAsync(c => {
                    c.QueryParameters.Filter = $"name eq '{name}' and file ne null";
                }, token);
                return children?.Value?.FirstOrDefault();
             } catch { return null; }
        }

        public record LedgerEntry(DateTime Date, string Vendor, decimal Amount);

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
                await AddLedgerEntryAsync(drive.Id, receiptDate, vendor, amount, description, standardizedName, uploadedFile.WebUrl);

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

        private async Task AddLedgerEntryAsync(
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
                if (_ledgerFileId == null)
                {
                    _logger.LogWarning("Ledger file ID is null, skipping ledger entry.");
                    StatusChanged?.Invoke("Warning: Ledger file not found, skipping entry.");
                    return;
                }

                _logger.LogInformation("Adding ledger entry: {Vendor} {Amount} on {Date}", vendor, amount, receiptDate);
                StatusChanged?.Invoke("Updating master ledger...");

                // Construct the row data matching the table columns: Date, Vendor, Amount, Description, Filename, Link, Reimbursed?
                // Graph API expects a JSON array of array of values
                var rowList = new List<UntypedNode>
                {
                    new UntypedString(receiptDate.ToString("yyyy-MM-dd")),
                    new UntypedString(vendor),
                    new UntypedDouble((double)amount),
                    new UntypedString(description ?? ""),
                    new UntypedString(fileName),
                    new UntypedString(fileUrl ?? ""),
                    new UntypedString("N")
                };

                var rowArray = new UntypedArray(rowList);
                var valuesArray = new UntypedArray(new List<UntypedNode> { rowArray });

                // Add row to the "Receipts" table
                // POST /drives/{drive-id}/items/{id}/workbook/tables/{name}/rows/add
                var requestBody = new Microsoft.Graph.Drives.Item.Items.Item.Workbook.Tables.Item.Rows.Add.AddPostRequestBody
                {
                    Values = valuesArray,
                    Index = null // Add to end
                };

                await _graphService.Client!.Drives[driveId].Items[_ledgerFileId].Workbook
                    .Tables["Receipts"]
                    .Rows
                    .Add
                    .PostAsync(requestBody);

                _logger.LogInformation("Ledger entry added successfully.");
                StatusChanged?.Invoke("Ledger updated successfully.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to add ledger entry");
                // Don't throw - we still uploaded the file, failing the whole operation here is annoying for the user
                StatusChanged?.Invoke($"Warning: Failed to update ledger ({ex.Message})");
            }
        }

        private byte[] CreateEmptyLedgerContent()
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    
                    // Create minimal SheetData with Header
                    var sheetData = new SheetData();
                    var headerRow = new Row() { RowIndex = 1 };
                    
                    string[] headers = { "Date", "Vendor", "Amount", "Description", "Filename", "Link", "Reimbursed?" };
                    foreach (var header in headers)
                    {
                        headerRow.Append(new Cell 
                        { 
                            CellValue = new CellValue(header), 
                            DataType = new EnumValue<CellValues>(CellValues.String) 
                        });
                    }
                    sheetData.Append(headerRow);
                    
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    // Create the Table Definition
                    var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
                    string tableRef = "A1:G2"; // Header + 1 empty row
                    
                    var table = new Table() 
                    { 
                        Id = 1, 
                        Name = "Receipts", 
                        DisplayName = "Receipts", 
                        Reference = tableRef, 
                        TotalsRowShown = false 
                    };

                    var autoFilter = new AutoFilter() { Reference = tableRef };
                    
                    var tableColumns = new TableColumns() { Count = 7 };
                    uint colId = 1;
                    foreach (var header in headers)
                    {
                        tableColumns.Append(new TableColumn() { Id = colId++, Name = header });
                    }

                    var tableStyleInfo = new TableStyleInfo() 
                    { 
                        Name = "TableStyleMedium9", 
                        ShowFirstColumn = false, 
                        ShowLastColumn = false, 
                        ShowRowStripes = true, 
                        ShowColumnStripes = false 
                    };

                    table.Append(autoFilter);
                    table.Append(tableColumns);
                    table.Append(tableStyleInfo);

                    tableDefinitionPart.Table = table;

                    // Link Table to Worksheet
                    var tableParts = new TableParts(new TablePart() { Id = worksheetPart.GetIdOfPart(tableDefinitionPart) });
                    worksheetPart.Worksheet.Append(tableParts);

                    // Add Sheets to Workbook
                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet() 
                    { 
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart), 
                        SheetId = 1, 
                        Name = "Sheet1" 
                    };
                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();
                }
                
                return memoryStream.ToArray();
            }
        }

        private async Task EnsureReimbursedColumnAsync(string driveId, string fileId, CancellationToken cancellationToken)
        {
            try
            {
                _logger.LogInformation("Checking for 'Reimbursed?' column...");
                
                // 1. Check if column exists
                var columns = await _graphService.Client!.Drives[driveId].Items[fileId].Workbook
                    .Tables["Receipts"].Columns
                    .GetAsync(c => c.QueryParameters.Select = new[] { "name" }, cancellationToken);
                
                bool exists = columns?.Value?.Any(c => c.Name == "Reimbursed?") ?? false;

                if (!exists)
                {
                    _logger.LogInformation("Column 'Reimbursed?' missing. Adding it...");
                    StatusChanged?.Invoke("Adding 'Reimbursed?' column...");

                    // 2. Add the column
                    var newCol = await _graphService.Client!.Drives[driveId].Items[fileId].Workbook
                        .Tables["Receipts"].Columns
                        .PostAsync(new WorkbookTableColumn
                        {
                            Name = "Reimbursed?",
                            Index = null // Append to end
                        }, cancellationToken: cancellationToken);

                    // 3. Default existing rows to "Y"
                    // We need to reference the DataBodyRange of this column.
                    // Note: If the table has no data rows, this might throw or return null range, so we safeguard.
                    try 
                    {
                        var dataBodyRange = await _graphService.Client.Drives[driveId].Items[fileId].Workbook
                            .Tables["Receipts"].Columns["Reimbursed?"].DataBodyRange
                            .GetAsync(cancellationToken: cancellationToken);

                        // If address or row count indicates we have data
                        if (dataBodyRange != null && dataBodyRange.RowCount > 0 && !string.IsNullOrEmpty(dataBodyRange.Address))
                        {
                            StatusChanged?.Invoke("Backfilling existing rows with 'Y'...");
                            
                            // Parse address to find sheet and range (e.g. Sheet1!G2:G5)
                            // We need to target the worksheet range directly to PATCH
                            string address = dataBodyRange.Address;
                            string[] parts = address.Split('!');
                            
                            if (parts.Length == 2)
                            {
                                string sheetName = parts[0].Replace("'", ""); 
                                string rangeRef = parts[1];

                                await PatchRangeAsync(driveId, fileId, sheetName, rangeRef, "Y", cancellationToken);
                                
                                _logger.LogInformation("Backfilled 'Reimbursed?' column with 'Y' for existing rows.");
                            }
                        }
                    }
                    catch (Exception valEx)
                    {
                        // If there were no rows, getting DataBodyRange might fail, or patching it might fail.
                        // We swallow this specific error as it likely means there was no data to backfill.
                        _logger.LogWarning(valEx, "Could not backfill 'Reimbursed?' column (table might be empty).");
                    }
                }
                else
                {
                    _logger.LogInformation("'Reimbursed?' column already exists.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to ensure 'Reimbursed?' column schema.");
                StatusChanged?.Invoke($"Warning: Schema update failed ({ex.Message})");
            }
        }

        private async Task PatchRangeAsync(string driveId, string fileId, string worksheet, string rangeAddress, string value, CancellationToken cancellationToken)
        {
             // Helper to PATCH a range using manual request construction to bypass SDK limitations
             var baseUrl = _graphService.Client!.RequestAdapter.BaseUrl;
             var url = $"{baseUrl}/drives/{driveId}/items/{fileId}/workbook/worksheets/{worksheet}/range(address='{rangeAddress}')";
             
             var requestInfo = new RequestInformation
             {
                 HttpMethod = Method.PATCH,
                 UrlTemplate = url
             };
             
             var body = new WorkbookRange
             {
                 Values = new UntypedArray(new List<UntypedNode> 
                 { 
                     new UntypedArray(new List<UntypedNode> { new UntypedString(value) }) 
                 })
             };

             requestInfo.SetContentFromParsable(_graphService.Client.RequestAdapter, "application/json", body);
             await _graphService.Client.RequestAdapter.SendPrimitiveAsync<Stream>(requestInfo, null, cancellationToken);
        }
    }
}
