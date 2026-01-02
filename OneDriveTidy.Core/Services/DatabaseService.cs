using LiteDB;
using Microsoft.Extensions.Logging;
using OneDriveTidy.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OneDriveTidy.Core.Services
{
    public class DatabaseService : IDisposable
    {
        private readonly LiteDatabase _db;
        private readonly ILogger<DatabaseService> _logger;
        private const string CollectionName = "driveItems";
        private const string ConfigCollectionName = "config";

        private bool _isDisposed = false;
        private readonly object _lock = new object();

        public DatabaseService(string dbPath, ILogger<DatabaseService> logger)
        {
            _logger = logger;
            // Use default connection (Direct) but manage concurrency with a lock
            // This avoids the "EnterTransaction" error and the "Mutex" crash on dispose
            _db = new LiteDatabase(dbPath);
            Initialize();
        }

        private void Initialize()
        {
            lock (_lock)
            {
                if (_isDisposed) return;
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                col.EnsureIndex(x => x.Id);
                col.EnsureIndex(x => x.ContentHash);
                col.EnsureIndex(x => x.ParentId);
                col.EnsureIndex(x => x.Size); // Index for sorting by size
            }
        }

        public void SaveDeltaLink(string deltaLink)
        {
            lock (_lock)
            {
                if (_isDisposed) return;
                var col = _db.GetCollection<BsonDocument>(ConfigCollectionName);
                col.Upsert("deltaLink", new BsonDocument { ["value"] = deltaLink });
            }
        }

        public string? GetDeltaLink()
        {
            lock (_lock)
            {
                if (_isDisposed) return null;
                var col = _db.GetCollection<BsonDocument>(ConfigCollectionName);
                var doc = col.FindById("deltaLink");
                return doc?["value"].AsString;
            }
        }

        public void UpsertItem(DriveItemModel item)
        {
            lock (_lock)
            {
                if (_isDisposed) return;
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                col.Upsert(item);
            }
        }

        public void UpsertItems(IEnumerable<DriveItemModel> items)
        {
            lock (_lock)
            {
                if (_isDisposed) return;
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                col.Upsert(items);
            }
        }

        public void DeleteItem(string id)
        {
            lock (_lock)
            {
                if (_isDisposed) return;
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                col.Delete(id);
            }
        }

        public void DeleteItems(IEnumerable<string> ids)
        {
            lock (_lock)
            {
                if (_isDisposed) return;
                try 
                {
                    var col = _db.GetCollection<DriveItemModel>(CollectionName);
                    var idValues = ids.Select(id => new BsonValue(id)).ToList();
                    if (idValues.Any())
                    {
                        col.DeleteMany("$._id IN @0", new BsonArray(idValues));
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "LiteDB DeleteItems Error");
                }
            }
        }

        public DriveItemModel? GetItem(string id)
        {
            lock (_lock)
            {
                if (_isDisposed) return null;
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                return col.FindById(id);
            }
        }

        public IEnumerable<DriveItemModel> GetAllItems()
        {
            lock (_lock)
            {
                if (_isDisposed) return Enumerable.Empty<DriveItemModel>();
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                return col.FindAll().ToList();
            }
        }

        public IEnumerable<IGrouping<string?, DriveItemModel>> GetDuplicates()
        {
            lock (_lock)
            {
                if (_isDisposed) return Enumerable.Empty<IGrouping<string?, DriveItemModel>>();
                
                _logger.LogInformation("Starting GetDuplicates query...");
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                
                // Optimized: Use SQL to find hashes with > 1 count first, then fetch items
                // This avoids loading ALL items into memory
                try 
                {
                    // 1. Find duplicate hashes
                    var duplicateHashes = new List<string>();
                    using (var reader = _db.Execute(@"
                        SELECT ContentHash 
                        FROM driveItems 
                        WHERE IsFolder = false AND ContentHash != null 
                        GROUP BY ContentHash 
                        HAVING COUNT(*) > 1
                    "))
                    {
                        while(reader.Read())
                        {
                            duplicateHashes.Add(reader.Current["ContentHash"].AsString);
                        }
                    }

                    _logger.LogInformation("Found {Count} duplicate groups.", duplicateHashes.Count);

                    if (!duplicateHashes.Any()) return Enumerable.Empty<IGrouping<string?, DriveItemModel>>();

                    // 2. Fetch items for these hashes
                    // We can't use IN clause with a huge list, so we might need to iterate or chunk
                    // Or just fetch all non-folders with hash and filter in memory (still better than grouping all)
                    // But actually, fetching all items with those hashes is what we want.
                    
                    // Let's try a direct query for items where hash is in our list
                    // LiteDB "IN" clause limit?
                    
                    var allDuplicateItems = new List<DriveItemModel>();
                    
                    // Chunking to avoid query limits
                    int chunkSize = 100;
                    for (int i = 0; i < duplicateHashes.Count; i += chunkSize)
                    {
                        var chunk = duplicateHashes.Skip(i).Take(chunkSize).Select(h => new BsonValue(h)).ToList();
                        var items = col.Find(Query.In("ContentHash", new BsonArray(chunk)));
                        allDuplicateItems.AddRange(items);
                    }

                    return allDuplicateItems.GroupBy(x => x.ContentHash).ToList();
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error in GetDuplicates optimization, falling back to slow method.");
                    // Fallback
                    var files = col.Find(x => !x.IsFolder && x.ContentHash != null);
                    return files.GroupBy(x => x.ContentHash)
                                .Where(g => g.Count() > 1)
                                .ToList();
                }
            }
        }

        public long GetItemCount()
        {
            lock (_lock)
            {
                if (_isDisposed) return 0;
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                return col.Count();
            }
        }

        public long GetTotalSize()
        {
            lock (_lock)
            {
                if (_isDisposed) return 0;
                try 
                {
                    using var result = _db.Execute("SELECT SUM(Size) FROM driveItems WHERE IsFolder = false");
                    if (result.Read() && !result.Current.IsNull)
                    {
                        // result.Current is a document like { "SUM(Size)": 12345 }
                        // We grab the first value
                        var val = result.Current.AsDocument.Values.FirstOrDefault();
                        return val.IsNumber ? val.AsInt64 : 0;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Error calculating total size via SQL, falling back.");
                    try 
                    {
                        var col = _db.GetCollection<DriveItemModel>(CollectionName);
                        return col.Find(x => !x.IsFolder).Sum(x => x.Size ?? 0);
                    }
                    catch {}
                }
                return 0;
            }
        }

        public (int GroupCount, long WastedSize) GetDuplicateStats()
        {
            lock (_lock)
            {
                if (_isDisposed) return (0, 0);
                
                _logger.LogInformation("Calculating duplicate stats...");
                try 
                {
                    // Use MAX(Size) instead of FIRST(Size)
                    // Ensure we dispose the reader to release locks!
                    using var result = _db.Execute(@"
                        SELECT COUNT(*) AS Cnt, MAX(Size) AS Sz
                        FROM driveItems
                        WHERE IsFolder = false AND ContentHash != null
                        GROUP BY ContentHash
                        HAVING COUNT(*) > 1
                    ");

                    int groupCount = 0;
                    long wastedSize = 0;

                    while(result.Read())
                    {
                        var row = result.Current;
                        groupCount++;
                        var count = row["Cnt"].AsInt32;
                        var size = row["Sz"].AsInt64;
                        wastedSize += (count - 1) * size;
                    }
                    
                    _logger.LogInformation("Stats calculated: {GroupCount} groups, {WastedSize} bytes wasted.", groupCount, wastedSize);
                    return (groupCount, wastedSize);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error calculating duplicate stats via SQL.");
                    return (0, 0);
                }
            }
        }
        
        public void ClearAll()
        {
             lock (_lock)
             {
                 if (_isDisposed) return;
                 var col = _db.GetCollection<DriveItemModel>(CollectionName);
                 col.DeleteAll();
             }
        }

        public void Dispose()
        {
            lock (_lock)
            {
                if (_isDisposed) return;
                _logger.LogInformation("DatabaseService Disposing...");
                _isDisposed = true;
                _db?.Dispose();
                _logger.LogInformation("DatabaseService Disposed.");
            }
        }
    }
}
