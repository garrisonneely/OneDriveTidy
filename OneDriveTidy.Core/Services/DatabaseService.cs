using LiteDB;
using OneDriveTidy.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OneDriveTidy.Core.Services
{
    public class DatabaseService : IDisposable
    {
        private readonly LiteDatabase _db;
        private const string CollectionName = "driveItems";

        private const string ConfigCollectionName = "config";

        private bool _isDisposed = false;

        public DatabaseService(string dbPath)
        {
            // Keep a single instance open for thread safety and to avoid file lock issues
            _db = new LiteDatabase(dbPath);
            Initialize();
        }

        private void Initialize()
        {
            if (_isDisposed) return;
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            col.EnsureIndex(x => x.Id);
            col.EnsureIndex(x => x.ContentHash);
            col.EnsureIndex(x => x.ParentId);
            col.EnsureIndex(x => x.Size); // Index for sorting by size
        }

        public void SaveDeltaLink(string deltaLink)
        {
            if (_isDisposed) return;
            var col = _db.GetCollection<BsonDocument>(ConfigCollectionName);
            col.Upsert("deltaLink", new BsonDocument { ["value"] = deltaLink });
        }

        public string? GetDeltaLink()
        {
            if (_isDisposed) return null;
            var col = _db.GetCollection<BsonDocument>(ConfigCollectionName);
            var doc = col.FindById("deltaLink");
            return doc?["value"].AsString;
        }

        public void UpsertItem(DriveItemModel item)
        {
            if (_isDisposed) return;
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            col.Upsert(item);
        }

        public void UpsertItems(IEnumerable<DriveItemModel> items)
        {
            if (_isDisposed) return;
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            col.Upsert(items);
        }

        public void DeleteItem(string id)
        {
            if (_isDisposed) return;
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            col.Delete(id);
        }

        public void DeleteItems(IEnumerable<string> ids)
        {
            if (_isDisposed) return;
            try 
            {
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                // Use DeleteMany with BsonExpression for atomicity and performance
                // "$._id IN ['id1', 'id2', ...]"
                var idValues = ids.Select(id => new BsonValue(id)).ToList();
                if (idValues.Any())
                {
                    col.DeleteMany("$._id IN @0", new BsonArray(idValues));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LiteDB DeleteItems Error: {ex.Message}");
            }
        }

        public DriveItemModel? GetItem(string id)
        {
            if (_isDisposed) return null;
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            return col.FindById(id);
        }

        public IEnumerable<DriveItemModel> GetAllItems()
        {
            if (_isDisposed) return Enumerable.Empty<DriveItemModel>();
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            return col.FindAll().ToList();
        }

        public IEnumerable<IGrouping<string?, DriveItemModel>> GetDuplicates()
        {
            if (_isDisposed) return Enumerable.Empty<IGrouping<string?, DriveItemModel>>();
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            
            // Find items that are not folders and have a hash
            var files = col.Find(x => !x.IsFolder && x.ContentHash != null);

            // Group by hash and return only groups with more than 1 item
            return files.GroupBy(x => x.ContentHash)
                        .Where(g => g.Count() > 1)
                        .ToList();
        }

        public long GetItemCount()
        {
            if (_isDisposed) return 0;
            var col = _db.GetCollection<DriveItemModel>(CollectionName);
            return col.Count();
        }

        public long GetTotalSize()
        {
            if (_isDisposed) return 0;
            // Use SQL for efficiency if possible, or fallback to LINQ
            // LiteDB v5 supports SQL
            try 
            {
                // Only sum files, not folders, to avoid double counting
                var result = _db.Execute("SELECT SUM(Size) FROM driveItems WHERE IsFolder = false");
                if (result.Read() && !result.Current.IsNull)
                {
                    return result.Current.AsInt64;
                }
            }
            catch 
            {
                // Fallback
                var col = _db.GetCollection<DriveItemModel>(CollectionName);
                return col.Find(x => !x.IsFolder).Sum(x => x.Size ?? 0);
            }
            return 0;
        }

        public (int GroupCount, long WastedSize) GetDuplicateStats()
        {
            if (_isDisposed) return (0, 0);
            var duplicates = GetDuplicates();
            int count = duplicates.Count();
            long size = 0;
            foreach(var group in duplicates)
            {
                // Wasted size is (Count - 1) * Size of one item
                // Assuming all items in group have same size (which they should if hash matches)
                var itemSize = group.First().Size ?? 0;
                size += (group.Count() - 1) * itemSize;
            }
            return (count, size);
        }
        
        public void ClearAll()
        {
             if (_isDisposed) return;
             var col = _db.GetCollection<DriveItemModel>(CollectionName);
             col.DeleteAll();
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            Console.WriteLine("DatabaseService Disposing...");
            _isDisposed = true;
            _db?.Dispose();
            Console.WriteLine("DatabaseService Disposed.");
        }
    }
}
