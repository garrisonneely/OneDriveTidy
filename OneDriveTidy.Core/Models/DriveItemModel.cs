using System;

namespace OneDriveTidy.Core.Models
{
    public class DriveItemModel
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? ParentId { get; set; }
        public string? Path { get; set; }
        public string? ContentHash { get; set; }
        public long? Size { get; set; }
        public DateTimeOffset? CreatedDateTime { get; set; }
        public DateTimeOffset? LastModifiedDateTime { get; set; }
        public bool IsFolder { get; set; }
        public string? WebUrl { get; set; }
        
        // Metadata for later
        public DateTime? PhotoTakenDate { get; set; }
        public string? CameraModel { get; set; }
        public bool IsTranscribed { get; set; }
        public string? Transcript { get; set; }
    }
}
