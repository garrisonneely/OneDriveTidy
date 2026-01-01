# OneDriveTidy

OneDriveTidy is a C# utility designed to help manage and organize large OneDrive repositories (approx. 1TB). It aims to automate the cleanup process, organize media files, and extract valuable metadata from content.

## Features

### 1. Duplicate Removal
- Scans the entire OneDrive instance to identify duplicate files.
- Utilizes a local database mirror of file names and folder structures to efficiently track and compare files.
- Provides functionality to remove identified duplicates.

### 2. Artifact Filtering & Cleanup
- Identifies files that may not be "keepworthy" (e.g., temporary files, system files, or non-media artifacts).
- Provides an interface or mechanism to review these files.
- Allows for selection and deletion of these files, either individually or in bulk.

### 3. Media Organization
- Scans for all photo and video files.
- Extracts metadata (EXIF, creation dates, etc.) from media files.
- Automatically moves and organizes media into a structured folder hierarchy based on the extracted metadata (e.g., Year/Month/Day or Event-based).

### 4. Video Content Analysis
- Scans video files to extract content metadata.
- Generates transcripts from video audio to better understand content.
- Stores extracted details and transcripts in a structured document format for easy searching and reference.

## Architecture & Plan of Attack

### Technology Stack
- **Framework:** .NET 8
- **UI:** Blazor Web App (Local web interface)
- **Cloud Interaction:** Microsoft Graph SDK (Server-side operations to avoid sync issues)
- **Local Cache:** LiteDB (NoSQL single-file DB for flexible metadata storage)
- **Metadata:** MetadataExtractor (EXIF/XMP parsing)
- **AI/Transcription:** Whisper.net (Local, GPU-accelerated transcription via `whisper.cpp`)

### Architecture: "The Hybrid Scanner"
To handle 1TB of data efficiently, especially with "Files On-Demand":
1.  **Indexer (Graph API + LiteDB):**
    - Uses Graph API **Delta Queries** to crawl the drive.
    - Caches `DriveItemId`, `Name`, `Path`, and `ContentHash` in LiteDB.
    - Avoids full downloads for simple listing/deduplication.
2.  **Duplicate Manager:**
    - Queries LiteDB for matching `ContentHash` values.
    - Executes deletions via Graph API.
3.  **Media Processor (Stream-Process-Discard):**
    - **Photos:** Downloads partial streams (headers) where possible to read EXIF.
    - **Videos:** Downloads to temp storage -> Extracts Audio -> Transcribes with Whisper -> Deletes temp file.
    - **Activation:** This feature is user-activated, not automatic, to manage bandwidth.
4.  **Organizer:**
    - Moves files using Graph API `DriveItem.Move` commands (instant server-side move).

### Implementation Steps
1.  **Solution Setup:** Initialize `OneDriveTidy.Core` (Logic) and `OneDriveTidy.App` (Blazor Web App).
2.  **Auth & Indexer:** Implement Graph Auth and the Delta Query scanner.
3.  **UI Dashboard:** Create the main view to login and visualize scan progress.
4.  **Duplicate Logic:** Implement hash comparison and cleanup UI.
5.  **Media Processor:** Build the download/transcribe pipeline.
6.  **Organizer UI:** Build the view to review and execute file moves.
