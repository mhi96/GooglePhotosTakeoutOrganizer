# Google Photos Takeout Organizer
Organize photos and videos into Year/Month/Type folders. Works best with Google Takeout .json metadata, supports multiple formats, handles missing metadata, avoids overwrites, and logs all moves in Excel.

Organize your photos and videos into a structured folder hierarchy automatically.

## Requirements

- Python 3.8+  
- Required packages:
  ```bash
  pip install pillow pymediainfo openpyxl



## Usage
- Place your photos/videos in a source folder.
- Update the source_folder and destination_folder paths in organizer.py.
- Run the script:
  ```
  python organizer.py
  ```
- Check the destination folder for your organized media.
- Check file_moves.xlsx in the destination folder for a detailed log.

## Features
- Organizes files into folders by **Year → Month → Media Type (Pictures/Videos)**.
- Reads **Google Takeout `.json` metadata** to determine the exact date a photo/video was taken.
- Supports multiple image formats: `.jpg`, `.jpeg`, `.png`, `.heic`, `.webp`.
- Supports multiple video formats: `.mp4`, `.mov`, `.3gp`, `.mkv`, `.ts`.
- Handles files with missing metadata by placing them into `No-Date-Taken` folders.
- Avoids file overwriting by appending `_1`, `_2`, etc., if duplicate filenames exist.
- Logs all moves to an **Excel file** with details:
  - Original Path  
  - New Path  
  - Method Used (`JSON`, `EXIF`, `MediaInfo`, `Unknown`)  
  - File Name Without Extension  
  - Only Extension
- **Multithreaded processing** for faster performance on large collections.
- Fully compatible with Google Takeout exports, but works for **any photo/video collection**.



## Notes
- Works best with Google Takeout exports because JSON metadata provides exact timestamps.
- If metadata is missing, the script will fall back on EXIF data for photos or MediaInfo for videos.
- Duplicate filenames are automatically renamed with a numeric suffix to prevent overwriting.
