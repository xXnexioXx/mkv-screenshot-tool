# MKV Screenshot Tool

A PowerShell 7 script that automatically captures 5 evenly spaced screenshots from an MKV video file, uploads them to [ImgBB](https://imgbb.com), and generates a BBCode-formatted `.txt` file ready to paste into forum posts.

---

## Features

- Automatically finds the MKV file in the script folder
- Screenshots taken between 10%–90% of the video, evenly distributed
- 3 frame candidates per position — the sharpest (largest file size) is kept
- Uploads to ImgBB with the video title and timestamp as the image name
- Generates a BBCode `.txt` file with 5 linked, resized images
- Handles filenames with spaces, special characters, and unicode
- Auto-restarts in PowerShell 7 if accidentally opened in PS5
- Cleans up temporary PNG files after upload

---

## Requirements

- [PowerShell 7](https://aka.ms/powershell)
- [ffmpeg + ffprobe](https://ffmpeg.org/download.html) — place the `.exe` files in the same folder as the script
- An [ImgBB account](https://imgbb.com) with an API key

---

## Setup

1. Download `mkv-screenshot-tool.ps1`
2. Place `ffmpeg.exe` and `ffprobe.exe` in the same folder
3. Open `mkv-screenshot-tool.ps1` and replace `YOUR_IMGBB_API_KEY` on line 26 with your actual API key
4. Run the script!

---

## Usage

Simply double-click or right-click the `.ps1` file and select **Run with PowerShell**.

### Folder structure

```
📁 your-folder/
├── mkv-screenshot-tool.ps1
├── ffmpeg.exe
├── ffprobe.exe
└── your-movie.mkv
```

### Output

After a successful run, the folder will contain:

```
📁 your-folder/
├── mkv-screenshot-tool.ps1
├── ffmpeg.exe
├── ffprobe.exe
├── your-movie.mkv
└── your-movie.txt        ← BBCode output, ready to paste
```

The `.txt` file will look like this:

```bbcode
[center]
[url=https://i.ibb.co/...][img=500]https://i.ibb.co/...[/img][/url] [url=https://i.ibb.co/...][img=500]https://i.ibb.co/...[/img][/url]

[url=https://i.ibb.co/...][img=500]https://i.ibb.co/...[/img][/url] [url=https://i.ibb.co/...][img=500]https://i.ibb.co/...[/img][/url]

[url=https://i.ibb.co/...][img=500]https://i.ibb.co/...[/img][/url]
[/center]
```

Images are uploaded to ImgBB with the name format: `Movie Title [HH-MM-SS]`

---

## How it works

1. Scans the folder for the first `.mkv` file
2. Reads the video duration via `ffprobe` (falls back to container header if stream data is unavailable)
3. Calculates 5 evenly spaced timestamps between 10% and 90% of the video length
4. For each timestamp, captures 3 candidate frames (1 second apart) and keeps the largest one to avoid black or blank frames
5. Uploads all 5 frames to ImgBB via the API
6. Writes a BBCode `.txt` file with the uploaded image URLs

---

## Notes

- Only one `.mkv` file should be present in the folder at a time
- Any existing `.txt` files in the folder are deleted at the start of each run
- If any upload fails, the `.txt` file will not be created to prevent incomplete output
- Images are uploaded with **no expiration** (unlimited)

---
