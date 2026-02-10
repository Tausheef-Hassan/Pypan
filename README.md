# PyPan - Wikimedia Commons Batch Uploader

**Version:** 0.2.1a0

## ⚠️ IMPORTANT SECURITY WARNING
This program **stores your Wikimedia Commons username and password** in temporary configuration files (`user-config.py` and `user-password.py`) so that Pywikibot can log in.  
Although the files are restricted to the current user and cleaned up after the program exits, **your password will be written in plain text** during runtime.  

**Recommendation:** Use [bot passwords](https://commons.wikimedia.org/wiki/Special:BotPasswords) if you have 2FA enabled for enhanced security.

**Do not use this tool on shared or insecure machines. Do not share the generated config files. Always log out and remove files after use.**

---

## Overview
PyPan is a batch uploader for Wikimedia Commons built on top of Pywikibot. It allows you to select an input file (Excel, CSV, or JSON) containing file paths and metadata, then automatically upload them with progress tracking, retry logic, verification, and comprehensive logging.

---

## Key Features

### Upload Capabilities
- **Multiple Input Formats**: Supports Excel (.xlsx, .xls), CSV (.csv), and JSON (.json) files
- **URL Downloads**: Upload files directly from URLs (including Wayback Machine fallback)
- **YouTube Support**: Download and upload YouTube videos (requires yt-dlp)
- **Video Conversion**: Automatically converts common video formats (MP4, AVI, MOV, etc.) to WebM
- **File Verification**: Validates uploads by comparing file sizes and wikitext content
- **Auto-increment Filenames**: Automatically handles duplicate filenames on Commons

### Reliability Features
- **Smart Retries**: Retries failed uploads up to 10 times (configurable)
- **Internet Resilience**: Waits and auto-retries if internet connection drops
- **Incremental Results**: Saves results after each successful upload
- **Pause/Resume**: Pause uploads and resume later
- **Stop Anytime**: Cleanly stop the upload process

### User Interface
- **Login Management**: Secure login with password visibility toggle
- **Progress Tracking**: Real-time progress bar with ETA
- **Live Statistics**: Success/failure counts and timing information
- **Internet Status**: Live internet connection indicator
- **Detailed Logging**: Timestamped logs of all operations
- **Reset Function**: Clear all settings and start fresh

### Configuration Options
- **Family & Language**: Configure wiki family and language (defaults to commons/commons)
- **Parallelization**: Control concurrent uploads (default: 1)
- **Retry Settings**: Configure max attempts and pause between retries
- **Upload Pause**: Set pause duration after successful uploads
- **Warning Handling**: Choose to ignore or respect upload warnings

---

## Differences from Pattypan
- Retries errors up to 10 times (configurable) where Pattypan might get stuck
- Waits and retries if internet connection is down, auto-resumes when connection restored
- Writes results back to output file with detailed status and verification columns
- Supports URL downloads including YouTube videos
- Automatic video format conversion to WebM
- Pause/resume functionality
- Provides accurate ETA and progress tracking
- Supports multiple input file formats (Excel, CSV, JSON)
- File verification after upload
- Incremental result saving

⚠️ **Note:** This tool focuses on batch uploading. For creating structured Excel files with metadata, consider using Pattypan or custom scripts.

---

## Supported File Types

### Images
`.png`, `.jpg`, `.jpeg`, `.gif`, `.tiff`, `.tif`, `.webp`, `.svg`, `.xcf`, `.apng`

### Audio
`.oga`, `.ogg`, `.mid`, `.midi`, `.wav`, `.flac`, `.mp3`, `.opus`

### Video
`.webm`, `.ogv`, `.ogx`

### Documents
`.pdf`, `.djvu`

### 3D Models
`.stl`

### Video Formats Auto-Converted to WebM
`.mp4`, `.avi`, `.mov`, `.mkv`, `.flv`, `.wmv`, `.m4v`, `.mpeg`, `.mpg`, `.3gp`, `.m2v`

---

## Input File Formats

### Excel (.xlsx, .xls)
The uploader expects an **Excel file with no header row**.  
Each row must have exactly three columns:

1. **File Path / URL** – Local path or URL to the file
   - Local: `C:\Users\Me\Pictures\photo.jpg`
   - URL: `https://example.com/image.jpg`
   - YouTube: `https://www.youtube.com/watch?v=VIDEO_ID`

2. **Target Filename** – Desired name on Wikimedia Commons
   - Example: `My_uploaded_photo.jpg`
   - Extension is auto-detected if not provided
   - Illegal characters are automatically sanitized
   - Duplicates are auto-incremented (e.g., `File (1).jpg`)

3. **Description/Wikitext** – Full description text for Commons page
   - Can include categories, templates, any valid wikitext
   - Program appends: `[[Category: Uploaded with pypan]]`

### CSV Format
Same three-column structure as Excel, comma-separated values.

### JSON Format
Array of objects with keys:
```json
[
  {
    "file_path": "C:\\path\\to\\file.jpg",
    "target_filename": "Example_file.jpg",
    "description": "{{Information\n|description=...\n}}"
  }
]
```

---

## Example Excel Table

⚠️ **Important:** There should be no header row. Data starts from row 1.

| File Path | Target Filename | Wikitext |
|-----------|-----------------|----------|
| `D:\Photos\photo1.jpg` | Syeda Rizwana Hasan Inspection (PID-0002220) | =={{int:filedesc}}==<br>{{Information<br>\|description = {{bn\|1=বর্ণনা}}{{en\|1=Description}}<br>\|date = 2025-04-24<br>\|source = {{Source-PID}}<br>\|author = Author<br>}}<br><br>=={{int:license-header}}==<br>{{PD-BDGov-PID}} |
| `https://example.com/image.jpg` | Example Image | File from URL<br>[[Category:Test]] |
| `https://www.youtube.com/watch?v=dQw4w9WgXcQ` | Example Video | YouTube video<br>[[Category:Videos]] |

### Excel Formula Example
```excel
="=={{int:filedesc}}==
{{Information
 |description = {{bn|1="&D1&"}}{{en|1="&C1&"}}
 |date = {{Date-PID|"&E1&"}}
 |source = {{Source-PID | url="&F1&"}}
 |author = {{Institution:Press Information Department}}
 |permission = 
 |other versions = 
}}

=={{int:license-header}}==
{{PD-BDGov-PID}}

[[Category:PID_BD]]"
```

⚠️ **Excel Formula Note:** When saving Excel files with formulas in the description column, ensure formulas are not evaluated. The program reads raw values, and formulas starting with `=` in the description column (column 3) are automatically escaped to prevent Excel from treating them as formulas.

---

## Installation & Dependencies

### Required
- Python 3.7+
- pywikibot
- pandas
- requests
- Pillow (PIL)
- openpyxl (for Excel support)

### Optional (for enhanced features)
- **yt-dlp** – For YouTube video downloads
  ```bash
  pip install yt-dlp
  ```
- **moviepy** – For video format conversion
  ```bash
  pip install moviepy
  ```

### Install All Dependencies
```bash
pip install pywikibot pandas requests Pillow openpyxl yt-dlp moviepy
```

---

## How It Works

### 1. Login
- Click "Login" button
- Enter Wikimedia Commons username and password
- Use bot password if you have 2FA enabled
- Program creates temporary config files
- Credentials are verified before proceeding

### 2. Configure Upload
- Select input file (Excel, CSV, or JSON)
- Configure settings:
  - **Family/Lang**: Usually `commons/commons`
  - **Parallelization**: Number of concurrent uploads
  - **Max Retry Attempts**: How many times to retry failed uploads
  - **Pause Between Retries**: Wait time before retrying
  - **Pause After Upload**: Brief pause after successful upload
  - **Ignore Warnings**: Whether to bypass upload warnings

### 3. Upload Process
- Click "Start Upload"
- For each file:
  - Downloads from URL if needed (with Wayback fallback)
  - Downloads YouTube videos if yt-dlp is available
  - Converts video formats to WebM if needed
  - Detects file type from content (not extension)
  - Sanitizes filename (removes illegal characters)
  - Checks for duplicates (auto-increments if exists)
  - Uploads to Commons
  - Verifies upload (compares size and wikitext)
  - Saves result incrementally
  - Waits configured pause duration

### 4. Results
- Output file created with same format as input
- Two new columns added:
  - `Upload_Status`: Success / Skipped / Failed with reason
  - `Verification`: Upload verification result
- Detailed logs available in application window

### 5. Cleanup
- Click "Logout" or close application
- Config files automatically deleted
- Pywikibot cache cleaned up

---

## Output File

Results are saved to a file with `_results` suffix in the same format as input:
- Excel input → Excel output with status columns
- CSV input → CSV output with status columns  
- JSON input → JSON output with status fields

### Status Values
- **Success** – Uploaded and verified successfully
- **Skipped: reason** – File skipped (already exists, invalid type, etc.)
- **Failed: reason** – Upload failed with error message

### Verification Values
- **Verified** – File size and wikitext match
- **Not OK: reason** – Verification failed with details

---

## Special Features

### URL Downloads
- Supports direct HTTP/HTTPS URLs
- Automatic fallback to Wayback Machine if download fails
- Retries with configured attempts

### YouTube Downloads
**Requires yt-dlp installation**

Strategy progression:
1. Anonymous download (works for public videos)
2. Chrome browser cookies (for restricted videos)
3. Firefox browser cookies
4. Edge browser cookies

If all strategies fail with bot detection, log into YouTube in your browser and retry.

### Video Conversion
**Requires moviepy installation**

Automatically converts these formats to WebM:
- MP4, AVI, MOV, MKV, FLV, WMV, M4V, MPEG, MPG, 3GP, M2V

Uses VP9 codec with Opus/Vorbis audio for Wikimedia Commons compatibility.

### Filename Sanitization
Automatically removes illegal characters:
- `: # < > [ ] | { } / \`
- Multiple tildes (`~~~`)
- Control characters
- Leading/trailing hyphens and spaces

### Auto-increment Duplicates
If `File.jpg` exists on Commons:
- First upload → `File (1).jpg`
- Second upload → `File (2).jpg`
- And so on...

---

## Configuration File Storage

### Compiled Executable
Config files stored in: `%USERPROFILE%\.pypan` (Windows)

### Script Mode
Config files stored in same directory as script

### Files Created
- `user-config.py` – Pywikibot configuration
- `user-password.py` – Login credentials (plain text)
- `apicache/` – Pywikibot API cache
- `throttle.ctrl` – Upload throttle control
- `pywikibot-USERNAME.lwp` – Login session

All files are automatically cleaned up on logout or exit.

---

## Troubleshooting

### "yt-dlp not installed"
```bash
pip install yt-dlp
```

### "moviepy not installed"
```bash
pip install moviepy
```

### YouTube download fails with bot detection
1. Open your web browser
2. Log into YouTube
3. Ensure you're logged in and can watch the video
4. Retry in PyPan (will use browser cookies)

### Upload verification fails
- Check internet connection stability
- Verify file wasn't corrupted during upload
- Check Commons page manually to confirm upload

### Excel formulas evaluated incorrectly
- Save formulas as text by prefixing with single quote: `'=formula`
- Or use CSV format instead of Excel

---

## Security Best Practices

1. **Use Bot Passwords**: Create bot passwords at https://commons.wikimedia.org/wiki/Special:BotPasswords
2. **Don't Share Configs**: Never share `user-password.py` or config directory
3. **Logout After Use**: Always click "Logout" when done
4. **Secure Machine**: Don't use on shared or public computers
5. **Monitor Activity**: Check your Commons contributions regularly

---

## Known Limitations

- Cannot create structured Excel input files (use Pattypan or custom scripts)
- Password stored in plain text during runtime
- Parallel uploads may hit rate limits (use conservative parallelization)
- Some YouTube videos may require manual download if heavily restricted
- Video conversion requires significant disk space and CPU

---

## Credits

Built with:
- [Pywikibot](https://www.mediawiki.org/wiki/Manual:Pywikibot) – MediaWiki bot framework
- [yt-dlp](https://github.com/yt-dlp/yt-dlp) – YouTube downloader
- [moviepy](https://zulko.github.io/moviepy/) – Video editing library

---

## License

This tool is provided as-is for Wikimedia Commons batch uploads. Use responsibly and in accordance with Wikimedia policies.

---

## Support

For issues, feature requests, or questions:
- Check the troubleshooting section
- Review Wikimedia Commons upload policies
- Consult Pywikibot documentation

**Remember:** Always verify your uploads on Commons and ensure they comply with licensing requirements.
