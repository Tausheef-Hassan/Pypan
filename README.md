# PyPan - Wikimedia Commons Batch Uploader

## ⚠️ IMPORTANT SECURITY WARNING
This program **stores your Wikimedia Commons username and password** in temporary configuration files (`user-config.py` and `user-password.py`) so that Pywikibot can log in.  
Although the files are restricted to the current user and cleaned up after the program exits, **your password will be written in plain text** during runtime.  

**Do not use this tool on shared or insecure machines. Do not share the generated config files. Always log out and remove files after use.**

---

## Overview
PyPan is a batch uploader for Wikimedia Commons built on top of Pywikibot. It allows you to select an Excel file containing file paths and metadata, then automatically upload them with progress tracking, retry logic, and logging.

---

## Difference from Pattypan
   - Retries errors 10 times and skipes where pattypan gets struck.  
   - Waits and retries if internet connection is down and auto starts uploading if internet is down.   
   - Writes results back to a new Excel file (`*_results.xlsx`) with a status column.
   - You can Pause uploading
   - Gives ETA

⚠️ This code does not have the ability to make a excel file like pattypan and I am currently working on that

---

## Excel File Format
The uploader expects an **Excel file (.xlsx or .xls)** with **no header row**.  
Each row must have exactly three columns:

1. **File Path** – Local path to the file on your computer.  
   Example: `C:\Users\Me\Pictures\photo.jpg`

2. **Target Filename** – The desired name of the file on Wikimedia Commons.  
   Example: `My_uploaded_photo.jpg`  
   - If no extension is provided, the program will add the correct one automatically.  
   - If the file already exists on Commons, it will be skipped.

3. **Description/Wikitext** – The full description text to include on the file’s Commons page.  
   Example: `A test image uploaded with PyPan. [[Category:Test uploads]]`  
   - You can include categories, templates, or any valid wikitext.  
   - The program appends a tracking category:  
     ```
     [[Category: Uploaded with pypan]]
     ```

### Example table

⚠️There should not be any header in the excel file, the items should start from row 1

| File Path | Target Filename | Wikitext |
|-----------|-----------------|----------|
| `D:\PID Project\Output\Batch 1\PID-0002220.jpg` | Syeda Rizwana Hasan Chunati Range Hill Cutting Inspection 2025-04-24 (PID-0002220) | =={{int:filedesc}}==<br> {{Information <br> \|description = {{bn\|1=পরিবেশ, বন ও জলবায়ু পরিবর্তন এবং পানি সম্পদ মন্ত্রণালয়ের উপদেষ্টা সৈয়দা রিজওয়ানা হাসান চট্টগ্রামের চুনতি রেঞ্জের আওতাধীন এলাকায় পাহাড় কাটা ও অবৈধ দখলের স্থানসমূহ সরেজমিনে পরিদর্শন করেন (বৃহস্পতিবার, ২৪ এপ্রিল ২০২৫)।}}{{en\|1=On Thursday, April 24, 2025, Syeda Rizwana Hasan, an advisor to the Ministry of Environment, Forest and Climate Change and Water Resources, conducted an on-site inspection of areas within Chattogram's Chunati Range that were affected by hill cutting and illegal encroachment. (Source: PID){{Auto-translated PID English description}}}} <br> \|date = {{Date-PID\|2025-04-24 11:41}} <br> \|source = {{Source-PID \| url=http://pressinform.gov.bd/sites/default/files/files/pressinform.portal.gov.bd/daily_photo/b85a9973_d92c_441c_a5e5_9382dbc43fff/2025-04-24-11-41-5e3716e3c901f81a77aa83b27733f65a.jpg}} <br> \|author = {{Institution:Press Information Department}} <br> \|permission =  <br> \|other versions =  }} <br><br> =={{int:license-header}}==<br> {{PD-BDGov-PID}}<br>[[Category: Syeda Rizwana Hasan]] |
| `D:\PID Project\Output\Batch 1\PID-0002221.jpg` | 2025-04-24 Syeda Rizwana Hasan Inspects Hill Cutting Chittagong Chunati Range (PID-0002221) | =={{int:filedesc}}==<br> {{Information <br> \|description = {{bn\|1=পরিবেশ, বন ও জলবায়ু পরিবর্তন এবং পানি সম্পদ মন্ত্রণালয়ের উপদেষ্টা সৈয়দা রিজওয়ানা হাসান চট্টগ্রামের চুনতি রেঞ্জের আওতাধীন এলাকায় পাহাড় কাটা ও অবৈধ দখলের স্থানসমূহ সরেজমিনে পরিদর্শন করেন (বৃহস্পতিবার, ২৪ এপ্রিল ২০২৫)।}}{{en\|1=On Thursday, April 24, 2025, Syeda Rizwana Hasan, Adviser to the Ministry of Environment, Forest and Climate Change and Water Resources, conducted an on-site inspection of hill cutting and illegal occupation sites within the Chunati Range area of Chittagong.{{Auto-translated PID English description}}}} <br> \|date = {{Date-PID\|2025-04-24 11:41}} <br> \|source = {{Source-PID \| url=http://pressinform.gov.bd/sites/default/files/files/pressinform.portal.gov.bd/daily_photo/6a6f71cb_cafc_4642_af85_3bb763483681/2025-04-24-11-41-c758e8ffba696106a8b3ad2c10285ab0.jpg}} <br> \|author = {{Institution:Press Information Department}} <br> \|permission =  <br> \|other versions =  }} <br><br> =={{int:license-header}}==<br> {{PD-BDGov-PID}}<br>[[Category: Syeda Rizwana Hasan]] |
| `D:\PID Project\Output\Batch 1\PID-0002222.jpg` | Bangladesh Embassy Tokyo E-passport Service Inauguration 2025-04-24 (PID-0002222) | =={{int:filedesc}}== <br{{Information <br> \|description = {{bn\|1=প্রধান উপদেষ্টার মুখ্য সচিব এম সিরাজ উদ্দিন মিয়া জাপানের টোকিওতে বাংলাদেশ দূতাবাসে ই-পাসপোর্ট সেবার কার্যক্রম উদ্বোধন করেন (বৃহস্পতিবার, ২৪ এপ্রিল ২০২৫)।}}{{en\|1=M. Siraj Uddin Miah, Chief Secretary to the Chief Adviser, inaugurated the e-passport service operations at the Bangladesh Embassy in Tokyo, Japan (Thursday, April 24, 2025). {{Auto-translated PID English description}}}} <br> \|date = {{Date-PID\|2025-04-24 11:34}} <br> \|source = {{Source-PID \| url=http://pressinform.gov.bd/sites/default/files/files/pressinform.portal.gov.bd/daily_photo/e0707e11_ab4c_476b_808e_5b5994545a3a/2025-04-24-11-34-06bb873d2e04669a537442eaf28f9774.jpg}} <br> \|author = {{Institution:Press Information Department}} <br> \|permission =  <br> \|other versions =  }} <br><br> =={{int:license-header}}==<br> {{PD-BDGov-PID}}<br>[[Category: Siraj Uddin Miah]] |

#### Results

https://commons.wikimedia.org/wiki/File:Syeda_Rizwana_Hasan_Chunati_Range_Hill_Cutting_Inspection_2025-04-24_(PID-0002220).jpg

https://commons.wikimedia.org/wiki/File:2025-04-24_Syeda_Rizwana_Hasan_Inspects_Hill_Cutting_Chittagong_Chunati_Range_(PID-0002221).jpg

https://commons.wikimedia.org/wiki/File:Bangladesh_Embassy_Tokyo_E-passport_Service_Inauguration_2025-04-24_(PID-0002222).jpg

### Excel formula

Excel formula example to make the wikitext column

```excel
="=={{int:filedesc}}==
{{Information
 |description = {{bn|1="&D1&"}}{{en|1="&C1&"{{Auto-translated PID English description}}}}
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


---

## How It Works
1. **Login Window**  
   - Enter your Wikimedia Commons username and password.  
   - Program generates temporary Pywikibot config files with your credentials.  
   - Password is stored in `user-password.py` (plain text) while program runs.

2. **Main Uploader**  
   - Select your Excel file.  
   - Configure retries, pauses, and warning handling.  
   - Start upload. Each file is checked for existence, extension validity, and Commons conflicts.  
   - Progress bar, logs, and statistics update live.

3. **Upload Logic**  
   - Skips missing or invalid files.  
   - Waits and retries if internet connection is down.  
   - Retries failed uploads up to the configured maximum attempts.  
   - Writes results back to a new Excel file (`*_results.xlsx`) with a status column.

4. **Cleanup**  
   - On exit, temporary login files and Pywikibot artifacts are deleted from the config directory.  

---

## Allowed File Types
Only Wikimedia Commons-supported extensions are allowed.  
Supported list in this tool:  .gif, .jpg, .jpeg, .mid, .midi, .ogg, .ogv, .oga, .png,
.svg, .xcf, .djvu, .pdf, .webm

---

## Output
- **Log messages** appear inside the application window.  
- **Result Excel file** is created alongside your input file with an additional column (`Upload_Status`) showing:  
  - `Success` – uploaded successfully  
  - `Skipped: reason` – skipped with explanation  
  - `Failed: reason` – failed with explanation  

---

## Notes
- Internet connection is checked before every upload.  
- All actions are logged with timestamps.  
- If running as a bundled executable, config files are stored in:  %USERPROFILE%.pypan
- If running as a script, config files are stored in the same directory as the script.  

---
