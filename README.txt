RECIPE PICTURE -> EXCEL  (works offline)

ONE-TIME SETUP (needs internet only this once)
1. Install Python 3.9+ from python.org (tick "Add Python to PATH").
2. Open Command Prompt in this folder and run:
       pip install -r requirements.txt
   (The OCR models are included in the package, so nothing else is downloaded later.)

USE (no internet needed)
- Double-click run.bat  -> click "Choose picture(s)..." -> pick your recipe photos.
- Or drag one or more pictures onto run.bat.
- An .xlsx file with the same name is saved next to each picture.

TIPS FOR BEST RESULTS
- Straight, sharp, well-lit picture; crop to just the table.
- Tables with visible borders are rebuilt cell-by-cell (most accurate).
- Always check the numbers once in Excel; OCR can misread blurry characters.
- The first row is formatted as the header.
