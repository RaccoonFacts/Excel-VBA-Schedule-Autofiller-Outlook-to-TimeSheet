# Excel VBA Schedule Auto-Filler: Outlook Email → Web Timesheet in Seconds

**From painful 40–60 minutes manual coloring → to a perfect timesheet with 3 clicks.**

### The Real Problem
Every week you receive your roster via Outlook email.  
The company web timesheet requires you to manually color **1,440 individual 1-minute blocks per day** using a clunky HTML table.  
Doing it by hand takes forever and is full of mistakes.

  <img width="1187" height="753" alt="image" src="https://github.com/user-attachments/assets/39af666c-e17d-47c2-bc96-2e0d6c8d8a66" />

  
### My Solution (What This Tool Does)
1. **Copy the entire roster email** (Ctrl+A → Ctrl+C in Outlook)
 <img width="645" height="753" alt="image" src="https://github.com/user-attachments/assets/bd341509-5bde-47ae-b99c-7dcab855512f" />
 
2. **Paste once** into the correct sheet:
   - PL1 roster → paste into **"PL1 Paste"** (starting at A7)
   - PL2 roster → paste into **"PL2 Paste"** (starting at A7)
   <img width="1526" height="776" alt="image" src="https://github.com/user-attachments/assets/675ee446-8763-44cc-9a86-26bf8585f182" />

3. Word detection to mark correct types of times
4. Click one button → choose **Pl1All** or **Pl2All**
5. The macro instantly:
   - Cleans and formats the pasted times
   - Breaks the day into every single minute (B6:B1445)
   <img width="1699" height="686" alt="image" src="https://github.com/user-attachments/assets/5aeb4d58-d0db-41be-accb-783c8c732d43" />

 

   - Applies the correct hex color to column E for every minute you’re working
<img width="1256" height="753" alt="image" src="https://github.com/user-attachments/assets/891a4bc5-ada2-4568-af17-d728a994e6db" />


   - Generates a perfect "COPY THIS" sheet
6. Press **Ctrl+A → Ctrl+C** on the "PL1/PL2 COPY THIS" tab
7. Open the web timesheet → F12 (DevTools) → paste → **entire week is filled perfectly**
<img width="1105" height="753" alt="image" src="https://github.com/user-attachments/assets/2e0098fb-6136-4628-9a9b-ee2ed8f3cf77" />

**Total time: under 15 seconds per week**





### Color Legend
| Color   | Hex       | Meaning                  |
|---------|-----------|--------------------------|
| Green   | `#59bf6b` | Normal shift (from B3–B4 or C3–C4) |
| Orange/Red | `#ff6633` | Empty / standby shifts   |
| Blue    | `#0099ff` | Test or training blocks  |
| Purple  | `#8000ff` | Default/background       |

### Key Features
- Automatically fixes Outlook’s weird dashes (`–` → `-`) and AM/PM junk
- Handles overnight shifts perfectly (e.g., 23:30–02:15)
- Tolerance matching — never misses a minute even if Outlook rounds differently
- One-click export to clean, tab-delimited text (opens automatically in Notepad)
- Full support for both PL1 and PL2 rosters in the same workbook

### How to Use (Step-by-Step)
1. Download → `Schedule-AutoFiller.xlsm`
2. Open your roster email in Outlook → Select All → Copy
3. In Excel, go to **PL1 Paste** or **PL2 Paste** → click cell A7 → Paste
4. Press **Alt + F8** → double-click **Pl1All** or **Pl2All**
5. When it finishes, go to **PL1 COPY THIS** or **PL2 COPY THIS**
6. Ctrl+A → Ctrl+C
7. Open your web timesheet → press F12 → paste into the color field

→ Done. Zero manual entries. Estimated time saved per week: 6 hours.
