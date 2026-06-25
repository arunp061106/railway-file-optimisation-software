# 🚂 Railway File Organiser

**Fully offline. Works on any Windows PC. No internet required after first setup.**

---

## What It Does

| Step | What Happens |
|---|---|
| 1 | A file appears in your Downloads folder |
| 2 | The software detects it automatically |
| 3 | It analyses the file name and suggests a category |
| 4 | A popup asks you to confirm or change the category |
| 5 | The file is moved to the correct subfolder |
| 6 | Excel log opens with the file name + clickable hyperlink |

---

## First-Time Setup

> **You only need internet for first-time setup. After that, fully offline.**

1. Make sure **Python 3.9+** is installed → [python.org](https://python.org)
   - During install, check ✅ **"Add Python to PATH"**
2. Double-click **`SETUP.bat`** — installs all required libraries
3. Double-click **`START.bat`** — launches the app in the system tray
4. On first launch, a setup wizard will ask you to:
   - Select your **watch folder** (usually `Downloads`)
   - Select your **organised files folder** (e.g. `D:\Railway Files`)

---

## Daily Use

- Double-click **`START.bat`** to run (or set it to start with Windows — see Settings)
- The app runs quietly in the **system tray** (bottom-right of taskbar)
- Right-click the tray icon for options:
  - ▶ **Start / ⏸ Stop** watching
  - 📊 **Open Excel Log**
  - 📁 **Open Organised Files Folder**
  - ⚙ **Settings**
  - ❌ **Quit**

---

## File Categories (Built-In)

| Category | Example File Names |
|---|---|
| Tenders | `TDR_2024_NIT_Mumbai.pdf`, `Tender_Notice_45.pdf` |
| Circulars | `Circular_45_Safety.pdf`, `Notification_June.pdf` |
| Reports | `Monthly_Report_March.pdf`, `MPR_2024.xlsx` |
| Schedules | `Timetable_2024.pdf`, `Duty_Roster_May.xlsx` |
| Budget & Finance | `Budget_Estimate_2024.pdf`, `Sanction_Order.pdf` |
| Technical Drawings | `Drawing_Bridge_12.dwg`, `Signal_Design.pdf` |
| Correspondence | `Letter_DRM_Office.pdf`, `Memo_June.docx` |
| Policy & Rules | `Railway_Manual_2024.pdf`, `SOP_Safety.pdf` |
| Minutes of Meeting | `MoM_DRM_Meeting.docx`, `Minutes_June.pdf` |
| Contracts | `Work_Order_Bridge.pdf`, `Agreement_Contractor.pdf` |
| Safety & Accident | `Accident_Inquiry_Report.pdf`, `Safety_Audit.pdf` |
| Staff & HR | `Transfer_Order_May.pdf`, `Promotion_List.pdf` |
| Miscellaneous | Everything else |

---

## Customising Categories

Edit **`categories.json`** in the software folder to:
- Add new keywords to existing categories
- Create entirely new categories
- Change category colours

Example — adding "loco" to the Technical category:
```json
"Technical Drawings": {
    "keywords": ["drawing", "dwg", "loco", "...your new keyword..."],
    ...
}
```

---

## Excel Log

The file **`Railway_Files_Log.xlsx`** is created automatically in your organised files folder. It contains:

| Column | Content |
|---|---|
| S.No | Serial number |
| Date & Time | When the file was organised |
| Original File Name | The file's original name |
| Category | Where it was filed |
| Final File Name | Name after duplicate handling |
| Open File | 📂 Click to open directly |

---

## Building a Standalone .exe

To deploy on a PC that doesn't have Python:

1. Run `SETUP.bat` first
2. Double-click **`BUILD_EXE.bat`**
3. Find `RailwayFileOrganiser.exe` in the `dist\` folder
4. Copy `categories.json` and `config.json` alongside the `.exe`

---

## Troubleshooting

| Problem | Solution |
|---|---|
| App doesn't start | Check `railway_optimiser.log` in the software folder |
| Excel won't update | Close the Excel file first, then retry |
| Wrong category suggested | Edit `categories.json` to add better keywords |
| File not detected | Make sure the file is in the correct watch folder |
| Antivirus blocks .exe | Whitelist it in Windows Defender |

---

*Built for Indian Railways — 100% offline, secure, and customisable.*
