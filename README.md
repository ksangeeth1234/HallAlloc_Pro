# Lecture Hall Allocation & Availability Reporting System (Web Edition)

A modern, professional, browser-based web application that converts and replaces the legacy Excel + VBA Macro workflow for lecture hall availability tracking and reporting.

---

## 📊 1. Detailed System Analysis

### A. Excel File & Workflow Analysis
The legacy reporting process relied on two primary worksheets inside `Lecture_Hall_Allocation_for_ongoing_programs.xlsx`:
1. **`Hall Availability Search`**: Contains the hall availability results table (rows 14–41) and interactive search inputs in cells `B7` (Day), `E7` (Time), and `B8` (Month Filter).
2. **`Class Schedule Analysis`**: Source table containing 163 class records. Key data columns include:
   - Column B (1): Weekday
   - Column E (4) & F (5): Start Time & End Time
   - Column G (6) & H (7): Start Minute & End Minute (since midnight)
   - Column J (9): Lecture Hall Name
   - Column L (11): Program / Batch Name
   - Column M (12): Delivery Mode (Physical, Hybrid, Online)
   - Column P (15) & Q (16): Start Date & End Date
   - Column S (18): Calculated Class Status (`Ongoing`, `Upcoming`, `Completed`, `Unknown`)

---

### B. VBA Macro Analysis & Recreation
The VBA procedure `RunHallSearch` in `HallSearchMacro.bas` executed the following algorithm:
1. Reads `searchDay` from `B7` and `searchTime` from `E7`. Converts `searchTime` to total minutes since midnight (`Hour*60 + Minute`).
2. Evaluates `searchMonth` from `B8`. If a month is selected, it sets a date window (`monthStart` to `monthEnd`). If `(Any Month)` is selected, the month filter is bypassed.
3. Loops through each lecture hall in `Hall Availability Search` (rows 14+).
4. Scans every record in `Class Schedule Analysis`:
   - Checks if `dDay == searchDay` AND `dHall == hallName` AND `dStatus` is either `Ongoing` or `Upcoming`.
   - Checks if `searchMinutes >= dStart` AND `searchMinutes < dEnd`.
   - If month filter is active: checks if `dStartDate <= monthEnd` AND `dEndDate >= monthStart`.
5. If matched, sets the hall status to **"Occupied"** and records the Program/Batch, Delivery Mode, and Time Slot. If unmatched, marks the hall as **"Free"**.

#### JavaScript Re-implementation
The function `runHallSearchLogic` in [`public/js/processing.js`](file:///c:/xampp/htdocs/php/Lecture%20Hall%20Allocation%20System/public/js/processing.js) faithfully port this exact algorithm without relying on Microsoft Excel or VBA execution.

---

## 🏗️ 2. Web Application Architecture

```
/project
├── public/
│   ├── index.html                    # Single Page Application Dashboard Shell
│   ├── css/
│   │   └── style.css                 # Custom CSS Design System
│   ├── js/
│   │   ├── app.js                    # Main UI Controller & Chart.js Integration
│   │   └── processing.js             # Recreated Macro Algorithm & Excel Parser
│   ├── Lecture_Hall_Allocation.xlsx  # Auto-loaded Sample Data
│   └── assets/                       # Icons & Images
├── api/
│   └── health.js                     # Vercel Serverless Function Health Check
├── package.json                      # Project dependencies & scripts
├── vercel.json                       # Vercel routing & static configuration
└── README.md                         # Documentation
```

---

## ✨ 3. Feature Highlights

- **Interactive Dashboard**: KPI metrics, conflict counter, room utilization stats, and Chart.js trend charts.
- **Drag-and-Drop Excel Upload**: Supports `.xlsx`, `.xls`, and `.xlsm` formats using SheetJS (`xlsx.full.min.js`).
- **Simulated VBA Macro Execution**: Instant calculation of hall occupancy by day, time, and month.
- **Comprehensive Search & Filtering**: Real-time text search, delivery mode filter (Physical, Hybrid, Online), and status filter with pagination.
- **Official Printable & Exportable Report**: Clean printable layout with CSV export capability.

---

## 💾 4. Database Assessment

**Database Requirement**: **Not Required** for basic single-user / file-driven reporting workflows.
- The web application operates purely client-side by parsing uploaded Excel workbooks into memory.
- If persistent report history, user authentication, or multi-user state is needed in future, **Vercel KV / Postgres** or **Supabase** can be seamlessly integrated.

---

## 🚀 5. Local Development Instructions

### Prerequisites
- [Node.js](https://nodejs.org/) (v16+) installed.

### Step-by-Step Run Instructions
1. Open VS Code terminal in the project directory.
2. Install dependencies (if needed):
   ```bash
   npm install
   ```
3. Start the local development web server:
   ```bash
   cmd /c "npx http-server public -p 3000"
   ```
4. Open your web browser and navigate to:
   ```
   http://localhost:3000
   ```
5. Click **"Execute Search Macro"** or drag and drop any updated `.xlsx` schedule file into the **Excel Upload** tab.

---

## ☁️ 6. Deployment to Vercel

1. Push this project folder to a GitHub, GitLab, or Bitbucket repository.
2. Log into [Vercel](https://vercel.com/) and click **"Add New Project"**.
3. Import your repository.
4. Framework Preset: **Other** (or Static Site).
5. Build Command: Leave blank (or `npm run build`).
6. Output Directory: `public`.
7. Click **Deploy**. Vercel will instantly host your app on a global CDN.

---

## 🔒 7. Security & Limitations

- **File Processing**: SheetJS operates locally inside the user's browser memory. No sensitive Excel data is transmitted to untrusted 3rd party servers.
- **Large File Limits**: Client-side parsing handles workbooks up to 50 MB smoothly.
