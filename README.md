# Student Attendance Analysis System

## Overview

Student Attendance Analysis System is a browser-based application for reviewing attendance data and generating focused attendance-range reports. Users can upload attendance workbooks in Excel (`.xlsx` or `.xls`) or PDF format, choose the relevant academic context and attendance range, then review student-wise and subject-wise results.

Attendance files are processed locally in the browser. Student data is not uploaded to a database or external processing service.

## Key Features

- Excel attendance file upload (`.xlsx` and `.xls`)
- PDF attendance file upload
- Excel sheet selection when multiple sheets are present
- Automatic subject detection from uploaded attendance data
- Dynamic subject-wise attendance analysis
- Year selection for I, II, III, or IV
- Section selection
- Report date selection
- Attendance range analysis:
  - Below 65%
  - 65%–74%
- Student-wise analysis
- Subject-wise student counts
- Overall subject-occurrence counts
- Subject-wise horizontal bar graph
- PDF report generation
- Excel report generation
- Automatic report filenames using report date, year, and section
- Multi-page PDF support
- PDF parsing for registration numbers, student names, and attendance percentages

## Data Processing Rules

- Subjects are detected dynamically from the uploaded file.
- Sample subject names are not hardcoded.
- `Counselling`, `Counseling`, `Total`, `TOTAL`, `TOTAL %`, and other aggregate or metadata columns are excluded from analysis.
- Registration number is the strongest student-row identifier.
- Student names can span multiple lines in PDFs.
- Multi-page PDFs are combined into one attendance dataset.
- When an Excel workbook has multiple sheets, only the selected sheet is analyzed.
- Uploaded attendance data remains in browser memory for the current session.

## Attendance Analysis Logic

- **Below 65%** means attendance `< 65%`.
- **65%–74%** means attendance `>= 65%` and `< 75%`.
- The `Students Matching Selected Range` value counts unique students with at least one matching subject.
- Subject-wise counts represent matching student-subject occurrences.
- `Overall Subject Occurrences` is the sum of all matching subject occurrences.

## Report Contents

Generated reports contain:

- Report date
- Year
- Section
- Department
- Selected attendance range
- Total students analyzed
- Students matching the selected range
- Actual subjects analyzed
- Subject-wise counts
- Subject-wise graph
- Student-wise analysis table

## Export Formats

The application can export the current analysis in both Excel and PDF formats. Both exports are generated from the same analyzed dataset and use the selected report date, year, section, attendance range, and dynamically detected subjects.

## PDF Report

The PDF export includes:

- Professional readable formatting
- Strong text contrast
- A structured summary
- Subject-wise counts
- A dynamically sized horizontal graph
- Wrapped subject labels
- Multi-page student tables
- Automatic table wrapping
- Repeated table headers on new pages
- Automatic page breaks
- No overlapping graph, table, or text content

## Technology

This project is part of a pnpm workspace and uses:

- React
- TypeScript
- Vite
- Wouter for client-side routing
- SheetJS (`xlsx`) for Excel parsing and export
- PDF.js (`pdfjs-dist`) for browser-side PDF text extraction
- Recharts for application charts
- Tailwind CSS and Radix UI components for the interface
- Lucide React for icons
- Framer Motion for interface motion
- TanStack React Query
- `date-fns` for date utilities

The application does not require authentication, a database, or a backend API for attendance processing.

## Project Structure

```text
.
├── artifacts/
│   └── student-attendance-analysis/
│       ├── public/
│       │   ├── favicon.svg
│       │   └── robots.txt
│       ├── src/
│       │   ├── App.tsx
│       │   ├── index.css
│       │   └── main.tsx
│       ├── components.json
│       ├── index.html
│       ├── package.json
│       ├── tsconfig.json
│       └── vite.config.ts
├── attached_assets/
├── package.json
├── pnpm-lock.yaml
├── pnpm-workspace.yaml
└── replit.md
```

Important application files:

- `artifacts/student-attendance-analysis/src/App.tsx` — application state, upload flow, Excel/PDF parsing, attendance analysis, charts, tables, and exports
- `artifacts/student-attendance-analysis/src/index.css` — visual theme and responsive layout rules
- `artifacts/student-attendance-analysis/src/main.tsx` — React application entry point
- `artifacts/student-attendance-analysis/package.json` — application scripts and dependencies
- `artifacts/student-attendance-analysis/vite.config.ts` — Vite configuration and artifact routing

## How to Run Locally

Install workspace dependencies from the repository root:

```bash
pnpm install
```

Run the attendance analysis web application:

```bash
PORT=5173 BASE_PATH=/ pnpm --filter @workspace/student-attendance-analysis run dev
```

The Vite configuration requires both `PORT` and `BASE_PATH`. The managed artifact workflow supplies these automatically.

Run the application typecheck:

```bash
pnpm --filter @workspace/student-attendance-analysis run typecheck
```

Build the application:

```bash
PORT=5173 BASE_PATH=/ pnpm --filter @workspace/student-attendance-analysis run build
```

Run the full workspace typecheck:

```bash
pnpm run typecheck
```

## Deployment

The production artifact builds the frontend as static files in:

```text
artifacts/student-attendance-analysis/dist/public
```

The project can be connected to GitHub and deployed using a suitable static or frontend hosting platform that supports the generated Vite build. The host should serve `index.html` for client-side routes and provide the required base path configuration when building.

## Future Enhancements

Marks analysis may be added in a future version. It is not currently part of the application.