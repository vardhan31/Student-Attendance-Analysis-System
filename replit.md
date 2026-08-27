# Student Attendance Analysis System

A local-first React application that analyzes uploaded attendance workbooks and exports focused student-subject reports.

## Run & Operate

- `pnpm --filter @workspace/student-attendance-analysis run dev` — run the web app
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm --filter @workspace/student-attendance-analysis run typecheck` — typecheck the web app
- `pnpm --filter @workspace/student-attendance-analysis run build` — build the web app

The attendance workbook is parsed and held in browser memory only. No database, authentication, backend API, or persistent student data storage is used.

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- Frontend: React + Vite + Wouter
- Workbook processing: SheetJS (`xlsx`)
- Visualizations: Recharts
- Icons and motion: Lucide React + Framer Motion

## Where things live

- `artifacts/student-attendance-analysis/src/App.tsx` — application state, workbook parsing, analysis, charts, tables, and export
- `artifacts/student-attendance-analysis/src/index.css` — visual theme and responsive layout rules
- `artifacts/student-attendance-analysis/package.json` — web app dependencies

## Architecture decisions

- Workbook data remains local to the browser to keep sensitive student information private.
- Sheet names and subject columns are detected from each workbook rather than hardcoded.
- Registration number is the stable student identity used for de-duplication and counts.
- The selected range is the single source of truth for student matches, subject counts, charts, and exports.

## Product

- Overview dashboard with empty and analyzed states
- Analysis workspace with `.xlsx` drag-and-drop upload, sheet selection, academic context, threshold selection, warnings, search, sort, and responsive charts
- Reports view with summary metadata, subject counts, student-wise results, and a four-sheet Excel export named from date, year, and section

## User preferences

- Display owner: A.N. Harshith Vardhan
- Department: Computer Science and Engineering

## Gotchas

- The web workflow must provide `PORT` and `BASE_PATH`; do not run the Vite app from the workspace root.
- Decimal ratios such as `0.5581`, numeric percentages such as `55.81`, and strings such as `55.81%` are normalized to the same percentage scale.
- `Counseling` and `TOTAL %` must never be treated as subjects.

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
