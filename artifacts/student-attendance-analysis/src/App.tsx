import { useMemo, useRef, useState, type ChangeEvent, type DragEvent, type ReactNode } from "react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { ErrorBoundary } from "@/components/error-boundary";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { BarChart, Bar, CartesianGrid, Cell, Pie, PieChart, ResponsiveContainer, Tooltip as ChartTooltip, XAxis, YAxis } from "recharts";
import { AlertTriangle, ArrowDownUp, BarChart3, BookOpen, CalendarDays, Check, ChevronDown, CircleHelp, Download, FileDown, FileSpreadsheet, FileText, Files, Filter, LayoutDashboard, Loader2, Menu, RefreshCw, Search, ShieldCheck, SlidersHorizontal, UploadCloud, X } from "lucide-react";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import pdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import { Link, Route, Switch, useLocation, Router as WouterRouter } from "wouter";
import NotFound from "@/pages/not-found";

const queryClient = new QueryClient();
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

type AttendanceRange = "below65" | "65to74";
type AttendanceRow = { serialNumber: string; registrationNumber: string; name: string; values: Record<string, number | null> };
type UploadedWorkbook = { fileName: string; fileType: "excel" | "pdf"; sheets: string[]; activeSheet: string };
type AnalysisConfig = { year: string; section: string; reportDate: string; sheetName: string; range: AttendanceRange };
type StudentMatch = { subject: string; percentage: number };
type StudentResult = { serialNumber: string; registrationNumber: string; year: string; section: string; name: string; matches: StudentMatch[]; matchCount: number };
type SubjectCount = { subject: string; studentCount: number };
type AnalysisSummary = { totalStudents: number; totalSubjects: number; uniqueStudents: number; totalCases: number; selectedYear: string; selectedSection: string; selectedDate: string; selectedDateIso: string; sourceSheet: string; rangeLabel: string };
type Analysis = { summary: AnalysisSummary; students: StudentResult[]; subjects: SubjectCount[]; rows: AttendanceRow[] };
type ParseResult = { workbook: UploadedWorkbook; rows: Record<string, AttendanceRow[]>; subjects: Record<string, string[]>; warnings: string[] };
type AppState = { workbook: UploadedWorkbook | null; rows: Record<string, AttendanceRow[]>; subjects: Record<string, string[]>; warnings: string[]; analysis: Analysis | null; config: AnalysisConfig; parsing: boolean; parseError: string };

const emptyConfig: AnalysisConfig = { year: "", section: "", reportDate: "", sheetName: "", range: "below65" };
const rangeLabels: Record<AttendanceRange, string> = { below65: "Below 65%", "65to74": "65% – 74%" };
const chartColors = ["#176b67", "#e77b5d", "#3d7aa8", "#a58658", "#8065a9", "#4c8d77"];

function safeText(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizeHeader(value: unknown): string {
  return safeText(value).toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

function parseAttendance(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return null;
    return value >= 0 && value <= 1 ? value * 100 : value >= 0 && value <= 100 ? value : null;
  }
  const text = safeText(value).replace(/,/g, "");
  if (!text || /^[-–—]$/.test(text)) return null;
  const hasPercent = text.includes("%");
  const parsed = Number(text.replace("%", "").trim());
  if (!Number.isFinite(parsed) || parsed < 0) return null;
  if (!hasPercent && parsed <= 1) return parsed * 100;
  return parsed <= 100 ? parsed : null;
}

function looksLikeMetadata(header: string): boolean {
  const normalized = normalizeHeader(header);
  return /^(SL|S NO|SERIAL|NO|COUNSEL(LING)?|TOTAL|TOTAL %|TOTAL PERCENTAGE|AVERAGE|ATTENDANCE|OVERALL|AGGREGATE|PERCENT|PERCENTAGE|SUMMARY|STATUS|REMARKS?|REGD|REGISTRATION|NAME)$/.test(normalized)
    || /(COUNSEL|TOTAL|AVERAGE|OVERALL|AGGREGATE|PERCENT|SUMMARY|STATUS|REMARK|CONDUCTED|HOURS|PROGRAM|BRANCH|DEPARTMENT|SEMESTER|YEAR|SECTION|CLASS|COLLEGE|ROLL|REGISTRATION|REGD)/.test(normalized);
}

function parseSheet(sheet: XLSX.WorkSheet, sheetName: string): { rows: AttendanceRow[]; subjects: string[]; warnings: string[] } {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null }) as unknown[][];
  const warnings: string[] = [];
  if (!matrix.length) return { rows: [], subjects: [], warnings: [`${sheetName} is blank and was skipped.`] };
  let headerIndex = matrix.findIndex((row) => row.slice(0, 25).some((cell) => /REGD|REGISTRATION|REG NO|ROLL NO/.test(normalizeHeader(cell))) && row.some((cell) => normalizeHeader(cell) === "NAME" || normalizeHeader(cell).includes("STUDENT NAME")));
  if (headerIndex < 0) headerIndex = matrix.slice(0, 12).findIndex((row) => row.some((cell) => normalizeHeader(cell) === "NAME" || normalizeHeader(cell).includes("STUDENT NAME")));
  if (headerIndex < 0) {
    warnings.push(`${sheetName}: could not find a registration number and name header row.`);
    return { rows: [], subjects: [], warnings };
  }
  const headers = matrix[headerIndex] ?? [];
  const serialIndex = headers.findIndex((cell) => /^(SL|S NO|SERIAL)$/.test(normalizeHeader(cell)));
  const regIndex = headers.findIndex((cell) => /REGD|REGISTRATION|REG NO|ROLL NO/.test(normalizeHeader(cell)));
  const nameIndex = headers.findIndex((cell) => normalizeHeader(cell) === "NAME" || normalizeHeader(cell).includes("STUDENT NAME"));
  if (regIndex < 0 || nameIndex < 0) {
    warnings.push(`${sheetName}: required REGD.NO and NAME columns were not detected.`);
    return { rows: [], subjects: [], warnings };
  }
  const numericColumns = headers.map((header, index) => {
    const headerText = safeText(header);
    if (!headerText || index === regIndex || index === nameIndex || looksLikeMetadata(headerText)) return false;
    const samples = matrix.slice(headerIndex + 1).map((row) => parseAttendance(row[index])).filter((value): value is number => value !== null);
    return samples.length > 0;
  });
  const subjects = headers.filter((header, index) => numericColumns[index] && safeText(header)).map((header) => safeText(header));
  if (!subjects.length) warnings.push(`${sheetName}: no numeric attendance subject columns were found.`);
  const duplicateRegs = new Set<string>();
  const seen = new Set<string>();
  const rows: AttendanceRow[] = [];
  matrix.slice(headerIndex + 1).forEach((row) => {
    const registrationNumber = safeText(row[regIndex]);
    const name = safeText(row[nameIndex]);
    if (!registrationNumber || !name || /^total|^average|^summary/i.test(name)) return;
    const key = registrationNumber.toUpperCase();
    if (seen.has(key)) duplicateRegs.add(registrationNumber);
    seen.add(key);
    const values: Record<string, number | null> = {};
    subjects.forEach((subject) => { values[subject] = parseAttendance(row[headers.findIndex((header) => safeText(header) === subject)]); });
    rows.push({ serialNumber: serialIndex >= 0 ? safeText(row[serialIndex]) : String(rows.length + 1), registrationNumber, name, values });
  });
  if (duplicateRegs.size) warnings.push(`${sheetName}: duplicate registration number${duplicateRegs.size === 1 ? "" : "s"} found (${Array.from(duplicateRegs).slice(0, 4).join(", ")}). The first record is used in analysis.`);
  if (!rows.length) warnings.push(`${sheetName}: no student records were found below the header.`);
  return { rows, subjects, warnings };
}

const PDF_DOCUMENT_KEY = "__pdf_document__";
type PdfCell = { text: string; x: number; y: number; width: number };
type PdfLine = { y: number; items: PdfCell[] };
type PdfHeaderMapping = {
  headerY: number;
  serial?: PdfCell;
  registration: PdfCell;
  name: PdfCell;
  subjects: PdfCell[];
};

function groupPdfLines(items: PdfCell[]): PdfLine[] {
  const lines: PdfLine[] = [];
  items.forEach((item) => {
    const line = lines.find((candidate) => Math.abs(candidate.y - item.y) <= 4);
    if (line) line.items.push(item);
    else lines.push({ y: item.y, items: [item] });
  });
  return lines.map((line) => ({ ...line, items: line.items.sort((a, b) => a.x - b.x) })).sort((a, b) => b.y - a.y);
}

function pdfColumnMatch(line: PdfLine, x: number): PdfCell | undefined {
  return line.items.filter((item) => Math.abs(item.x - x) < 80).sort((a, b) => Math.abs(a.x - x) - Math.abs(b.x - x))[0];
}

function findPdfHeaderCell(line: PdfLine, pattern: RegExp): PdfCell | undefined {
  for (let index = 0; index < line.items.length; index += 1) {
    for (let size = 1; size <= Math.min(3, line.items.length - index); size += 1) {
      const group = line.items.slice(index, index + size);
      const text = normalizeHeader(group.map((item) => item.text).join(" "));
      if (pattern.test(text)) {
        const first = group[0];
        const last = group[group.length - 1];
        return { text: group.map((item) => item.text).join(" "), x: first.x, y: first.y, width: Math.max(first.width, last.x + last.width - first.x) };
      }
    }
  }
  return undefined;
}

function isPdfRegistrationCandidate(value: unknown): boolean {
  const text = safeText(value).replace(/\s+/g, "");
  if (parseAttendance(text) !== null || text.length < 4) return false;
  return /^[A-Z0-9./-]+$/i.test(text) && /\d/.test(text) && /[A-Z]/i.test(text);
}

function findPdfHeaderMapping(lines: PdfLine[]): PdfHeaderMapping | null {
  const completeHeaderLine = lines.find((line) => {
    const text = normalizeHeader(line.items.map((item) => item.text).join(" "));
    return /(REGD|REGISTRATION|REG NO|ROLL NO)/.test(text) && /(NAME|STUDENT NAME)/.test(text);
  });
  let headerLine = completeHeaderLine;
  if (!headerLine) {
    const registrationLine = lines.find((line) => /(REGD|REGISTRATION|REG NO|ROLL NO)/.test(normalizeHeader(line.items.map((item) => item.text).join(" "))));
    const nameLine = lines.find((line) => /(NAME|STUDENT NAME)/.test(normalizeHeader(line.items.map((item) => item.text).join(" "))));
    if (registrationLine && nameLine && Math.abs(registrationLine.y - nameLine.y) <= 36) {
      const related = lines.filter((line) => Math.abs(line.y - Math.max(registrationLine.y, nameLine.y)) <= 36);
      headerLine = { y: Math.max(...related.map((line) => line.y)), items: related.flatMap((line) => line.items).sort((a, b) => a.x - b.x) };
    }
  }
  if (!headerLine) return null;
  const registration = findPdfHeaderCell(headerLine, /(REGD|REGISTRATION|REG NO|ROLL NO)/);
  const name = findPdfHeaderCell(headerLine, /^(NAME|STUDENT NAME)$/);
  const serial = findPdfHeaderCell(headerLine, /^(SL|S NO|SERIAL)$/);
  if (!registration || !name) return null;
  const dataLines = lines.filter((line) => line.y < headerLine.y && !/(REGD|REGISTRATION|REG NO|ROLL NO)/.test(normalizeHeader(line.items.map((item) => item.text).join(" "))));
  const subjects = headerLine.items
    .filter((item) => item !== registration && item !== name && !looksLikeMetadata(item.text))
    .filter((item) => dataLines.some((line) => {
      const cell = pdfColumnMatch(line, item.x);
      return cell && parseAttendance(cell.text) !== null;
    }));
  return { headerY: headerLine.y, serial, registration, name, subjects: Array.from(new Map(subjects.map((item) => [safeText(item.text), item])).values()) };
}

function samePdfSubjects(first: PdfHeaderMapping, next: PdfHeaderMapping): boolean {
  const firstNames = first.subjects.map((item) => normalizeHeader(item.text)).join("|");
  const nextNames = next.subjects.map((item) => normalizeHeader(item.text)).join("|");
  return firstNames === nextNames;
}

function parsePdfRows(lines: PdfLine[], mapping: PdfHeaderMapping, pageName: string): { rows: AttendanceRow[]; duplicateRegs: Set<string> } {
  const dataLines = lines.filter((line) => {
    const lineText = normalizeHeader(line.items.map((item) => item.text).join(" "));
    return !/(REGD|REGISTRATION|REG NO|ROLL NO)/.test(lineText) && line.y < mapping.headerY;
  });
  const subjectColumns = mapping.subjects;
  const firstSubjectX = subjectColumns[0]?.x ?? Number.POSITIVE_INFINITY;
  const rows: AttendanceRow[] = [];
  const duplicateRegs = new Set<string>();
  const seenRegs = new Set<string>();
  let lastRow: AttendanceRow | null = null;
  let lastRowY = Number.POSITIVE_INFINITY;
  dataLines.forEach((line) => {
    const regCell = pdfColumnMatch(line, mapping.registration.x);
    const registrationNumber = safeText(regCell?.text);
    if (!registrationNumber || !/[A-Za-z0-9]{3,}/.test(registrationNumber)) {
      const continuation = line.items
        .filter((item) => item.x >= mapping.name.x - 8 && item.x < firstSubjectX - 5)
        .map((item) => item.text)
        .join(" ")
        .trim();
      const hasSubjectValue = subjectColumns.some((column) => parseAttendance(pdfColumnMatch(line, column.x)?.text) !== null);
      if (lastRow && continuation && !hasSubjectValue && lastRowY - line.y <= 32 && !looksLikeMetadata(continuation)) {
        lastRow.name = `${lastRow.name} ${continuation}`.replace(/\s+/g, " ").trim();
      }
      return;
    }
    const name = line.items
      .filter((item) => item !== regCell && item.x >= mapping.name.x - 8 && item.x < firstSubjectX - 5)
      .map((item) => item.text)
      .join(" ")
      .trim();
    if (!name || /^(TOTAL|AVERAGE|SUMMARY)$/i.test(name)) return;
    const key = registrationNumber.toUpperCase();
    if (seenRegs.has(key)) duplicateRegs.add(registrationNumber);
    seenRegs.add(key);
    const values: Record<string, number | null> = {};
    subjectColumns.forEach((column) => {
      const cell = pdfColumnMatch(line, column.x);
      values[safeText(column.text)] = parseAttendance(cell?.text);
    });
    const serialNumber = mapping.serial ? safeText(pdfColumnMatch(line, mapping.serial.x)?.text) : String(rows.length + 1);
    lastRow = { serialNumber, registrationNumber, name, values };
    lastRowY = line.y;
    rows.push(lastRow);
  });
  return { rows, duplicateRegs };
}

function parsePdfRowsFromAnchors(lines: PdfLine[], serialOffset: number): { rows: AttendanceRow[]; percentageCount: number } {
  const rows: AttendanceRow[] = [];
  let lastRow: AttendanceRow | null = null;
  let lastRowY = Number.POSITIVE_INFINITY;
  let lastRegistrationX = 0;
  let lastFirstNumericX = Number.POSITIVE_INFINITY;
  let percentageCount = 0;
  lines.forEach((line) => {
    const registrationCell = line.items.find((item) => isPdfRegistrationCandidate(item.text));
    if (!registrationCell) {
      const continuation = line.items
        .filter((item) => item.x > lastRegistrationX && item.x < lastFirstNumericX && !isPdfRegistrationCandidate(item.text) && parseAttendance(item.text) === null)
        .map((item) => item.text)
        .join(" ")
        .trim();
      if (lastRow && continuation && lastRowY - line.y <= 32 && !looksLikeMetadata(continuation)) lastRow.name = `${lastRow.name} ${continuation}`.replace(/\s+/g, " ").trim();
      return;
    }
    const numericItems = line.items.filter((item) => item.x > registrationCell.x && parseAttendance(item.text) !== null);
    if (!numericItems.length) return;
    const firstNumericX = numericItems[0].x;
    const name = line.items.filter((item) => item !== registrationCell && item.x > registrationCell.x && item.x < firstNumericX && parseAttendance(item.text) === null).map((item) => item.text).join(" ").trim();
    if (!name || looksLikeMetadata(name)) return;
    const values: Record<string, number | null> = {};
    numericItems.forEach((item, index) => { values[`Subject ${index + 1}`] = parseAttendance(item.text); });
    percentageCount += numericItems.length;
    lastRow = { serialNumber: String(serialOffset + rows.length + 1), registrationNumber: safeText(registrationCell.text), name, values };
    lastRowY = line.y;
    lastRegistrationX = registrationCell.x;
    lastFirstNumericX = firstNumericX;
    rows.push(lastRow);
  });
  return { rows, percentageCount };
}

function debugPdf(label: string, payload: unknown): void {
  if (import.meta.env.DEV && typeof window !== "undefined" && new URLSearchParams(window.location.search).has("debugPdf")) console.debug(`[attendance-pdf] ${label}`, payload);
}

async function readPdf(file: File): Promise<ParseResult> {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data, useWorkerFetch: false, isEvalSupported: false }).promise;
  if (!pdf.numPages) throw new Error("This PDF does not contain any pages.");
  const rows: Record<string, AttendanceRow[]> = { [PDF_DOCUMENT_KEY]: [] };
  const subjects: Record<string, string[]> = { [PDF_DOCUMENT_KEY]: [] };
  const warnings: string[] = [];
  let mapping: PdfHeaderMapping | null = null;
  let headerFound = false;
  const duplicateRegs = new Set<string>();
  let fallbackOffset = 0;
  debugPdf("document", { pages: pdf.numPages, fileName: file.name });
  for (let index = 0; index < pdf.numPages; index += 1) {
    const page = await pdf.getPage(index + 1);
    const content = await page.getTextContent();
    const items = (content.items as Array<{ str?: unknown; transform?: number[]; width?: number }>)
      .filter((item) => typeof item.str === "string" && item.str.trim() && Array.isArray(item.transform))
      .map((item) => ({ text: safeText(item.str), x: item.transform?.[4] ?? 0, y: item.transform?.[5] ?? 0, width: item.width ?? 0 }));
    const lines = groupPdfLines(items);
    debugPdf(`page ${index + 1}`, { itemCount: items.length, items: items.slice(0, 50).map((item) => ({ text: item.text, x: item.x, y: item.y })), registrationCandidates: items.filter((item) => isPdfRegistrationCandidate(item.text)).map((item) => item.text), percentageValues: items.filter((item) => parseAttendance(item.text) !== null).map((item) => item.text) });
    const detectedMapping = findPdfHeaderMapping(lines);
    if (detectedMapping) {
      if (mapping && !samePdfSubjects(mapping, detectedMapping)) {
        warnings.push(`Page ${index + 1}: a different table structure was skipped.`);
        continue;
      }
      mapping = detectedMapping;
      headerFound = true;
    }
    const parsed = mapping?.subjects.length ? parsePdfRows(lines, mapping, `Page ${index + 1}`) : parsePdfRowsFromAnchors(lines, fallbackOffset);
    rows[PDF_DOCUMENT_KEY].push(...parsed.rows);
    if ("duplicateRegs" in parsed) parsed.duplicateRegs.forEach((registrationNumber) => duplicateRegs.add(registrationNumber));
    if (mapping?.subjects.length) subjects[PDF_DOCUMENT_KEY] = mapping.subjects.map((item) => safeText(item.text)).filter(Boolean);
    else if (!subjects[PDF_DOCUMENT_KEY].length && parsed.rows.length) subjects[PDF_DOCUMENT_KEY] = Object.keys(parsed.rows[0].values);
    fallbackOffset += parsed.rows.length;
  }
  if (!headerFound) warnings.push("The PDF did not contain a recognizable registration number, student name, and subject header.");
  if (mapping && !mapping.subjects.length) warnings.push("The PDF header was found, but no numeric subject attendance columns were detected.");
  if (duplicateRegs.size) warnings.push("Duplicate registration numbers were found across the PDF. The first record is used in analysis.");
  debugPdf("final", { headerFound, students: rows[PDF_DOCUMENT_KEY].length, subjects: subjects[PDF_DOCUMENT_KEY] });
  if (!rows[PDF_DOCUMENT_KEY].length) throw new Error("No attendance table was detected in this PDF. Please upload a PDF with registration number, student name, and subject attendance columns.");
  return { workbook: { fileName: file.name, fileType: "pdf", sheets: [], activeSheet: PDF_DOCUMENT_KEY }, rows, subjects, warnings };
}

async function readAttendanceFile(file: File): Promise<ParseResult> {
  const extension = file.name.toLowerCase().split(".").pop();
  if (extension === "pdf") return readPdf(file);
  if (extension !== "xlsx" && extension !== "xls") throw new Error("Please choose an Excel (.xlsx or .xls) or PDF (.pdf) attendance file.");
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
  const sheets = workbook.SheetNames;
  if (!sheets.length) throw new Error("This workbook does not contain any sheets.");
  const rows: Record<string, AttendanceRow[]> = {};
  const subjects: Record<string, string[]> = {};
  const warnings: string[] = [];
  sheets.forEach((sheetName: string) => {
    const parsed = parseSheet(workbook.Sheets[sheetName], sheetName);
    rows[sheetName] = parsed.rows;
    subjects[sheetName] = parsed.subjects;
    warnings.push(...parsed.warnings);
  });
  return { workbook: { fileName: file.name, fileType: "excel", sheets, activeSheet: sheets[0] }, rows, subjects, warnings };
}

function formatDateForReport(value: string): string {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  return year && month && day ? `${day}-${month}-${year}` : value;
}

function buildAnalysis(state: AppState): Analysis | null {
  const { config, rows, subjects } = state;
  if (!config.year.trim() || !config.section.trim() || !config.reportDate || !config.sheetName || !rows[config.sheetName]) return null;
  const threshold = config.range === "below65" ? (value: number) => value < 65 : (value: number) => value >= 65 && value < 75;
  const unique = new Map<string, AttendanceRow>();
  rows[config.sheetName].forEach((row) => {
    const key = row.registrationNumber.toUpperCase();
    if (!unique.has(key)) unique.set(key, row);
  });
  const students: StudentResult[] = Array.from(unique.values()).map((row) => {
    const matches = (subjects[config.sheetName] ?? []).filter((subject) => row.values[subject] !== null && threshold(row.values[subject] as number)).map((subject) => ({ subject, percentage: Math.round((row.values[subject] as number) * 100) / 100 }));
    return { serialNumber: row.serialNumber, registrationNumber: row.registrationNumber, name: row.name, year: config.year.trim(), section: config.section.trim(), matches, matchCount: matches.length };
  }).filter((student) => student.matchCount > 0);
  const subjectMap = new Map<string, number>((subjects[config.sheetName] ?? []).map((subject) => [subject, 0]));
  students.forEach((student) => student.matches.forEach(({ subject }) => subjectMap.set(subject, (subjectMap.get(subject) ?? 0) + 1)));
  const subjectCounts = Array.from(subjectMap.entries()).map(([subject, studentCount]) => ({ subject, studentCount }));
  return {
    summary: {
      totalStudents: unique.size,
      totalSubjects: (subjects[config.sheetName] ?? []).length,
      uniqueStudents: students.length,
      totalCases: students.reduce((sum, student) => sum + student.matchCount, 0),
      selectedYear: config.year.trim(),
      selectedSection: config.section.trim(),
      selectedDate: formatDateForReport(config.reportDate),
      selectedDateIso: config.reportDate,
      sourceSheet: state.workbook?.fileType === "pdf" ? "Complete PDF document" : config.sheetName,
      rangeLabel: rangeLabels[config.range],
    },
    students,
    subjects: subjectCounts,
    rows: Array.from(unique.values()),
  };
}

function exportAnalysis(analysis: Analysis): void {
  const summary = analysis.summary;
  const workbook = XLSX.utils.book_new();
  const safe = (value: unknown) => safeText(value).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "-") || "report";
  const studentRows = analysis.students.map((student) => ({
    "S.No": student.serialNumber,
    "Registration Number": student.registrationNumber,
    Name: student.name,
    Year: student.year,
    Section: student.section,
    "Low Attendance Subjects": student.matches.map((match) => `${match.subject} - ${match.percentage.toFixed(2)}%`).join("; "),
    "Match count": student.matchCount,
  }));
  const subjectRows = analysis.subjects.map((subject) => ({ Subject: subject.subject, "Student count": subject.studentCount }));
  const detailRows = analysis.students.flatMap((student) => student.matches.map((match) => ({
    "S.No": student.serialNumber, "Registration Number": student.registrationNumber, Name: student.name, Year: student.year, Section: student.section, Subject: match.subject, Attendance: `${match.percentage.toFixed(2)}%`,
  })));
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([{
      "Report date": summary.selectedDate, "Source sheet": summary.sourceSheet, Year: summary.selectedYear, Section: summary.selectedSection, Range: summary.rangeLabel,
    "Total students": summary.totalStudents, "Subjects detected": summary.totalSubjects, "Students with matches": summary.uniqueStudents, "Subject-wise cases": summary.totalCases,
  }]), "Summary");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(studentRows), "Student-wise Analysis");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(subjectRows), "Subject-wise Analysis");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detailRows), "Detailed Analysis");
  XLSX.writeFile(workbook, `${safe(summary.selectedDateIso)}-${safe(summary.selectedYear)}-${safe(summary.selectedSection)}.xlsx`);
}

function pdfSafe(value: unknown): string {
  return safeText(value).replace(/[^\x20-\x7E]/g, "-").replace(/[\\()]/g, (character) => `\\${character}`);
}

function buildPdfBlob(analysis: Analysis): Blob {
  const summary = analysis.summary;
  const margin = 42;
  const encoder = new TextEncoder();
  const pageContents: string[] = [];
  const text = (commands: string[], value: unknown, x: number, y: number, size = 9, bold = false) => {
    commands.push(`0.12 0.15 0.18 rg BT /${bold ? "F2" : "F1"} ${size} Tf ${x} ${y} Td (${pdfSafe(value)}) Tj ET`);
  };
  const wrap = (value: unknown, maxChars: number) => {
    const lines: string[] = [];
    let current = "";
    pdfSafe(value).split(/\s+/).forEach((word) => {
      if (!word) return;
      if (current && `${current} ${word}`.length > maxChars) {
        lines.push(current);
        current = word;
      } else current = current ? `${current} ${word}` : word;
    });
    if (current || !lines.length) lines.push(current);
    return lines;
  };
  const drawTable = (commands: string[], startY: number, columns: Array<{ title: string; width: number; maxChars: number }>, rows: string[][], headerColor: "teal" | "sand" = "sand") => {
    let y = startY;
    const totalWidth = columns.reduce((sum, column) => sum + column.width, 0);
    commands.push(`${headerColor === "teal" ? "0.09 0.42 0.40" : "0.94 0.91 0.84"} rg ${margin} ${y - 20} ${totalWidth} 20 re f`);
    let x = margin;
    columns.forEach((column) => { text(commands, column.title, x + 5, y - 14, 8, true); x += column.width; });
    y -= 20;
    rows.forEach((row) => {
      const wrapped = row.map((cell, index) => wrap(cell, columns[index].maxChars));
      const rowHeight = Math.max(24, ...wrapped.map((lines) => lines.length * 10 + 10));
      commands.push(`0.86 0.84 0.79 RG ${margin} ${y - rowHeight} ${totalWidth} ${rowHeight} re S`);
      x = margin;
      wrapped.forEach((lines, index) => {
        lines.forEach((line, lineIndex) => text(commands, line, x + 5, y - 13 - lineIndex * 10, 8));
        commands.push(`0.86 0.84 0.79 RG ${x} ${y - rowHeight} m ${x} ${y} l S`);
        x += columns[index].width;
      });
      commands.push(`0.86 0.84 0.79 RG ${x} ${y - rowHeight} m ${x} ${y} l S`);
      y -= rowHeight;
    });
    return y;
  };

  const firstPage: string[] = [];
  text(firstPage, "STUDENT ATTENDANCE ANALYSIS REPORT", margin, 748, 16, true);
  text(firstPage, `${summary.selectedYear} YEAR - ${summary.rangeLabel.toUpperCase()} ATTENDANCE ANALYSIS`, margin, 720, 11, true);
  text(firstPage, `Report Date: ${summary.selectedDate}`, margin, 702, 10);
  text(firstPage, `Section: ${summary.selectedSection}`, margin, 686, 10);
  text(firstPage, "Generated by: A.N. Harshith Vardhan", margin, 670, 9);
  text(firstPage, "Department: Computer Science and Engineering", margin, 656, 9);
  firstPage.push(`0.35 0.38 0.40 RG ${margin} 642 m 570 642 l S`);
  text(firstPage, "SUMMARY", margin, 620, 12, true);
  [["TOTAL STUDENTS", summary.totalStudents], ["MATCHING STUDENTS", summary.uniqueStudents], ["SUBJECTS ANALYZED", summary.totalSubjects], ["SUBJECT OCCURRENCES", summary.totalCases]].forEach(([label, value], index) => {
    const x = margin + index * 132;
    firstPage.push(`0.90 0.92 0.91 rg ${x} 570 124 34 re f`);
    text(firstPage, label, x + 7, 591, 7, true);
    text(firstPage, value, x + 7, 578, 13, true);
  });
  text(firstPage, "SUBJECT-WISE COUNT", margin, 546, 11, true);
  const subjectEndY = drawTable(firstPage, 526, [{ title: "Subject", width: 265, maxChars: 28 }, { title: "Students in Selected Range", width: 180, maxChars: 28 }], analysis.subjects.map((subject) => [subject.subject, String(subject.studentCount)]));
  const drawChart = (commands: string[], headingY: number) => {
    text(commands, "GRAPH - STUDENTS IN SELECTED RANGE", margin, headingY, 11, true);
    commands.push(`0.96 0.97 0.96 rg ${margin} ${headingY - 196} 528 170 re f`);
    const maxCount = Math.max(1, ...analysis.subjects.map((subject) => subject.studentCount));
    const chartStep = Math.min(19, Math.max(12, Math.floor(132 / Math.max(1, analysis.subjects.length))));
    analysis.subjects.forEach((subject, index) => {
      const y = headingY - 30 - index * chartStep;
      const width = Math.max(subject.studentCount ? 3 : 0, Math.round((subject.studentCount / maxCount) * 265));
      text(commands, subject.subject, margin + 8, y + 2, 8);
      commands.push(`0.09 0.42 0.40 rg ${margin + 105} ${y} ${width} 9 re f`);
      text(commands, String(subject.studentCount), margin + 116 + width, y + 2, 8, true);
    });
  };
  if (subjectEndY >= 360 && analysis.subjects.length <= 7) drawChart(firstPage, 334);
  pageContents.push(firstPage.join("\n"));
  if (subjectEndY < 360 || analysis.subjects.length > 7) {
    const chartPage: string[] = [];
    text(chartPage, "SUBJECT-WISE GRAPH", margin, 748, 14, true);
    text(chartPage, `${summary.selectedDate} - ${summary.selectedYear} YEAR - ${summary.rangeLabel}`, margin, 728, 9);
    drawChart(chartPage, 680);
    pageContents.push(chartPage.join("\n"));
  }

  const studentColumns = [
    { title: "S.No", width: 35, maxChars: 5 },
    { title: "Registration No", width: 105, maxChars: 15 },
    { title: "Student Name", width: 150, maxChars: 22 },
    { title: "Low Attendance Subjects", width: 238, maxChars: 35 },
  ];
  const studentRows = analysis.students.map((student) => [
    student.serialNumber,
    student.registrationNumber,
    student.name,
    student.matches.map((match) => `${match.subject} - ${match.percentage.toFixed(2)}%`).join(", "),
  ]);
  const studentPages: string[][][] = [];
  let currentRows: string[][] = [];
  let currentHeight = 0;
  studentRows.forEach((row) => {
    const wrapped = row.map((cell, index) => wrap(cell, studentColumns[index].maxChars));
    const rowHeight = Math.max(24, ...wrapped.map((lines) => lines.length * 10 + 10));
    if (currentRows.length && currentHeight + rowHeight > 620) {
      studentPages.push(currentRows);
      currentRows = [];
      currentHeight = 0;
    }
    currentRows.push(row);
    currentHeight += rowHeight;
  });
  if (currentRows.length || !studentPages.length) studentPages.push(currentRows);
  studentPages.forEach((rows) => {
    const commands: string[] = [];
    text(commands, "STUDENT-WISE ANALYSIS", margin, 748, 14, true);
    text(commands, `${summary.selectedDate} · Year ${summary.selectedYear} · Section ${summary.selectedSection}`, margin, 728, 9);
    drawTable(commands, 700, studentColumns, rows, "teal");
    pageContents.push(commands.join("\n"));
  });

  const numberedPages = pageContents.map((content, index) => `${content}\n0.45 0.46 0.45 rg\nBT /F1 8 Tf 520 28 Td (Page ${index + 1} of ${pageContents.length}) Tj ET`);
  const objects: string[] = ["<< /Type /Catalog /Pages 2 0 R >>", "", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>"];
  const pageRefs: number[] = [];
  numberedPages.forEach((content) => {
    const contentIndex = objects.push(`<< /Length ${encoder.encode(content).length} >>\nstream\n${content}\nendstream`);
    const pageIndex = objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${contentIndex} 0 R >>`);
    pageRefs.push(pageIndex);
  });
  objects[1] = `<< /Type /Pages /Kids [${pageRefs.map((reference) => `${reference} 0 R`).join(" ")}] /Count ${pageRefs.length} >>`;
  const chunks = ["%PDF-1.4\n"];
  const offsets = [0];
  let byteLength = encoder.encode(chunks[0]).length;
  objects.forEach((object, index) => {
    offsets[index + 1] = byteLength;
    const chunk = `${index + 1} 0 obj\n${object}\nendobj\n`;
    chunks.push(chunk);
    byteLength += encoder.encode(chunk).length;
  });
  const xrefOffset = byteLength;
  let xref = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => { xref += `${String(offset).padStart(10, "0")} 00000 n \n`; });
  xref += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  chunks.push(xref);
  return new Blob(chunks, { type: "application/pdf" });
}

function exportPdfAnalysis(analysis: Analysis): void {
  const summary = analysis.summary;
  const safe = (value: unknown) => safeText(value).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "-") || "report";
  const url = URL.createObjectURL(buildPdfBlob(analysis));
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safe(summary.selectedDateIso)}-${safe(summary.selectedYear)}-${safe(summary.selectedSection)}.pdf`;
  link.click();
  URL.revokeObjectURL(url);
}

function AppShell({ children, hasAnalysis }: { children: ReactNode; hasAnalysis: boolean }) {
  const [location] = useLocation();
  const [mobileOpen, setMobileOpen] = useState(false);
  const nav = [
    { href: "/", label: "Overview", icon: LayoutDashboard },
    { href: "/analysis", label: "Analysis workspace", icon: SlidersHorizontal },
    { href: "/reports", label: "Reports", icon: Files },
  ];
  return <div className="app-shell md:flex">
    <aside className="desktop-sidebar sidebar sidebar-grid sticky top-0 hidden h-[100dvh] w-[254px] shrink-0 flex-col px-4 py-5 md:flex">
      <div className="mb-10 flex items-center gap-3 px-2">
        <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-[hsl(var(--accent))] text-sm font-bold text-[hsl(var(--accent-foreground))]">AH</div>
        <div><p className="text-sm font-bold tracking-tight">Attendance Atlas</p><p className="mono mt-0.5 text-[10px] text-[hsl(var(--sidebar-foreground)/.5)]">CSE / LOCAL WORKSPACE</p></div>
      </div>
      <p className="mono mb-3 px-3 text-[10px] uppercase tracking-[.16em] text-[hsl(var(--sidebar-foreground)/.4)]">Workspace</p>
      <nav className="space-y-1">
        {nav.map(({ href, label, icon: Icon }) => <Link key={href} href={href} onClick={() => setMobileOpen(false)} data-testid={`link-nav-${label.toLowerCase().replace(/\s+/g, "-")}`} className={`sidebar-link flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm ${location === href ? "active" : ""}`}><Icon size={17} strokeWidth={1.8} /><span>{label}</span>{href === "/reports" && hasAnalysis ? <span className="ml-auto h-2 w-2 rounded-full bg-[hsl(var(--accent))]" /> : null}</Link>)}
      </nav>
      <div className="mt-auto rounded-xl border border-[hsl(var(--sidebar-foreground)/.12)] bg-[hsl(var(--sidebar-foreground)/.05)] p-3.5">
        <ShieldCheck size={17} className="mb-3 text-[hsl(var(--accent))]" />
        <p className="text-xs font-semibold">Private by design</p>
        <p className="mt-1 text-[11px] leading-relaxed text-[hsl(var(--sidebar-foreground)/.58)]">Workbooks are parsed in this browser and never uploaded.</p>
      </div>
      <p className="mono mt-5 px-2 text-[10px] text-[hsl(var(--sidebar-foreground)/.35)]">A.N. Harshith Vardhan · 2024</p>
    </aside>
    <div className="main-shell flex min-h-[100dvh] flex-1 flex-col">
      <div className="mobile-topbar sticky top-0 z-20 items-center justify-between border-b border-[hsl(var(--border))] bg-[hsl(var(--background)/.94)] px-4 py-3 backdrop-blur md:hidden">
        <button type="button" onClick={() => setMobileOpen(!mobileOpen)} className="action-btn action-ghost !p-2" data-testid="button-toggle-mobile-navigation" aria-label="Toggle navigation">{mobileOpen ? <X size={18} /> : <Menu size={18} />}</button>
        <p className="text-sm font-bold">Attendance Atlas</p><div className="h-8 w-8 rounded-lg bg-[hsl(var(--accent))] text-center text-xs font-bold leading-8">AH</div>
      </div>
      {mobileOpen ? <div className="mobile-topbar fixed inset-x-0 top-[57px] z-10 flex flex-col gap-1 border-b border-[hsl(var(--border))] bg-[hsl(var(--sidebar))] p-3 md:hidden">{nav.map(({ href, label, icon: Icon }) => <Link key={href} href={href} onClick={() => setMobileOpen(false)} className={`sidebar-link flex items-center gap-3 rounded-lg px-3 py-3 text-sm ${location === href ? "active" : ""}`} data-testid={`mobile-link-${label.toLowerCase().replace(/\s+/g, "-")}`}><Icon size={17} /><span>{label}</span></Link>)}</div> : null}
      {children}
    </div>
  </div>;
}

function PageHeader({ eyebrow, title, description, action }: { eyebrow: string; title: string; description: string; action?: ReactNode }) {
  return <header className="mb-8 flex flex-col justify-between gap-5 lg:flex-row lg:items-end">
    <div><p className="eyebrow mb-3">{eyebrow}</p><h1 className="max-w-2xl text-3xl font-bold tracking-[-.045em] md:text-4xl">{title}</h1><p className="mt-2 max-w-2xl text-sm leading-relaxed text-[hsl(var(--muted-foreground))]">{description}</p></div>
    {action}
  </header>;
}

function EmptyState({ icon: Icon, title, description, action }: { icon: typeof BookOpen; title: string; description: string; action?: ReactNode }) {
  return <div className="panel-flat flex min-h-[300px] flex-col items-center justify-center px-6 py-12 text-center"><div className="mb-5 flex h-14 w-14 items-center justify-center rounded-2xl bg-[hsl(var(--primary)/.1)] text-[hsl(var(--primary))]"><Icon size={25} strokeWidth={1.6} /></div><h2 className="text-lg font-bold tracking-tight">{title}</h2><p className="mt-2 max-w-sm text-sm leading-relaxed text-[hsl(var(--muted-foreground))]">{description}</p>{action ? <div className="mt-6">{action}</div> : null}</div>;
}

function Dashboard({ state }: { state: AppState }) {
  const analysis = state.analysis;
  return <main className="page-pad flex-1">
    <PageHeader eyebrow="Overview / local analysis" title="Know who needs a closer look." description="A focused attendance workspace for A.N. Harshith Vardhan, Computer Science and Engineering." action={<Link href="/analysis" data-testid="link-start-analysis" className="action-btn action-primary"><UploadCloud size={16} />{analysis ? "Open workspace" : "Start an analysis"}</Link>} />
    {!analysis ? <div className="space-y-5">
      <section className="panel soft-grid relative overflow-hidden p-6 md:p-8"><div className="relative max-w-2xl"><p className="eyebrow mb-4">Your next useful action</p><h2 className="max-w-xl text-2xl font-bold leading-tight tracking-[-.04em] md:text-3xl">Turn one attendance file into a clear, defensible picture.</h2><p className="mt-4 max-w-lg text-sm leading-relaxed text-[hsl(var(--muted-foreground))]">Upload an Excel or PDF attendance file, choose the report date and threshold, then review student-level cases before exporting.</p><Link href="/analysis" data-testid="link-upload-workbook-empty" className="action-btn action-primary mt-6"><FileSpreadsheet size={16} />Upload attendance file</Link></div><div className="absolute -right-12 -top-16 hidden h-64 w-64 rounded-full border-[26px] border-[hsl(var(--accent)/.2)] md:block" /><div className="absolute -bottom-28 right-14 hidden h-64 w-64 rounded-full border-[1px] border-[hsl(var(--primary)/.22)] md:block" /></section>
      <div className="grid gap-5 md:grid-cols-3"><InfoCard number="01" icon={<FileSpreadsheet size={18} />} title="Upload locally" text="Drop an .xlsx, .xls, or .pdf file. Nothing leaves this browser." /><InfoCard number="02" icon={<Filter size={18} />} title="Set the lens" text="Choose Year, Section, report date, source page, and one range." /><InfoCard number="03" icon={<Download size={18} />} title="Export with confidence" text="Download the same analyzed information as Excel or PDF." /></div>
    </div> : <div className="space-y-5">
      <section className="panel overflow-hidden p-5 md:p-7"><div className="flex flex-col justify-between gap-5 sm:flex-row sm:items-start"><div><p className="eyebrow mb-2">Latest local report</p><h2 className="text-2xl font-bold tracking-tight">{analysis.summary.selectedDate}</h2><p className="mt-1 text-sm text-[hsl(var(--muted-foreground))]">Year {analysis.summary.selectedYear} <span className="mx-2 text-[hsl(var(--border))]">/</span> Section {analysis.summary.selectedSection} <span className="mx-2 text-[hsl(var(--border))]">/</span> {analysis.summary.rangeLabel}</p></div><div className="flex gap-2"><Link href="/reports" data-testid="link-review-report" className="action-btn action-primary"><Files size={16} />Review report</Link><Link href="/analysis" data-testid="link-new-analysis-dashboard" className="action-btn action-ghost"><RefreshCw size={16} />New analysis</Link></div></div><div className="mt-8 grid grid-cols-2 gap-3 lg:grid-cols-4"><MetricCard label="Students scanned" value={analysis.summary.totalStudents} detail="unique registrations" /><MetricCard label="At-risk students" value={analysis.summary.uniqueStudents} detail="with one or more matches" accent /><MetricCard label="Subject cases" value={analysis.summary.totalCases} detail="student × subject matches" /><MetricCard label="Subjects detected" value={analysis.summary.totalSubjects} detail="numeric columns" /></div></section>
      <div className="grid gap-5 lg:grid-cols-[1.1fr_.9fr]"><section className="panel-flat p-5"><SectionHeading title="Most affected subjects" aside="students with a match" />{analysis.subjects.length ? <div className="mt-4 space-y-3">{analysis.subjects.slice(0, 5).map((subject, index) => <div key={subject.subject} data-testid={`dashboard-subject-${index}`}><div className="mb-1.5 flex justify-between text-xs"><span className="font-semibold">{subject.subject}</span><span className="mono text-[hsl(var(--muted-foreground))]">{subject.studentCount}</span></div><div className="h-2 overflow-hidden rounded-full bg-[hsl(var(--muted))]"><div className="h-full rounded-full bg-[hsl(var(--primary))]" style={{ width: `${Math.max(6, (subject.studentCount / Math.max(1, analysis.summary.uniqueStudents)) * 100)}%` }} /></div></div>)}</div> : <p className="mt-5 text-sm text-[hsl(var(--muted-foreground))]">No subject matches in the selected range.</p>}</section><section className="panel-flat p-5"><SectionHeading title="Analysis notes" /><div className="mt-4 space-y-3 text-sm"><NoteLine label="Source file" value={state.workbook?.fileName ?? "Local file"} /><NoteLine label="Report date" value={analysis.summary.selectedDate} /><NoteLine label={state.workbook?.fileType === "pdf" ? "Source document" : "Selected sheet"} value={analysis.summary.sourceSheet} /><NoteLine label="Range logic" value={analysis.summary.rangeLabel} /><div className="mt-5 flex items-start gap-2 rounded-lg bg-[hsl(var(--primary)/.07)] p-3 text-xs leading-relaxed text-[hsl(var(--primary))]"><ShieldCheck size={15} className="mt-0.5 shrink-0" />No source file leaves this browser. Reports are generated only when you choose download.</div></div></section></div>
    </div>}
  </main>;
}

function InfoCard({ number, icon, title, text }: { number: string; icon: ReactNode; title: string; text: string }) {
  return <div className="panel-flat p-5"><div className="flex items-center justify-between text-[hsl(var(--primary))]"><span>{icon}</span><span className="mono text-[10px] text-[hsl(var(--muted-foreground))]">{number}</span></div><h3 className="mt-8 font-bold">{title}</h3><p className="mt-1.5 text-sm leading-relaxed text-[hsl(var(--muted-foreground))]">{text}</p></div>;
}

function MetricCard({ label, value, detail, accent = false }: { label: string; value: number; detail: string; accent?: boolean }) {
  return <div className={`rounded-xl border p-4 ${accent ? "border-[hsl(var(--accent)/.35)] bg-[hsl(var(--accent)/.09)]" : "border-[hsl(var(--border))] bg-[hsl(var(--background)/.55)]"}`}><p className="text-xs font-semibold text-[hsl(var(--muted-foreground))]">{label}</p><p className={`stat-number mt-3 ${accent ? "text-[hsl(var(--accent-foreground))]" : ""}`} data-testid={`metric-${label.toLowerCase().replace(/\s+/g, "-")}`}>{value}</p><p className="mt-2 text-[11px] text-[hsl(var(--muted-foreground))]">{detail}</p></div>;
}

function SectionHeading({ title, aside }: { title: string; aside?: string }) {
  return <div className="flex items-center justify-between gap-3"><div><div className="kicker-line mb-3" /><h2 className="text-base font-bold tracking-tight">{title}</h2></div>{aside ? <span className="mono text-[10px] text-[hsl(var(--muted-foreground))]">{aside}</span> : null}</div>;
}

function NoteLine({ label, value }: { label: string; value: string }) {
  return <div className="flex items-center justify-between gap-4 border-b border-[hsl(var(--border)/.7)] pb-3 text-xs last:border-0"><span className="text-[hsl(var(--muted-foreground))]">{label}</span><span className="max-w-[62%] truncate text-right font-semibold">{value}</span></div>;
}

function UploadZone({ onFile, parsing, error }: { onFile: (file: File) => void; parsing: boolean; error: string }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const handleDrop = (event: DragEvent<HTMLDivElement>) => { event.preventDefault(); setDragging(false); const file = event.dataTransfer.files?.[0]; if (file) onFile(file); };
  const handleInput = (event: ChangeEvent<HTMLInputElement>) => { const file = event.target.files?.[0]; if (file) onFile(file); };
  return <div className="space-y-3"><div className={`drop-zone flex min-h-[210px] cursor-pointer flex-col items-center justify-center rounded-xl px-6 py-9 text-center ${dragging ? "dragging" : ""}`} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={handleDrop} onClick={() => inputRef.current?.click()} data-testid="dropzone-workbook"><input ref={inputRef} type="file" accept=".xlsx,.xls,.pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/pdf" className="hidden" onChange={handleInput} data-testid="input-workbook" /><div className="mb-4 flex h-12 w-12 items-center justify-center rounded-xl bg-[hsl(var(--primary)/.11)] text-[hsl(var(--primary))]">{parsing ? <Loader2 className="animate-spin" size={23} /> : <UploadCloud size={23} />}</div><p className="text-sm font-bold">{parsing ? "Reading attendance file locally…" : "Drop your Excel or PDF attendance file here"}</p><p className="mt-1 text-xs text-[hsl(var(--muted-foreground))]">or click to browse · .xlsx, .xls, or .pdf</p></div>{error ? <div className="error-box flex items-start gap-2 rounded-lg p-3 text-xs" data-testid="status-upload-error"><AlertTriangle size={15} className="mt-0.5 shrink-0" />{error}</div> : null}</div>;
}

function AnalysisPage({ state, setState }: { state: AppState; setState: (next: (current: AppState) => AppState) => void }) {
  const replaceInputRef = useRef<HTMLInputElement>(null);
  const updateConfig = (updates: Partial<AnalysisConfig>) => setState((current) => ({
    ...current,
    config: { ...current.config, ...updates },
    workbook: current.workbook && updates.sheetName ? { ...current.workbook, activeSheet: updates.sheetName } : current.workbook,
    analysis: null,
  }));
  const onFile = async (file: File) => {
    setState((current) => ({ ...current, parsing: true, parseError: "", analysis: null }));
    try {
      const parsed = await readAttendanceFile(file);
      setState((current) => ({ ...current, ...parsed, config: { ...current.config, sheetName: parsed.workbook.activeSheet }, parsing: false, parseError: "" }));
    } catch (error) {
      setState((current) => ({ ...current, parsing: false, parseError: error instanceof Error ? error.message : "Could not read this attendance file." }));
    }
  };
  const runAnalysis = () => {
    if (!state.workbook) return;
    if (!state.config.year.trim()) { setState((current) => ({ ...current, parseError: "Select a Year before analyzing." })); return; }
    if (!state.config.section.trim()) { setState((current) => ({ ...current, parseError: "Enter a Section before analyzing." })); return; }
    if (!state.config.reportDate) { setState((current) => ({ ...current, parseError: "Select a Report date before analyzing." })); return; }
    if (!state.config.sheetName) { setState((current) => ({ ...current, parseError: "Choose a sheet or date to analyze." })); return; }
    const analysis = buildAnalysis(state);
    setState((current) => ({ ...current, analysis, parseError: analysis ? "" : "Complete the required fields before analyzing." }));
  };
  return <main className="page-pad flex-1">
    <PageHeader eyebrow="Analysis workspace" title="Set the lens, then inspect the evidence." description="Everything below is computed from the workbook in this browser. No attendance data leaves this session." action={state.workbook ? <button type="button" className="action-btn action-ghost" onClick={() => replaceInputRef.current?.click()} data-testid="button-replace-workbook"><RefreshCw size={16} />Replace workbook</button> : null} />
    <input ref={replaceInputRef} type="file" accept=".xlsx,.xls,.pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/pdf" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) onFile(file); event.target.value = ""; }} data-testid="input-replace-workbook" />
    {!state.workbook ? <section className="mx-auto max-w-3xl"><div className="panel p-5 md:p-7"><SectionHeading title="Upload attendance file" aside="step 01 / 03" /><p className="mb-5 mt-4 text-sm text-[hsl(var(--muted-foreground))]">Use an Excel or PDF attendance file. The detector finds registration numbers, student names, and numeric subject columns without relying on fixed positions.</p><UploadZone onFile={onFile} parsing={state.parsing} error={state.parseError} /></div><div className="mt-4 flex items-start gap-2 px-1 text-xs text-[hsl(var(--muted-foreground))]"><CircleHelp size={14} className="mt-0.5 shrink-0 text-[hsl(var(--primary))]" />Total, Counselling, blank sheets, and ambiguous columns are reported as warnings or excluded.</div></section> : <div className="space-y-5">
      <section className="grid gap-5 xl:grid-cols-[1.02fr_.98fr]"><div className="panel p-5 md:p-7"><SectionHeading title="Configure analysis" aside="steps 02–03" /><div className="mt-6 grid gap-4 sm:grid-cols-2"><label className="block"><span className="mb-2 block text-xs font-semibold">Year <span className="text-[hsl(var(--accent))]">*</span></span><select className="field" value={state.config.year} onChange={(event) => updateConfig({ year: event.target.value })} data-testid="select-year"><option value="" disabled>Select Year</option><option value="I">I</option><option value="II">II</option><option value="III">III</option><option value="IV">IV</option></select></label><label className="block"><span className="mb-2 block text-xs font-semibold">Section <span className="text-[hsl(var(--accent))]">*</span></span><input className="field" value={state.config.section} onChange={(event) => updateConfig({ section: event.target.value })} placeholder="e.g. 11, 19, A, B" data-testid="input-section" /></label><label className="block"><span className="mb-2 block text-xs font-semibold">Report date <span className="text-[hsl(var(--accent))]">*</span></span><div className="relative"><input className="field pr-10" type="date" value={state.config.reportDate} onChange={(event) => updateConfig({ reportDate: event.target.value })} data-testid="input-report-date" /><CalendarDays size={16} className="pointer-events-none absolute right-3 top-3 text-[hsl(var(--muted-foreground))]" /></div></label>{state.workbook.fileType === "excel" && state.workbook.sheets.length > 1 ? <label className="block"><span className="mb-2 block text-xs font-semibold">Select Sheet <span className="text-[hsl(var(--accent))]">*</span></span><div className="relative"><select className="field appearance-none pr-10" value={state.config.sheetName} onChange={(event) => updateConfig({ sheetName: event.target.value })} data-testid="select-sheet">{state.workbook.sheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}</select><ChevronDown className="pointer-events-none absolute right-3 top-3 text-[hsl(var(--muted-foreground))]" size={16} /></div></label> : null}</div><div className="mt-6"><p className="mb-2 text-xs font-semibold">Attendance range</p><div className="grid gap-3 sm:grid-cols-2">{(["below65", "65to74"] as AttendanceRange[]).map((range) => <label key={range} className={`range-card flex cursor-pointer items-start gap-3 rounded-xl p-3.5 ${state.config.range === range ? "selected" : ""}`} data-testid={`range-option-${range}`}><input type="radio" name="range" checked={state.config.range === range} onChange={() => updateConfig({ range })} className="mt-1" /><span><span className="block text-sm font-bold">{rangeLabels[range]}</span><span className="mt-1 block text-xs leading-relaxed text-[hsl(var(--muted-foreground))]">{range === "below65" ? "Attendance strictly below 65%." : "65% inclusive, 75% exclusive."}</span></span>{state.config.range === range ? <Check className="ml-auto shrink-0 text-[hsl(var(--primary))]" size={17} /> : null}</label>)}</div></div>{state.parseError ? <div className="error-box mt-5 flex items-start gap-2 rounded-lg p-3 text-xs" data-testid="status-analysis-error"><AlertTriangle size={15} className="mt-0.5 shrink-0" />{state.parseError}</div> : null}<button type="button" onClick={runAnalysis} className="action-btn action-primary mt-6 w-full sm:w-auto" data-testid="button-run-analysis"><BarChart3 size={16} />Analyze {state.workbook.fileType === "pdf" ? "PDF" : "selected sheet"}</button></div>
          <div className="panel-flat p-5 md:p-7"><SectionHeading title="File map" aside="detected locally" /><div className="mt-5 flex items-center gap-3 rounded-xl bg-[hsl(var(--muted)/.55)] p-3.5"><div className="flex h-10 w-10 items-center justify-center rounded-lg bg-[hsl(var(--card))] text-[hsl(var(--primary))]">{state.workbook.fileType === "pdf" ? <FileText size={19} /> : <FileSpreadsheet size={19} />}</div><div className="min-w-0"><p className="truncate text-sm font-bold" data-testid="text-workbook-name">{state.workbook.fileName}</p><p className="mono mt-1 text-[10px] text-[hsl(var(--muted-foreground))]">{state.workbook.fileType === "pdf" ? "PDF · complete document" : `EXCEL · ${state.workbook.sheets.length} ${state.workbook.sheets.length === 1 ? "sheet" : "sheets"} detected`}</p></div></div>{state.workbook.fileType === "excel" && state.workbook.sheets.length > 1 ? <div className="mt-6"><p className="eyebrow mb-3">Select one sheet to analyze</p><div className="flex max-h-[190px] flex-wrap content-start gap-2 overflow-auto pr-1">{state.workbook.sheets.map((sheet) => <button type="button" key={sheet} onClick={() => updateConfig({ sheetName: sheet })} className={`rounded-lg border px-3 py-2 text-left text-xs font-semibold transition-colors ${state.config.sheetName === sheet ? "border-[hsl(var(--primary))] bg-[hsl(var(--primary)/.1)] text-[hsl(var(--primary))]" : "border-[hsl(var(--border))] bg-[hsl(var(--card))] text-[hsl(var(--muted-foreground))] hover:border-[hsl(var(--primary)/.4)]"}`} data-testid={`button-sheet-${sheet.replace(/\s+/g, "-")}`}>{sheet}</button>)}</div></div> : <div className="mt-6 rounded-lg bg-[hsl(var(--muted)/.55)] p-3 text-xs leading-relaxed text-[hsl(var(--muted-foreground))]">{state.workbook.fileType === "pdf" ? "All PDF pages have been combined into one attendance document." : "This workbook has one sheet, which was selected automatically."}</div>}<div className="mt-7 border-t border-[hsl(var(--border))] pt-5"><p className="text-xs font-semibold">Detection guardrails</p><ul className="mt-3 space-y-2 text-xs leading-relaxed text-[hsl(var(--muted-foreground))]"><li className="flex gap-2"><Check size={14} className="mt-0.5 shrink-0 text-[hsl(var(--primary))]" />Ratios, percentages, blanks, and invalid values are normalized safely.</li><li className="flex gap-2"><Check size={14} className="mt-0.5 shrink-0 text-[hsl(var(--primary))]" />Total, Counselling, and non-subject metadata are always excluded.</li><li className="flex gap-2"><Check size={14} className="mt-0.5 shrink-0 text-[hsl(var(--primary))]" />Duplicate registration numbers are surfaced before review.</li></ul></div></div></section>
      {state.warnings.length ? <div className="warning-box flex items-start gap-3 rounded-xl p-4 text-xs leading-relaxed" data-testid="status-workbook-warnings"><AlertTriangle size={17} className="mt-0.5 shrink-0" /><div><p className="font-bold">Review workbook notes</p><ul className="mt-1.5 list-disc space-y-1 pl-4">{state.warnings.slice(0, 7).map((warning, index) => <li key={`${warning}-${index}`}>{warning}</li>)}</ul>{state.warnings.length > 7 ? <p className="mt-1 font-semibold">+ {state.warnings.length - 7} more notes</p> : null}</div></div> : null}
      {state.analysis ? <Results analysis={state.analysis} workbookName={state.workbook.fileName} /> : <div className="panel-flat flex items-center gap-4 p-5 text-sm text-[hsl(var(--muted-foreground))]"><div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-[hsl(var(--muted))]"><ArrowDownUp size={18} /></div><div><p className="font-semibold text-[hsl(var(--foreground))]">Results will appear here</p><p className="mt-1 text-xs">Choose your context and range, then run the analysis.</p></div></div>}
    </div>}
  </main>;
}

function Results({ analysis, workbookName }: { analysis: Analysis; workbookName: string }) {
  const [query, setQuery] = useState("");
  const [sort, setSort] = useState<"count" | "registration" | "name">("count");
  const [tab, setTab] = useState<"students" | "subjects">("students");
  const filteredStudents = useMemo(() => analysis.students.filter((student) => `${student.registrationNumber} ${student.name}`.toLowerCase().includes(query.toLowerCase())).sort((a, b) => sort === "count" ? b.matchCount - a.matchCount || a.name.localeCompare(b.name) : sort === "name" ? a.name.localeCompare(b.name) : a.registrationNumber.localeCompare(b.registrationNumber)), [analysis.students, query, sort]);
  const chartData = analysis.subjects.slice(0, 8);
  return <section className="space-y-5 rise-in" data-testid="section-analysis-results"><div className="grid grid-cols-2 gap-3 lg:grid-cols-4"><MetricCard label="Students scanned" value={analysis.summary.totalStudents} detail={workbookName} /><MetricCard label="At-risk students" value={analysis.summary.uniqueStudents} detail={analysis.summary.rangeLabel} accent /><MetricCard label="Subject cases" value={analysis.summary.totalCases} detail="all matching values" /><MetricCard label="Subjects detected" value={analysis.summary.totalSubjects} detail="numeric columns" /></div><div className="grid gap-5 lg:grid-cols-[1.25fr_.75fr]"><div className="panel-flat min-h-[310px] p-5"><SectionHeading title="Cases by subject" aside="top 8 subjects" />{chartData.length ? <div className="mt-5 h-[235px] w-full"><ResponsiveContainer width="100%" height="100%"><BarChart data={chartData} margin={{ top: 5, right: 6, left: -18, bottom: 42 }}><CartesianGrid stroke="hsl(37 24% 86%)" vertical={false} /><XAxis dataKey="subject" angle={-35} textAnchor="end" interval={0} tick={{ fontSize: 10, fill: "#6d7782" }} /><YAxis allowDecimals={false} tick={{ fontSize: 10, fill: "#6d7782" }} /><ChartTooltip cursor={{ fill: "rgba(23,107,103,.07)" }} contentStyle={{ borderRadius: 8, border: "1px solid #e2ded5", background: "#fbfaf6", fontSize: 12 }} /><Bar dataKey="studentCount" name="Students" radius={[5, 5, 0, 0]} fill="#176b67" /></BarChart></ResponsiveContainer></div> : <ChartEmpty text="No matches for this range." />}</div><div className="panel-flat min-h-[310px] p-5"><SectionHeading title="Share of cases" aside="subject distribution" />{chartData.length ? <div className="relative mt-2 h-[245px]"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={chartData} dataKey="studentCount" nameKey="subject" innerRadius={57} outerRadius={88} paddingAngle={3} stroke="none">{chartData.map((entry, index) => <Cell key={entry.subject} fill={chartColors[index % chartColors.length]} />)}</Pie><ChartTooltip contentStyle={{ borderRadius: 8, border: "1px solid #e2ded5", background: "#fbfaf6", fontSize: 12 }} /></PieChart></ResponsiveContainer><div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center"><span className="stat-number text-xl">{analysis.summary.totalCases}</span><span className="mono mt-1 text-[9px] text-[hsl(var(--muted-foreground))]">CASES</span></div></div> : <ChartEmpty text="No matches for this range." />}</div></div><div className="panel-flat overflow-hidden"><div className="flex flex-col gap-4 border-b border-[hsl(var(--border))] p-5 md:flex-row md:items-center md:justify-between"><div><SectionHeading title="Review matches" aside={`${filteredStudents.length} shown`} /><div className="mt-4 flex gap-1 rounded-lg bg-[hsl(var(--muted)/.65)] p-1"><button type="button" className={`rounded-md px-3 py-1.5 text-xs font-semibold ${tab === "students" ? "bg-[hsl(var(--card))] shadow-sm" : "text-[hsl(var(--muted-foreground))]"}`} onClick={() => setTab("students")} data-testid="button-tab-students">Student-wise</button><button type="button" className={`rounded-md px-3 py-1.5 text-xs font-semibold ${tab === "subjects" ? "bg-[hsl(var(--card))] shadow-sm" : "text-[hsl(var(--muted-foreground))]"}`} onClick={() => setTab("subjects")} data-testid="button-tab-subjects">Subject-wise</button></div></div>{tab === "students" ? <div className="flex flex-col gap-2 sm:flex-row"><label className="relative"><Search size={15} className="absolute left-3 top-2.5 text-[hsl(var(--muted-foreground))]" /><input className="field !w-full !py-2 !pl-9 sm:w-[210px]" type="search" placeholder="Search student…" value={query} onChange={(event) => setQuery(event.target.value)} data-testid="input-search-students" /></label><div className="relative"><select className="field !w-full appearance-none !py-2 pr-8 sm:w-[160px]" value={sort} onChange={(event) => setSort(event.target.value as typeof sort)} data-testid="select-sort-students"><option value="count">Sort: match count</option><option value="registration">Sort: registration</option><option value="name">Sort: name</option></select><ChevronDown className="pointer-events-none absolute right-2.5 top-2.5 text-[hsl(var(--muted-foreground))]" size={15} /></div></div> : <span className="mono text-[10px] text-[hsl(var(--muted-foreground))]">ranked by student count</span>}</div>{tab === "students" ? <StudentTable students={filteredStudents} /> : <SubjectTable subjects={analysis.subjects} />}</div></section>;
}

function ChartEmpty({ text }: { text: string }) { return <div className="flex h-[245px] items-center justify-center text-xs text-[hsl(var(--muted-foreground))]">{text}</div>; }
function StudentTable({ students }: { students: StudentResult[] }) {
  return students.length ? <div className="table-wrap"><table className="data-table" data-testid="table-student-results"><thead><tr><th>S.No</th><th>Registration No</th><th>Student Name</th><th>Low Attendance Subjects</th><th className="text-right">Count</th></tr></thead><tbody>{students.map((student, index) => <tr key={`${student.registrationNumber}-${index}`} data-testid={`row-student-${index}`}><td className="mono text-xs text-[hsl(var(--muted-foreground))]">{student.serialNumber}</td><td className="mono whitespace-nowrap text-xs text-[hsl(var(--primary))]">{student.registrationNumber}</td><td className="min-w-[150px] max-w-[220px] whitespace-normal font-semibold">{student.name}</td><td><div className="flex max-w-[520px] flex-wrap gap-1.5">{student.matches.map((match) => <span key={match.subject} className="inline-flex items-center gap-1 rounded-md bg-[hsl(var(--accent)/.13)] px-2 py-1 text-[11px] font-medium text-[hsl(var(--accent-foreground))]">{match.subject} <span className="mono text-[10px] opacity-70">{match.percentage.toFixed(2)}%</span></span>)}</div></td><td className="text-right"><span className="inline-flex min-w-7 items-center justify-center rounded-md bg-[hsl(var(--primary)/.1)] px-2 py-1 mono text-xs font-bold text-[hsl(var(--primary))]">{student.matchCount}</span></td></tr>)}</tbody></table></div> : <div className="flex min-h-[190px] items-center justify-center px-5 text-center text-sm text-[hsl(var(--muted-foreground))]">No students match that search.</div>;
}
function SubjectTable({ subjects }: { subjects: SubjectCount[] }) {
  return subjects.length ? <div className="table-wrap"><table className="data-table" data-testid="table-subject-results"><thead><tr><th>Rank</th><th>Subject</th><th className="text-right">Students with matches</th><th className="w-1/3">Relative volume</th></tr></thead><tbody>{subjects.map((subject, index) => <tr key={subject.subject}><td className="mono text-[hsl(var(--muted-foreground))]">{String(index + 1).padStart(2, "0")}</td><td className="font-semibold">{subject.subject}</td><td className="text-right font-bold">{subject.studentCount}</td><td><div className="h-2 rounded-full bg-[hsl(var(--muted))]"><div className="h-full rounded-full bg-[hsl(var(--primary))]" style={{ width: `${Math.max(5, subject.studentCount / Math.max(...subjects.map((entry) => entry.studentCount)) * 100)}%` }} /></div></td></tr>)}</tbody></table></div> : <div className="flex min-h-[190px] items-center justify-center text-sm text-[hsl(var(--muted-foreground))]">No subject cases to show.</div>;
}

function ReportsPage({ state }: { state: AppState }) {
  const analysis = state.analysis;
  return <main className="page-pad flex-1"><PageHeader eyebrow="Reports / export" title="A clean handoff, when you are ready." description="Review the exact scope of the current analysis, then download the same report as Excel or PDF." action={analysis ? <div className="flex flex-wrap gap-2"><button type="button" className="action-btn action-primary" onClick={() => exportAnalysis(analysis)} data-testid="button-download-report"><FileSpreadsheet size={16} />Download Excel</button><button type="button" className="action-btn action-ghost" onClick={() => exportPdfAnalysis(analysis)} data-testid="button-download-pdf-report"><FileDown size={16} />Download PDF</button></div> : null} />{!analysis ? <EmptyState icon={Files} title="No report in this session" description="Run an analysis first. Your export will be built from the selected source and range, without uploading the original file." action={<Link href="/analysis" className="action-btn action-primary" data-testid="link-go-analysis-report"><BarChart3 size={16} />Go to analysis</Link>} /> : <div className="space-y-5 rise-in"><section className="panel overflow-hidden p-5 md:p-7"><div className="flex flex-col justify-between gap-5 border-b border-[hsl(var(--border))] pb-6 sm:flex-row sm:items-start"><div><p className="eyebrow mb-3">Current analysis</p><h2 className="text-2xl font-bold tracking-tight">{analysis.summary.selectedDate}</h2><p className="mt-2 text-sm text-[hsl(var(--muted-foreground))]">Year {analysis.summary.selectedYear} <span className="mx-2">·</span> Section {analysis.summary.selectedSection} <span className="mx-2">·</span> {analysis.summary.sourceSheet}</p></div><div className="flex items-center gap-2 rounded-lg bg-[hsl(var(--primary)/.08)] px-3 py-2 text-xs font-semibold text-[hsl(var(--primary))]"><Check size={15} />Ready to download</div></div><div className="mt-6 grid gap-3 sm:grid-cols-2 lg:grid-cols-4"><ReportStat label="Scope" value={analysis.summary.rangeLabel} /><ReportStat label="Students" value={String(analysis.summary.uniqueStudents)} /><ReportStat label="Subject cases" value={String(analysis.summary.totalCases)} /><ReportStat label="Detected subjects" value={String(analysis.summary.totalSubjects)} /></div></section><section className="panel-flat p-5 md:p-7"><SectionHeading title="Report contents" aside="same data in two formats" /><div className="mt-5 grid gap-3 md:grid-cols-2">{[["01", "Summary", "Report date, Year, Section, range, and headline counts."], ["02", "Student-wise Analysis", "One row per student with matched subjects."], ["03", "Subject-wise Analysis", "Subjects ranked by affected students."], ["04", "Detailed Analysis", "One row per student × subject case."]].map(([number, title, text]) => <div key={title} className="flex items-start gap-3 rounded-xl border border-[hsl(var(--border))] p-4"><span className="mono text-[10px] text-[hsl(var(--accent-foreground))]">{number}</span><div><p className="text-sm font-bold">{title}</p><p className="mt-1 text-xs leading-relaxed text-[hsl(var(--muted-foreground))]">{text}</p></div></div>)}</div><div className="mt-6 flex flex-wrap gap-2"><button type="button" onClick={() => exportAnalysis(analysis)} className="action-btn action-primary" data-testid="button-download-report-bottom"><FileSpreadsheet size={16} />Download {analysis.summary.selectedDate}-{analysis.summary.selectedYear}-{analysis.summary.selectedSection}.xlsx</button><button type="button" onClick={() => exportPdfAnalysis(analysis)} className="action-btn action-ghost" data-testid="button-download-pdf-report-bottom"><FileDown size={16} />Download PDF</button></div></section><div className="flex items-start gap-2 px-1 text-xs leading-relaxed text-[hsl(var(--muted-foreground))]"><ShieldCheck size={15} className="mt-0.5 shrink-0 text-[hsl(var(--primary))]" />Both report formats use the selected report date and exclude Total, Counselling, and other aggregate columns.</div></div>}</main>;
}

function ReportStat({ label, value }: { label: string; value: string }) { return <div className="rounded-xl bg-[hsl(var(--muted)/.55)] p-4"><p className="text-[11px] font-semibold text-[hsl(var(--muted-foreground))]">{label}</p><p className="mt-3 text-base font-bold tracking-tight">{value}</p></div>; }

function Router({ state, setState }: { state: AppState; setState: (next: (current: AppState) => AppState) => void }) {
  return <AppShell hasAnalysis={Boolean(state.analysis)}><ErrorBoundary resetKey={window.location.pathname}><Switch><Route path="/" component={() => <Dashboard state={state} />} /><Route path="/analysis" component={() => <AnalysisPage state={state} setState={setState} />} /><Route path="/reports" component={() => <ReportsPage state={state} />} /><Route component={NotFound} /></Switch></ErrorBoundary></AppShell>;
}

function App() {
  const [state, setState] = useState<AppState>({ workbook: null, rows: {}, subjects: {}, warnings: [], analysis: null, config: emptyConfig, parsing: false, parseError: "" });
  return <QueryClientProvider client={queryClient}><TooltipProvider><WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, "")}><Router state={state} setState={setState} /></WouterRouter><Toaster /></TooltipProvider></QueryClientProvider>;
}

export default App;