// Tab-separated text is what Excel, Google Sheets and most apps put on the clipboard;
// comma-separated is the CSV file format. Both quote fields the same way.

export function toDelimited(rows: string[][], delimiter: string): string {
  const needsQuotes = (v: string) => v.includes(delimiter) || /[\n\r"]/.test(v);
  return rows
    .map((row) => row.map((v) => (needsQuotes(v) ? `"${v.replace(/"/g, '""')}"` : v)).join(delimiter))
    .join("\n");
}

export const toTSV = (rows: string[][]) => toDelimited(rows, "\t");
export const toCSV = (rows: string[][]) => toDelimited(rows, ",");
export const parseTSV = (text: string) => parseDelimited(text, "\t");
export const parseCSV = (text: string) => parseDelimited(text.replace(/^﻿/, ""), ",");

export function parseDelimited(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let i = 0;
  const src = text.replace(/\r\n?/g, "\n");

  while (i < src.length) {
    const ch = src[i];
    if (ch === '"' && field === "") {
      // Quoted field: runs to the closing quote and may contain tabs and newlines.
      let j = i + 1;
      let closed = false;
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            field += '"';
            j += 2;
            continue;
          }
          closed = true;
          break;
        }
        field += src[j++];
      }
      if (closed) {
        i = j + 1;
        continue;
      }
      // No closing quote: treat the quote as plain text.
      field = '"';
      i++;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = "";
    } else if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += ch;
    }
    i++;
  }
  if (field !== "" || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}
