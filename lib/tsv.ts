// Tab-separated text is what Excel, Google Sheets and most apps put on the clipboard.

const needsQuotes = (s: string) => /[\t\n\r"]/.test(s);

export function toTSV(rows: string[][]): string {
  return rows
    .map((row) => row.map((v) => (needsQuotes(v) ? `"${v.replace(/"/g, '""')}"` : v)).join("\t"))
    .join("\n");
}

export function parseTSV(text: string): string[][] {
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
    if (ch === "\t") {
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
