---
name: xlsx-reader
description: Read and inspect local Excel .xlsx workbooks, list worksheets, and extract cell values for analysis without modifying the workbook.
---

# Reading XLSX workbooks

Use this skill when the user asks to read, inspect, search, or summarize a local `.xlsx` workbook.

1. Use only the workbook the user identifies or the relevant workbook already in the workspace. Do not upload workbook contents to external services.
2. List worksheets first:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File '.github\skills\xlsx-reader\scripts\read-xlsx.ps1' -Path 'path\to\workbook.xlsx' -ListSheets
   ```

3. Read only the relevant sheet or sheets. Use `-MaxRows` to keep large outputs manageable:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File '.github\skills\xlsx-reader\scripts\read-xlsx.ps1' -Path 'path\to\workbook.xlsx' -Sheet 'Sheet name' -MaxRows 100
   ```

   Omit `-Sheet` to read every worksheet. Omit `-MaxRows` to read all non-empty rows.

4. Report the worksheet names and any extraction limitations. The script reads stored cell values; formula cells without cached results may be empty, and numeric date/time values may appear as Excel serials. Do not modify the workbook unless explicitly asked.
