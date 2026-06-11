---
title: FlexiMapper
emoji: 🗺️
colorFrom: indigo
colorTo: purple
sdk: docker
app_port: 7860
pinned: false
---

# FlexiMapper — Intelligent Excel Column Mapping & Merger

FlexiMapper is a lightweight utility designed to merge Excel spreadsheet data into a styled target Excel template. 

## Features
- **All Source Columns Present**: Lists all source sheet headers on the Left Hand Side (LHS).
- **Target Mapping Dropdowns**: Choose the destination column in the target sheet (RHS) for each source column.
- **Smart Match**: Click to auto-align columns based on name similarity.
- **Excel Style Preservation**: Preserves font styling, backgrounds, cell widths, and formatting using `openpyxl`.
- **Gracious Typecast Handling**: Automatically parses string dates, currencies, and percentage forms into native cell formats without throwing exceptions.
- **Secure File Cleanup**: Auto-cleans all uploaded and merged spreadsheets immediately after downloads are complete.
