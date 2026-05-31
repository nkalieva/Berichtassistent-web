# Berichtsassistent 2 — Web / JavaScript Excel Add-in

A JavaScript/TypeScript Excel Add-in (Office.js) with custom functions for automated data retrieval and report generation.

## What it does
- Custom Excel functions for live data access from a REST API
- Task pane UI for server configuration
- Automated layout and formula insertion into Excel sheets
- Supports P&L and balance sheet data retrieval

## Technologies
- **JavaScript / TypeScript** — core logic
- **Office.js / Excel API** — Excel Add-in framework
- **REST API** — async fetch calls
- **Webpack** — bundling and build

## Structure
```
src/functions/    # Custom Excel UDF functions
src/taskpane/     # Task pane UI (HTML/CSS/JS)
src/commands/     # Ribbon command handlers
```

## Background
JavaScript companion to the C# version of Berichtsassistent 2. Enables cross-platform use of the same reporting functions via the Office.js Add-in model.
