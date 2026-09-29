# Dashboard Architecture

This document describes the current production shape of the reconciliation dashboard. It is intentionally split into three diagrams so data flow, deployment, and screen behavior can be reviewed independently.

## 1. System and deployment

```mermaid
flowchart LR
  SRC[Local source code] --> GH[GitHub master]
  GH -->|auto-deploy| R[Render web service]
  SRC -->|python app.py| L[Local Flask]
  R --> APP[Flask app]
  L --> APP
  APP --> AUTH[Session login gate]
  AUTH --> ROUTES[Dashboard and protected API routes]
```

## 2. Data pipeline

```mermaid
flowchart LR
  SHEETS[(Google Sheets)] -->|read tabs| PARSE[sheets_service.py]
  PARSE -->|header mapping + normalization| MODEL[Normalized source models]
  MODEL -->|shared overview metrics| CACHE[Safe in-memory snapshot]
  CACHE --> API[/api/data]
  UI[Browser dashboard] -->|fetch| API
  UI -->|POST /api/update-row| WRITE[Google Sheets write-back]
  WRITE -->|refresh after write| CACHE
  CACHE --> HEALTH[Refresh age / partial errors / stale status]
```

The Excel-to-Google-Sheet sync remains a separate, manual upstream workflow. The dashboard does not run that sync.

## 3. Dashboard calculation and view mapping

```mermaid
flowchart TB
  ROWS[Normalized rows] --> BASE[Backend KPIs and overview metrics]
  ROWS --> BROWSER[Dashboard view model]
  BASE --> CARDS[Overview KPI cards]
  BASE --> DIFF[Final Diff / Sum Difference metrics]
  DIFF --> BUCKET[Bucket Breakdown]
  BROWSER --> YTD[Claims Bucket Year-to-Date]
  BROWSER --> CHANNEL[Claims by Channel and Month]
  BROWSER --> OPS[Operational Pivot tabs]
  BROWSER --> CHARTS[Monthly status and recovery charts]
  BROWSER --> UPS[UPS claims table]
  YTD --> MODAL[Claim detail modal]
  CHANNEL --> MODAL
  OPS --> EXPAND[Expand month to channels]
```

## Refresh and reliability behavior

- Manual Refresh calls `/api/refresh` and then reloads `/api/data`.
- The browser refreshes automatically every 20 minutes.
- The backend keeps the last successful snapshot if a full refresh fails.
- `/api/data` reports refresh age, partial source errors, and stale status.
- A green/amber/red status indicator is shown in the dashboard header.
