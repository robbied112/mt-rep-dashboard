# Missing Thorn Rep Dashboard - Project Context

## Overview

Sales representative performance and opportunity tracking dashboard for Missing Thorn, a CPG wine brand selling through distributors. Single-file HTML application deployed on GitHub Pages with Firebase Firestore for CRM data persistence.

## Architecture

### Single-File HTML Dashboard
- **File**: `index.html` (~356 KB, ~4,535 lines)
- **Sections**:
  - Lines 1-657: HTML head, external CDN libraries, embedded CSS
  - Lines 658-1360: HTML body structure (header, nav tabs, filter bar, tab content areas)
  - Lines 1360-4535: JavaScript application logic
  - Lines 1361-2670: `__DATA_START__` / `__DATA_END__` markers (injected data)
  - Lines 3983-4535: Firebase config, AccountStore, CRM panel, Action Items tab

### External Dependencies (CDN)
- **Google Fonts**: Catamaran
- **Chart.js** v4.4.1: All charts and visualizations
- **SheetJS (XLSX)** v0.18.5: Excel export functionality
- **Firebase** v10.8.0: Firestore for CRM data persistence

### Build Pipeline
Python scripts (not in this repo) read 6+ Excel source files:
- VIP distributor portal exports (depletions, inventory, placements)
- QuickBooks exports (orders, revenue)

Scripts generate JavaScript data objects and inject them between `__DATA_START__` / `__DATA_END__` markers in the HTML template. The injected data objects are:

| Object | Description |
|--------|-------------|
| `stateNames` | State abbreviation → full name mapping |
| `regionMap` | State → region mapping (East/West) |
| `distScorecard` | 18 distributors with depletion data, CE, momentum, sell-through, inventory, DOH, weekly velocity |
| `accountsTop` | 80+ top accounts with monthly depletion data (Nov-Feb), trend, growth potential |
| `inventoryData` | Stock levels by location, depletion rates, SKU breakdown |
| `distHealth` | SKU-level sell-in/sell-through by wine type |
| `depletionsData` | New market and re-engagement opportunities |
| `reEngagementData` | 100+ accounts with prior volume for re-engagement |
| `newWins` | Recently activated accounts (2026) |
| `reorderData` | 80+ accounts with order cycles, last order dates, days since order |
| `qbDistOrders` | QuickBooks order history keyed by distributor name |
| `warehouseInventory` | Classic vs Contemporary wine inventory by location |
| `classicTracker` | Historical inventory tracking with burn rate analysis |
| `buildDate` | Build timestamp |
| `dataThrough` | Data freshness date |

### Firebase / Firestore

**Project**: `mt-dashboard-dab8e` (shared with executive dashboard)

**Collections**:
- `accounts/{accountId}` — Document per account (ID = sanitized account name via `replace(/[^a-zA-Z0-9]/g, '_')`)
  - Fields: `tags` (array), `status` (string), `followUp` (date), `nextAction` (string)
  - Subcollection: `notes/{noteId}` — Notes with `text`, `author`, `ts`, `date`

**AccountStore** (lines ~3983-4051):
- Same dual-mode pattern as executive dashboard
- Firebase primary, localStorage fallback with key `mt_acct_{sanitizedId}`
- Rep name stored in `localStorage` as `mt_rep_name`

### Authentication
- **None**. No login or access control.
- Rep name prompted on first panel open and stored in localStorage
- Firebase API keys are hardcoded (standard for client-side, secured by Firestore rules)

## Dashboard Tabs

1. **Performance Overview** — KPIs, monthly trends, distributor scorecard
2. **Depletions** — Top distributors, 13W CE trends, sell-through ratios
3. **Distributor Health** — Sell-in vs sell-through by SKU, monthly purchasing, account penetration, inventory coverage
4. **Inventory** — Stock status by location, days-on-hand, reorder opportunities, overstock alerts
5. **Account Insights** — Top accounts ranking, depletion trend analysis
6. **Opportunities** — Re-engagement targets, new wins, growth potential identification
7. **Reorder Forecast** — 80+ accounts with order cycle tracking, days since order, priority scoring
8. **My Action Items** — Tagged accounts, action follow-ups, urgency indicators (purple-highlighted tab)
9. **Key / Legend** — Definitions and status explanations

## CRM Features

### Account Panel (slide-in overlay)
Opens on account click from any table. Gathers data from 4 sources:
- `accountsTop` (depletion data)
- `reorderData` (order cycle data)
- `reEngagementData` (re-engagement info)
- `newWins` (new activation info)

### Tags
8 predefined: VIP, At Risk, Hot Lead, Needs Visit, Seasonal, Chain, Independent, Follow Up

### Notes
- Textarea entry with author attribution (rep name from localStorage)
- Timestamped, sorted descending
- Stored in Firestore subcollection

### Next Actions
Dropdown: Call, Email, Visit, Send Samples, Follow Up, Reorder Check
- Paired with follow-up date field

### My Action Items Tab
- Filters: All, Needs Visit, Follow Up, At Risk, Hot Lead, VIP, Has Next Action, Overdue Follow-ups
- KPI cards for action metrics
- Card-based display with urgency indicators
- Shows tagged/action-flagged accounts across all data

## Key Business Context

- **Product**: Wine (12 SKUs across Still Red/White/Rose, Sparkling White/Rose)
- **Distribution**: Through 18+ distributors across 13 states in East/West regions
- **East**: NJ, FL, PA, NC, GA, CT, ME, SC
- **West**: TX, AZ, NV, CA
- **Key Metrics**: CE (Case Equivalents), DOH (Days on Hand), Sell-Through Rate, Consistency, Momentum
- **Users**: 1 rep currently, potentially growing to 4-6 reps

## Deployment
- GitHub Pages (static hosting)
- Repo: `robbied112/mt-rep-dashboard`
- Branch: `main` (deployed)

## Companion Repository
- **mt-dashboard** (`robbied112/mt-dashboard`): Executive/CRO-facing version with revenue analytics, Account CRM manager view, and budget tracking
- Shares the same Firebase project (`mt-dashboard-dab8e`) for CRM data
