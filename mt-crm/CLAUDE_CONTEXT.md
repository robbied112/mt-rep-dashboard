# Missing Thorn CRM - Monorepo

## What This Repo Is
Single source of truth for the Missing Thorn BI dashboard system. Contains HTML templates,
build scripts, deployment config, and documentation. The two public GitHub Pages repos
(`mt-dashboard`, `mt-rep-dashboard`) are deployment targets only.

## Architecture

```
mt-crm/  (THIS REPO - private)
├── templates/
│   ├── executive_dashboard.html    ← Template with __DATA_START__/__DATA_END__ markers
│   └── rep_dashboard_template.html ← Template with __DATA_START__/__DATA_END__ markers
├── scripts/
│   ├── build_dashboard.py          ← Reads Excel → injects data → dist/index.html
│   └── build_rep_dashboard.py      ← Reads Excel → injects data → dist/index_rep.html
├── docs/
│   ├── FIREBASE_SETUP_GUIDE.md
│   └── CRM_PANEL_INTEGRATION_SUMMARY.md
├── dist/
│   ├── index.html                  ← Built executive dashboard (gitignored? or committed)
│   └── index_rep.html              ← Built rep dashboard
├── deploy.sh                       ← Pushes dist/ files to GitHub Pages repos
├── Makefile                        ← make build / make deploy / make all
├── .env.example
├── .gitignore
└── CLAUDE_CONTEXT.md
```

## Deployment Repos (public, GitHub Pages)
- `robbied112/mt-dashboard` → serves dist/index.html as executive dashboard
- `robbied112/mt-rep-dashboard` → serves dist/index_rep.html as rep dashboard

## Build Flow
1. Excel files live in OneDrive/SharePoint (NOT in this repo)
2. Build scripts read Excel files via `--data` flag
3. Scripts generate JavaScript data objects and inject between `__DATA_START__`/`__DATA_END__` markers
4. Output goes to `dist/`
5. `deploy.sh` clones Pages repos, copies built HTML, commits, pushes

```bash
# Build with OneDrive data
make build DATA_DIR=~/OneDrive/MT\ Dashboard\ Data/

# Deploy to GitHub Pages
make deploy

# Or do everything
make all DATA_DIR=~/OneDrive/MT\ Dashboard\ Data/
```

## Template Structure

### Executive Dashboard (~2,390 lines without data)
- Lines 1-298: HTML head, CDN libs, CSS
- Lines 299-660: HTML body (header, nav, tabs, filter bar)
- Lines 660-675: JS app setup + `__DATA_START__` marker
- Lines 675+: `__DATA_END__` marker → rest of JS logic
- Final ~500 lines: Firebase config, AccountStore, CRM panel, Account CRM Manager tab

### Rep Dashboard (~3,228 lines without data)
- Lines 1-657: HTML head, CDN libs, CSS
- Lines 658-1360: HTML body (header, nav, tabs, filter bar)
- Line 1361: `__DATA_START__` marker
- Line 1363: `__DATA_END__` marker → rest of JS logic
- Final ~550 lines: Firebase config, AccountStore, CRM panel, My Action Items tab

## Data Variables

### Executive Dashboard
`stateNames`, `regionMap`, `distScorecard`, `inventoryData`, `placementData`,
`newAccounts`, `accountsTop`, `reorderData`, `qbDistOrders`, `warehouseInventory`,
`classicTracker`, `revMonths`, `revMonthLabels`, `revTotal`, `revTrend`, `revTxns`,
`revUnits`, `productMix`, `revMonths2026`, `revMonthLabels2026`, `revTotal2026`,
`revTrend2026`, `revTxns2026`, `revUnits2026`, `productMix2026`, `ytdMonths`,
`ytdMonthLabels`, `ytdChannelRev`, `ytdDistRevByState`, `ytdBudget`,
`ytdBudgetTotal`, `revTotalByMonth`, `customerRevenue`, `acctConcentration`,
`distDetail`, `sampleSummary`, `buildDate`, `dataThrough`

### Rep Dashboard
`stateNames`, `regionMap`, `distScorecard`, `accountsTop`, `inventoryData`,
`distHealth`, `reEngagementData`, `newWins`, `reorderData`, `qbDistOrders`,
`warehouseInventory`, `classicTracker`, `distDetail`, `placementSummary`,
`acctConcentration`, `sampleSummary`, `buildDate`, `dataThrough`

## Firebase
- Project: `mt-dashboard-dab8e`
- Both dashboards share the same Firestore database
- See `docs/FIREBASE_SETUP_GUIDE.md` for details

## Key Business Context
- **Company**: Missing Thorn (CPG wine brand)
- **Products**: 12 SKUs across Still Red/White/Rose, Sparkling White/Rose
- **Distribution**: 16-18 distributors across 12-13 states (East/West regions)
- **Users**: CRO (executive dashboard) + 1 rep (rep dashboard), growing to 4-6 reps
- **Key Metrics**: CE (Case Equivalents), DOH (Days on Hand), Sell-Through Rate, Health Score
