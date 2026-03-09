# CRM Panel Integration Summary

## Overview
Both dashboards include a slide-in CRM panel that opens when clicking any account name
in any table. The panel provides per-account notes, tags, and next-action tracking backed
by Firebase Firestore.

## Components

### AccountStore
Dual-mode data layer (Firebase primary, localStorage fallback):
- `getNotes(acctId)` — Fetch notes from Firestore, fall back to localStorage
- `addNote(acctId, text, author)` — Add timestamped note
- `getMeta(acctId)` — Get tags, status, followUp, nextAction
- `setMeta(acctId, meta)` — Persist metadata with merge

### CRM Panel (both dashboards)
- Tag selector (8 predefined tags)
- Next action dropdown + follow-up date
- Notes textarea with author attribution
- Displays account context data from multiple data sources

### Account CRM Manager (executive dashboard only)
- Tab 9: filterable/sortable report of all CRM-tagged accounts
- Filters by tag and state
- Sort by priority, days since order, CE, tags
- Alert cards: stale VIPs (>30 days), at-risk, overdue follow-ups

### My Action Items (rep dashboard only)
- Tab 8: card-based display of accounts with CRM activity
- Filters: All, Needs Visit, Follow Up, At Risk, Hot Lead, VIP, Has Next Action, Overdue
- KPI cards for action metrics
- Urgency indicators

## Data Sources for CRM Panel
The panel aggregates context from multiple data objects:
- `accountsTop` — Depletion data, monthly trends
- `reorderData` — Order cycle data, days since order
- `reEngagementData` — Re-engagement info (rep dashboard)
- `newWins` — New activation info (rep dashboard)
- `distScorecard` — Distributor-level metrics
