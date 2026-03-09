# Firebase Setup Guide

## Project Info
- **Project**: mt-dashboard-dab8e
- **Console**: https://console.firebase.google.com/project/mt-dashboard-dab8e

## Firestore Collections

### `accounts/{accountId}`
Account ID = sanitized account name: `name.replace(/[^a-zA-Z0-9]/g, '_')`

**Fields:**
- `tags` (array) — VIP, At Risk, Hot Lead, Needs Visit, Seasonal, Chain, Independent, Follow Up
- `status` (string) — Account status
- `followUp` (date) — Next follow-up date
- `nextAction` (string) — Call, Email, Visit, Send Samples, Follow Up, Reorder Check

**Subcollection: `notes/{noteId}`**
- `text` (string) — Note content
- `author` (string) — Rep name
- `ts` (timestamp) — Server timestamp
- `date` (string) — Human-readable date

## Authentication
- No user authentication currently implemented
- Rep name stored in localStorage as `mt_rep_name`
- Firebase API keys are hardcoded in templates (standard for client-side apps)
- Security enforced via Firestore rules

## Shared Access
Both dashboards (executive + rep) share the same Firebase project and Firestore data.
CRM notes and tags entered in either dashboard are visible in both.
