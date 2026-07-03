# AlertTracker ↔ Chargeblast Sync

Google Apps Script that keeps two Google Sheets in sync via a shared `alert_id` key.

## What it does
- **Chargeblast → Tracker:** pulls `Chargeblast Action Date`, `Reference Information`, and `Notes from Chargeblast` into the Alert Tracker.
- **Tracker → Chargeblast:** upserts any row marked `Refunded? = NO` into the escalations sheet.

## Features
- Header-name column matching (resilient to column reordering)
- Only writes rows that actually changed
- Duplicate `alert_id`s are detected, skipped, and reported by email
- Batched reads/writes for performance at scale (2,000+ rows)
- Lock-based concurrency protection
- Runs automatically every 15 min + on edit in either sheet

## Setup
1. Bind the script to the Tracker spreadsheet (Extensions → Apps Script)
2. Set `AlertSync_FILE1_ID` / `AlertSync_FILE2_ID` and header config at the top
3. Run `AlertSync_testConfig` to validate
4. Run `AlertSync_setupTriggers` to go live