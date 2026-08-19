# DSE Maths Paper 1 Trainer

A mark-tracking trainer for **HKDSE Mathematics (Compulsory Part) Paper 1**, built as a single Google Apps Script web app with a Google Sheet as the database.

The app does **not** store any exam content. HKEAA owns the copyright of all past papers, so students work from paper copies or official PDFs and only enter their marks here.

## What it does

- Organises papers by **year**, each split into **A(1) / A(2) / B**, 35 marks per section.
- Records marks **part by part** (a, b, c), then rolls them up to question and section totals.
- Keeps every attempt (re-takes are never overwritten) and shows trend, weak topics and a year x section overview.
- Gives teachers a class overview with per-student progress and the five weakest topics.

## Design decisions

All 13 decisions (architecture, granularity, identity, topics, offline behaviour, UI) are recorded in `docs/CONSENSUS.md`.

Highlights:

| Area | Decision |
|---|---|
| Architecture | Single Apps Script web app, `doGet` + HtmlService SPA, Sheet as DB, `google.script.run` (no CORS) |
| Scoring | Per-part marks, integers 0..MaxMark, blanks block submission |
| Identity | Google account email + `Students` whitelist, `LockService` on writes |
| Topics | 28 topic IDs: 20 official Learning Units (SS01-SS20) + 8 junior secondary categories (JS01-JS08) |
| Reliability | Local queue, 2s/5s/15s/30s backoff, 20s timeout, idempotent by `AttemptID` |
| UI | Chinese/English toggle, mobile first, hand written CSS |

## Repository layout

```
apps-script/    files to paste into the Apps Script editor
data/           CSV templates for the Topics and Questions sheets
docs/           consensus document
SETUP.md        step by step installation and acceptance tests
```

## Quick start

1. Create a Google Sheet, then **Extensions > Apps Script**.
2. Paste the files from `apps-script/` (see `SETUP.md` for the exact file names).
3. Run `initSheets`, paste `data/Topics_seed.csv` into the `Topics` sheet.
4. Enter your question bank into `Questions`, then run **DSE Trainer > validate question bank**.
5. Deploy as a web app with **Execute as: User accessing the web app** and **Who has access: Anyone with Google Account**.

The deployment settings are not optional: with any other combination `Session.getActiveUser().getEmail()` returns an empty string and the whitelist cannot work.

## Scope of the MVP

Years 2023-2025 only. No timer, no ranking, no CSV/PDF export, no service worker.
