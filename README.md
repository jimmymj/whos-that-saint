# Who’s That Saint? — All Saints' Day Find Activity

A mobile-first, full-screen parish community web application for All Saints' Day featuring ~100 participating children in Saint costumes, live Google Sheets cloud synchronization, and an Admin & Prize Ceremony Center.

## Key Features
- **Full-Screen Mobile-First UI (Fall Harvest Theme):** Zero wasted horizontal space, large touch targets, automatic native full-screen mode (`requestFullscreen`), and high-contrast Fall Mahogany & Harvest Gold palette.
- **Two Saints Find Activity Modes:**
  1. **Find Saint Name by Kid's Name** (`+10 pts` per match)
  2. **Find Kid's Name Dressed Up by Saint** (`+10 pts` per match)
- **Educational & Fair-Play Design:**
  - **✞ About the Saint:** Shown by default on every card.
  - **🧥 Outfit & Common Property Clue (`-2 pts`):** Optional clue stored in Google Sheets (`Outfit_Clue`) describing the child's Saint costume and props (`+8 pts` if matched with clue).
  - **Hidden Answer Key:** Individual correct answers are never revealed to participants during or after play.
  - **Unique Participant Names:** Duplicate participant names are blocked both locally and in Google Sheets.
- **Admin & Prize Ceremony Center (`PIN: 1101`):**
  - **Tab 1:** Cloud Webhook & CSV Data Manager
  - **Tab 2:** Kids & Saints Add / Edit / Delete Manager (two-way sync with Google Sheets `Roster_Saints_Kids`)
  - **Tab 3:** Master Participant Scores & Single-File Prize CSV Export (`whos_that_saint_prizes.csv`)
  - **Tab 4:** Top 3 Prize Podium & 100-Kid Saints Honor Roll Ticker
  - **Tab 5:** 8-Point Automated Security & End-to-End QA Diagnostic Suite

## Project Files
- `index.html` — Standalone interactive web application
- `Code.gs` — Google Apps Script backend for two-way Google Sheets sync
- `Roster_Saints_Kids_Template.csv` — Master 100-Kid Input Table CSV template
- `Quiz_Participants_Scores_Template.csv` — Participant Scores Output Table CSV template

