# Translation volume reporter (Google Apps Script)

An internal tool I built while running operations for a translation agency. Billing was calculated by hand: someone opened every client's folder on Google Drive and added up the volume of translated text. That took 6–8 hours a month and produced errors. This script does it automatically and builds a monthly report in Google Sheets.

## What it does

- Finds each client's folder on Google Drive. Folder names rarely matched the client list exactly (typos, different word order, extra notes), so matching uses Levenshtein distance with a small tolerance.
- Extracts text from Google Docs, Sheets and Slides, Microsoft Word, Excel and PowerPoint files, and PDFs (via conversion), then counts characters.
- Skips documents that shouldn't be billed, based on a configurable list of phrases checked at the start of each file.
- Writes a summary sheet plus a detail sheet per client, with character counts and links to every file.

## How it works

1. Reads the client list and matches each client to a Drive folder (fuzzy matching, see `isClientNameMatch`).
2. Walks the year folders and extracts text from every supported file (`getTextAndCount`).
3. Counts characters per file and per client.
4. Writes a summary sheet and a detail sheet per client.
5. Before the 6-minute limit, saves progress and schedules `continueProcessingClients` to pick up where it stopped.

## Entry points

| Function | Use |
|---|---|
| `initializeClientProcessing()` | start a fresh monthly run |
| `continueProcessingClients()` | resumes automatically via trigger |
| `reprocessLastClient()` | rerun one client after fixing its files |

## Example output (sample data)

| Client | Files | Characters |
|---|---|---|
| Client A | 12 | 48,310 |
| Client B | 5 | 17,902 |
| **Total** | **17** | **66,212** |

## Working around the execution limit

Apps Script stops a run after 6 minutes, and a full month of files takes longer than that. The script checks the elapsed time, saves its position in `PropertiesService` before the limit, and schedules a time-based trigger to continue from the same place a minute later. Any amount of files gets processed without anyone babysitting the run.

## Setup

1. Create a new Apps Script project and paste `Code.gs`.
2. Fill in the constants at the top: source folder IDs, the output spreadsheet ID, and the list of excluded phrases.
3. Enable the Drive advanced service (needed for converting Office files and PDFs).
4. Run initializeClientProcessing() once.

All IDs in this repository are placeholders. No client data is included.

## Stack

Google Apps Script (JavaScript) · Google Drive API · Google Docs/Sheets/Slides services


## What I'd improve next

- Move IDs and settings to Script Properties instead of constants
- Unit tests for the matching logic (clasp + Jest)
- OCR for scanned PDFs, which currently count as empty
- A log sheet listing files that failed to convert
