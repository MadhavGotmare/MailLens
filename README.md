# MailLens

MailLens is a Google Apps Script companion for a Google Sheet that scans Gmail in safe batches, creates a durable sender inventory, identifies likely newsletters, and turns the result into a simple decision dashboard.

## What it does

- Adds a **MailLens** menu to the bound Google Sheet.
- Scans Gmail with a user-supplied Gmail search query (`in:anywhere` by default).
- Persists the scan offset and query in `PropertiesService`, so a time-limited execution continues with a one-minute time trigger.
- Aggregates sender email, display name, message count, first/last seen date, newsletter signals, and unsubscribe signals in `Senders`.
- Creates a `Dashboard` with summary metrics, a top-senders table, and a newsletter-versus-other pie chart.

## Installation

1. Create a Google Sheet that will hold the MailLens report.
2. In the Sheet, choose **Extensions → Apps Script**.
3. Add each `.gs` file from this repository to the Apps Script project, and replace the project manifest with `appsscript.json`.
4. Save the project, reload the spreadsheet, and approve the requested Spreadsheet, Gmail, and trigger permissions when you first choose **MailLens → Start scan**.
5. Choose **MailLens → Start scan**. Enter a Gmail query such as `label:inbox newer_than:1y`, or leave it blank to scan all mail.

> The script is intended to be **bound to the destination spreadsheet**. Run it from that spreadsheet so reports are written to the correct file.

## Operation and limits

A batch contains 100 Gmail threads. The current query, thread offset, and destination spreadsheet ID are saved in document properties after each batch, then the next batch is scheduled using a time-based trigger. This lets trigger-based executions reopen the report spreadsheet even though they do not have an active spreadsheet context. This limits individual runs and makes long scans resumable. Use **Continue scan now** to advance an existing scan manually, or **Reset scan** to remove its saved progress and continuation trigger.

Gmail search ordering can change as new mail arrives. For the most reproducible one-time audit, use a bounded query (for example, `before:2026/01/01`). Sender counts are message counts, not thread counts. Newsletter detection is heuristic: `List-Unsubscribe`, `List-Id`, common newsletter-like subject terms, and unsubscribe text are signals, not guarantees.

## Privacy

MailLens runs entirely in your Google account through Apps Script. It stores aggregate sender information in your selected spreadsheet; it does not send mailbox contents to an external service.

## License

[MIT](LICENSE)
