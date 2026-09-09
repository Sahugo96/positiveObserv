# Roulette Observer

Chrome Manifest V3 extension that observes roulette rounds, keeps their Fairplay
data in extension storage, and periodically exports the collected entries to an
Excel workbook. It also adds totals and row-exclusion controls to the site's
history table.

## Structure

- `background.js` owns durable extension storage and handles content-script
  messages.
- `content/roulette-observer.js` observes rounds, reads Fairplay data, manages
  the local fallback, and creates Excel downloads.
- `content/history-summary.js` calculates history-table totals and manages the
  row-exclusion buttons.
- `lib/xlsx.full.min.js` is the vendored SheetJS browser build used for exports.

## Running locally

Open `chrome://extensions`, enable **Developer mode**, choose **Load unpacked**,
and select this repository. The extension currently runs on every URL, but each
content module exits safely when its target page elements are absent.

The roulette observer exposes these helpers in the page's content-script context:

- `downloadExcelFile()`
- `getUnsavedCount()`
- `resetUnsavedCount()`
- `setAutoDownloadThreshold(number)`
- `getAutoDownloadThreshold()`
