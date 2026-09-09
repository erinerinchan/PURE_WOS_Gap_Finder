# PURE vs WoS Gap Finder

A browser-only tool for comparing Web of Science exports with PURE data and identifying missing records.

## Run locally

Open `index.html` in a modern browser, or serve the folder with any static web server.

## Supported inputs

- `WOS.xlsx` with columns such as `Author Full Names`, `Article Title`, `Source Title`, `DOI`, `Publication Year`, `Addresses`, and `UT (Unique WOS ID)`
- `PURE.xlsx` with columns such as `Title of the contribution in original language`, `Subtitle of the contribution in original language`, `DOI`, `Journal`, and `Publication Year`

## Features

- Auto-detects key columns by header name
- Manual column mapping when headers differ
- Affiliation filtering for HKUST / Guangzhou rules
- DOI match, exact title match, and fuzzy title review logic
- Summary cards and result tabs
- Export missing records to `.xlsx`
- All processing stays in the browser
