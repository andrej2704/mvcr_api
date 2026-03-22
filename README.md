# mvcr_api
Simple API to search for visa status on MVCR site

## About

This Node.js/Express API searches for a visa application number in the publicly available Excel spreadsheet published by the Czech Ministry of Interior ([MVCR](https://www.mvcr.cz/clanek/informace-o-stavu-rizeni.aspx)). It fetches and caches the spreadsheet in memory, then searches it for the provided query string.

## Usage

### Install dependencies

```bash
npm install
```

### Start the server

```bash
npm start
```

The server listens on port `3000` by default (configurable via the `PORT` environment variable).

### API

**GET** `/`

| Query parameter | Required | Description |
|---|---|---|
| `search` | Yes | The string to search for (e.g. visa application number) |
| `page` | No | Zero-based sheet index in the Excel file (default: `0`) |

#### Example request

```
GET /?search=OAM-12345-2024
```

#### Example responses

Found:
```json
{ "found": "Found!!! Pick it Up!", "fileName": "soubor/priloha/file.xlsx" }
```

Not found:
```json
{ "found": "NOT Found!!!", "fileName": "soubor/priloha/file.xlsx" }
```

Missing search parameter:
```json
{ "found": "Please specify what to search!" }
```

External fetch error (e.g. MVCR site unreachable):
```json
{ "error": "Failed to fetch data from MVCR." }
```
