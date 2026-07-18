# Heavy Metal Lyrics Analysis

A data pipeline that collects, cleans, stores, and analyzes lyrics from hundreds of heavy metal bands across multiple subgenres — built as a portfolio project to demonstrate end-to-end data engineering and NLP skills.

---

## Project Overview

What does heavy metal actually talk about? This project answers that question by building a word frequency database from thousands of songs across subgenres like thrash, death, black, doom, power, and progressive metal. The goal is to identify linguistic patterns, compare vocabulary across subgenres and eras, and ultimately find what makes metal lyrics distinct from everyday English.

**Dataset scale:**
- 224 bands across 10 defined subgenres (UPDATE NUMBER)
- 20,255 songs analyzed (UPDATE)
- Countries represented: USA (101), UK (43), Sweden (23), Germany (16), Norway (14), Finland (11), and more (UPDATE)

---

## Pipeline Architecture

The project is structured as four independent scripts that form a sequential pipeline:

```
preview_discography.py
        ↓
  [Manual Review]
        ↓
get_band_lyrics_raw.py
        ↓
clean_lyrics.py
        ↓
analyze_lyrics.py
```

### 1. `preview_discography.py`
Queries **Discogs** and **MusicBrainz** APIs to build a structured Excel preview of a band's full discography — albums, EPs, singles, tracklists, subgenres, release years, and vocalist gender. Includes filtering to exclude live albums, compilations, and samplers automatically.

### 2. Manual Review
The Excel preview is reviewed before any lyrics are fetched. Tracks can be included or excluded, subgenres corrected, and release types verified. This human checkpoint is a deliberate design choice to ensure data quality before committing API calls.

### 3. `get_band_lyrics_raw.py`
Reads the approved Excel preview and fetches raw lyrics from the **Genius API** for each included track. Saves raw lyric files, updates the Excel with fetch status, and builds a JSON metadata file for each band. Includes:
- Fallback search logic for tracks not found under their full title
- Artist verification using fuzzy string matching to reject wrong results
- Detection of live/demo versions in returned results
- Duplicate detection across releases

### 4. `clean_lyrics.py`
Processes raw lyric files to produce clean versions ready for analysis:
- Strips Genius annotations, contributor headers, and section markers
- Expands inline repeat markers (e.g. `x3` → repeated lines)
- Normalizes accented characters and punctuation
- Flags songs with ambiguous section structure for manual review
- Updates the Excel with clean status per track

### 5. `analyze_lyrics.py`
Reads cleaned lyric files and populates a **SQLite database** with word frequency counts broken down by subgenre. Handles contraction normalization, possessive stripping, and subgenre mapping.

---

## Database Schema

```sql
artists       (id, name, country)
releases      (id, artist_id, artist_name, title, year, type)
subgenres     (id, name)
release_subgenres (release_id, artist_id, subgenre_id, release_name, artist_name, subgenre_name)
songs         (id, release_id, artist_name, title)
words         (id, word, traditional, thrash, death, black, groove, progressive, doom, power, glam, nu, other)
```

---

## Key Technical Features

- **Multi-API orchestration** — Discogs, MusicBrainz, and Genius APIs chained together, each used for what it does best
- **Fuzzy string matching** — `difflib` used to match album titles across APIs with different naming conventions, and to verify that Genius returned the correct artist
- **Rate limit handling** — automatic retry logic and adaptive delays for all three APIs
- **Human-in-the-loop design** — Excel review step between discovery and fetching prevents wasted API calls on bad data
- **Incremental processing** — all scripts are safe to re-run; already-processed files are skipped
- **Overnight batch processing** — bands list managed via Excel, scripts run unattended across hundreds of bands

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python | Core scripting language |
| SQLite | Word frequency database |
| openpyxl | Excel preview generation and status tracking |
| lyricsgenius | Genius API wrapper |
| discogs_client | Discogs API wrapper |
| musicbrainzngs | MusicBrainz API wrapper |
| difflib | Fuzzy string matching |
| re | Regex-based lyric cleaning |
| python-dotenv | Secure API key management |

---

## Reproducing the Dataset

Raw lyric files are not included in this repository out of respect for copyright — lyrics are the intellectual property of their respective songwriters and publishers.

To reproduce the dataset:

1. Clone this repo
2. Create a `.env` file with your API keys (see `.env.example`)
3. Install dependencies: `pip install -r requirements.txt`
4. Add bands to `bands_to_process.xlsx`
5. Run `preview_discography.py` and review the generated Excel files
6. Run `get_band_lyrics_raw.py` to fetch lyrics
7. Run `clean_lyrics.py` to clean the raw files
8. Run `analyze_lyrics.py` to populate the database

API keys required:
- [Genius API](https://genius.com/developers) — free
- [Discogs API](https://www.discogs.com/settings/developers) — free

---

## Analysis & Findings

*Coming soon — analysis and visualizations in progress.*

---

## About

Built by William Anderson as part of a data science portfolio. This project demonstrates skills in data collection, API integration, text processing, database design, and pipeline architecture.

[LinkedIn](https://www.linkedin.com/in/wanderson33245/) · [Portfolio](https://wandersondata.com/)
