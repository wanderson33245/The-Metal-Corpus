# Heavy Metal Lyrics Analysis

A data pipeline that collects, cleans, stores, and analyzes lyrics from hundreds of heavy metal bands across multiple subgenres. Built as a portfolio project to demonstrate end-to-end data engineering and NLP skills.

---

## Project Overview

What does heavy metal actually talk about? This project answers that question by building a word frequency database from thousands of songs across subgenres like thrash, death, black, doom, power, and progressive metal. The goal is to find what makes metal lyrics distinct from everyday English, and how it varies by subgenre.

**Dataset scale:**
- 539 bands across 10 defined subgenres (UPDATE NUMBER)
- 39,116 songs analyzed (UPDATE)
- 45 countries represented: USA (245), UK (62), Germany (31), Sweden (28), Norway (17), Finland (18), and more (UPDATE)

---

## The Story Behind This Project

I've been a metalhead ever since I first heard the album *The Stage* by Avenged Sevenfold. Sometime back then, I read a post where someone had analyzed a few thousand metal songs and concluded that the most common words were along the lines of "death" and "fire" and the least common was something like "senators". I remember my immediate reaction was that it just sounded like something someone who doesn't listen to metal would make up.

So, I wrote a rudimentary program in JavaScript, copy/pasted 2,625 songs' lyrics into it, and it returned two arrays: a list of words and a count of their occurrence. That revealed that the most common word in heavy metal is "the" and the least common word is a ten thousand-way tie between however many words only appeared once in all 2600 lyrics. It was a good start, but I lacked the expertise to analyze it any further at that time.

Years later, looking for a portfolio project that wasn't another sales data analysis, I remembered the core question of that project: What **does** metal music talk about? How do the subgenres differ? This time, I had the skills to do it properly and get answers to questions that counting alone can't provide.

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
Queries **Discogs** and **MusicBrainz** APIs to build a structured Excel preview of a band's full discography: albums, EPs, singles, tracklists, subgenres, release years, and vocalist gender. Filters to exclude live albums, compilations, and samplers automatically.

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

- **Multi-API orchestration** — Discogs, MusicBrainz, and Genius APIs chained together. Discogs for band discography information, MusicBrainz for release type information, Genius for lyrics
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

The raw lyric files I used are not included in this repository out of respect for copyright. Lyrics are the intellectual property of their respective songwriters and publishers.

To reproduce the dataset:

1. Clone this repo
2. Create a `.env` file with your API keys (see `.env.example`)
3. Install dependencies: `pip install -r requirements.txt`
4. Create an Excel file `bands_to_process.xlsx` with band names, countries of origin, and vocalist genders
6. Run `preview_discography.py` and review the generated Excel files
7. Run `get_band_lyrics_raw.py` to fetch lyrics
8. Run `clean_lyrics.py` to clean the raw files
9. Run `analyze_lyrics.py` to populate the database

API keys required:
- [Genius API](https://genius.com/developers) — free
- [Discogs API](https://www.discogs.com/settings/developers) — free

---

## Analysis & Findings

The full analysis lives in [The Metal Corpus.ipynb](The Metal Corpus.ipynb). Here's what's covered:

- **The Most Common Words in Heavy Metal** — stop words removed, what's left is revealing
- **Most Unique Words by Subgenre** — what makes death metal vocabulary distinct from doom?
- **The You/I Ratio** — which genres address the audience, which turn inward
- **Profanity by Subgenre** — where the f-bombs actually land might surprise you
- **Black Metal's Language Problem** — Finnish and Norwegian words in the corpus
- **Glam vs Death: Contractions** — nothin', livin', runnin' vs clinical precision
- **Bands per Capita vs World Happiness Index** — Scandinavia tops both lists
- ...and more

---

## What's Next?

This project answers a lot of questions, but it opens the floodgates to many more. I've achieved my original goal, and so in an effort to avoid scope creep, I'm deliberately publishing this project with some further questions left unanswered. I'd like to come back someday and answer some of those in a V2, but today is not that day. Here are some of those further questions:

### Cross-genre comparison

The most natural extension to this project is to examine lyrics from **all** genres. Country, R&B, Gospel, Funk, Rock, Heavy Metal, Rap, etc... I've already built the infrastructure, it would just require hundreds of thousands more API calls, and quite a long time to make them. One comparison that's fascinated me for a couple years is *Satan Is Real* by The Louvin Brothers (1959, Country gospel) and *Satan Is Real* by Kreator (2017, Thrash metal). Both are a warning against evil and temptation, but if you judge Kreator's album by its cover, you would think different. I'd posit that thrash metal and country are more thematically similar than most people would expect, but my proof for that claim will have to wait until I do V2. That's the kind of assumption this project exists to test.

### Phrase and idiom analysis

Word frequency can only go so far. A logical next step is to analyze whole phrases. What idioms are used and by which genres most often? Are any genre-specific? Which genres lean most heavily on figurative language like similes, metaphors, and personification? Which genres favor direct and literal interpretation? That would reveal a lot of *artistic* information and context behind the words, and would likely require advanced NLP or n-gram modeling.

### Sentiment analysis

Beyond the words or phrases themselves, how they're used matters. Sentiment and positive/negative analysis by subgenre would be a challenging addition. Which subgenre is the most positive? Which is the most negative? Which are the most emotionally intense? What nouns appear with the most frequently negative adjectives? For example, a box-and-whisker plot showing the sentiment distribution of adjectives paired with common religious or political nouns by subgenres, and eventually by broader genres like Country vs. Rap, would reveal *how* the genres feel about those topics.

### Non-English lyric translation

Many of the lyrics collected from Scandinavian countries are in non-English languages, limiting the breadth of lyric collection from those countries and, specifically, black metal. Integrating a translation API would allow those lyrics to contribute to this project rather than appearing as noise.

### Metal Archives scraping

The [Metal Archives](https://www.metal-archives.com/) (Encyclopaedia Metallum) is one of the most comprehensive databases for metal music, with band information, line-ups, reviews, lyrical themes, and often the lyrics themselves for tens of thousands of artists. It doesn't have a dedicated public API, but scraping it responsibly would be a fun networking challenge and would significantly improve this dataset's coverage of underground artists. 

### AI integration

The human-in-the-loop validation step is the most time-consuming part of the pipeline. In a future extension of this project, I would like to train a classification model to automate it. With 2,048 validated previews in hand, the training data is there.

---

## About

Built by William Anderson as part of a data science portfolio. This project demonstrates skills in data collection, API integration, text processing, database design, and pipeline architecture.

[LinkedIn](https://www.linkedin.com/in/wanderson33245/) · [Portfolio](https://wandersondata.com/)
