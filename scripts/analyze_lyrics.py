# The following bands' lyrics were manually verified as training data while
# building the cleaning pipeline. I tested the cleaning regex accuracy against
# this trusted, hand-checked source before running the pipeline for the full
# 1,100+ band corpus.
#
# Ad Cinerem -- Full discography --                           1 album, 1 EP, 1 demo
# Alestorm   -- Full Discography --                           8 albums, 2 EPs, 1 single
# Amon Amarth -- Full discography                             12 albums, 2 EPs
# ...and.so.civilization.strides.on... -- Full discography -- 1 album
# Anthrax -- Full discography --                              13 albums, 2 EPs, 4 singles
# Black Sabbath -- Full discography --                        19 albums and some extras in The End (2016) EP
# Decapitated -- Full discography --                          8 albums, 1 demo
# Dream Theater -- Full discography --                        16 albums, 1 EP
# Mick Gordon -- Anything from the DOOM soundtracks           2 albums (Kinda)
# Mortal Sin -- First album                                   1 album
# 0 x i s t -- Full discography --                            2 albums, 1 EP
# Pinkly Smooth -- Full discography                           1 album
# Rings of Saturn -- Full discography --                  All 5 albums with lyrics
# Sabaton -- Full discography --                             ~12 albums
# The Beatles -- Yes, I'm counting Helter Skelter             1 song


import re
import sqlite3
import os
import json
from dotenv import load_dotenv
import time

load_dotenv()

vs_folder = os.environ.get("VS_FOLDER")
lyrics_folder_name = "Lyrics"
lyrics_folder_path = os.path.join(vs_folder, lyrics_folder_name)

start_time = time.time()

conn = sqlite3.connect("metal_lyrics.db")
cursor = conn.cursor()
cursor.execute("PRAGMA journal_mode=WAL")
cursor.execute("PRAGMA synchronous=NORMAL")


# ------- INITIALIZATION -----------------------------------------------------------------------


cursor.execute("""
    CREATE TABLE IF NOT EXISTS artists (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        country TEXT
    )
""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS subgenres (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE
    )
""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS releases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        artist_id INTEGER,
        artist_name TEXT,
        title TEXT,
        year INTEGER,
        type TEXT,
        vocalist_gender TEXT,
        FOREIGN KEY (artist_id) REFERENCES artists(id)
    )
""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS release_subgenres (
        release_id INTEGER,
        artist_id INTEGER,
        subgenre_id INTEGER,
        release_name TEXT,
        artist_name TEXT,
        subgenre_name TEXT,
        vocalist_gender TEXT,
        FOREIGN KEY (release_id) REFERENCES releases(id),
        FOREIGN KEY (artist_id) REFERENCES artists(id),
        FOREIGN KEY (subgenre_id) REFERENCES subgenres(id),
        PRIMARY KEY (release_id, subgenre_id)
    )
""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS songs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        release_id INTEGER,
        artist_name TEXT,
        title TEXT,
        FOREIGN KEY (release_id) REFERENCES releases(id)
    )
""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS words (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        word TEXT UNIQUE,
        heavy INTEGER DEFAULT 0,
        thrash INTEGER DEFAULT 0,
        death INTEGER DEFAULT 0,
        black INTEGER DEFAULT 0,
        groove INTEGER DEFAULT 0,
        progressive INTEGER DEFAULT 0,
        doom INTEGER DEFAULT 0,
        power INTEGER DEFAULT 0,
        glam INTEGER DEFAULT 0,
        nu INTEGER DEFAULT 0,
        metalcore INTEGER DEFAULT 0,
        alt INTEGER DEFAULT 0,
        gothic INTEGER DEFAULT 0,
        other INTEGER DEFAULT 0
    )
""")

conn.commit()

# --------------- END INITIALIZATION ------------------------------------------------------------------


def get_or_create_artist(name, country):
    cursor.execute("INSERT OR IGNORE INTO artists (name, country) VALUES (?, ?)", (name, country))
    cursor.execute("SELECT id FROM artists WHERE name = ?", (name,))
    return cursor.fetchone()[0]

def get_or_create_release(artist_id, artist_name, title, year, release_type, vocalist_gender):
    cursor.execute("INSERT OR IGNORE INTO releases (artist_id, artist_name, title, year, type, vocalist_gender) VALUES (?, ?, ?, ?, ?, ?)", (artist_id, artist_name, title, year, release_type, vocalist_gender))
    cursor.execute("SELECT id FROM releases WHERE artist_id = ? AND title = ?", (artist_id, title))
    return cursor.fetchone()[0]

def get_or_create_release_subgenre(release_id, artist_id, release_name, artist_name, subgenre, vocalist_gender):
    cursor.execute("INSERT OR IGNORE INTO subgenres (name) VALUES (?)", (subgenre,))
    cursor.execute("SELECT id FROM subgenres WHERE name = ?", (subgenre,))
    subgenre_id = cursor.fetchone()[0]
    cursor.execute("INSERT OR IGNORE INTO release_subgenres (release_id, artist_id, subgenre_id, release_name, artist_name, subgenre_name, vocalist_gender) VALUES (?, ?, ?, ?, ?, ?, ?)", (release_id, artist_id, subgenre_id, release_name, artist_name, subgenre, vocalist_gender))

def get_or_create_song(release_id, artist_name, title):
    cursor.execute("INSERT OR IGNORE INTO songs (release_id, artist_name, title) VALUES (?, ?, ?)", (release_id, artist_name, title))
    cursor.execute("SELECT id FROM songs WHERE release_id = ? AND title = ?", (release_id, title))
    return cursor.fetchone()[0]


# -----------------------------FUNCTION TO DETECT THE WORDS ----------------------------

style_to_subgenre = {
    "heavy metal": "heavy",
    "speed metal": "heavy",
    "traditional": "heavy", # I used to call "heavy" "traditional", but it just got complicated because "heavy metal" has changed drastically over time and it just started issues
    "neo-classical metal": "heavy",
    "doom metal": "doom",
    "funeral doom metal": "doom",
    "doom": "doom",
    "thrash": "thrash",
    "crossover thrash": "thrash",
    "speedcore": "thrash",
    "death metal": "death",
    "death": "death",
    "technical death metal": "death",
    "brutal death metal": "death",
    "melodic death metal": "death",
    "deathcore": "death",
    "deathrock": "death",
    "black metal": "black",
    "atmospheric black metal": "black",
    "unblack metal": "black",
    "black": "black",
    "blackgaze": "black",
    "groove metal": "groove",
    "groove": "groove",
    "progressive metal": "progressive",
    "progressive": "progressive",
    "prog": "progressive",
    "power metal": "power",
    "power": "power",
    "symphonic metal": "power",
    "glam metal": "glam",
    "glam": "glam",
    "hair metal": "glam", #I don't think "hair metal" ever occurs in here, but just in case
    "nu metal": "nu",
    "nu": "nu",
    "metalcore": "metalcore",
    "metallic hardcore": "metalcore",
    "alt metal": "alt",
    "alternative metal": "alt",
    "gothic metal": "gothic",
    "other": "other",
    "viking metal": "other",
    "post-metal": "other",
    "folk metal": "other",
    "sludge metal": "other",
    "industrial metal": "other",
    "grindcore": "other",
    "goregrind": "other",
    "funk metal": "other",
}

contractions = {"it's", "he's", "she's", "that's", "what's", "there's", "here's", "who's", "how's"}

replacements = { #not redundant
    "`": "'",
    "ü": "u",
    "ö": "o",
    "ä": "a",
    "ë": "e",
    "ï": "i",
    "á": "a",
    "é": "e",
    "í": "i",
    "ó": "o",
    "ú": "u",
    "ñ": "n",
    "å": "a",
    "ø": "o",
    "æ": "ae",
    "œ": "oe"
}

def normalize_text(text): #not redundant
    for accented, replacement in replacements.items():
        text = text.replace(accented, replacement)
    return text

def get_unique_subgenres(raw_subgenres):
    mapped = set()
    for style in raw_subgenres:
        #print(f"  Checking style: '{style}'")
        result = style_to_subgenre.get(style.strip().lower())
        if result:
            mapped.add(result)
        else:
            print(f"  [unmapped subgenre, skipping] {style}")
    return mapped

def get_words(raw_subgenres, text):
    unique_subgenres = get_unique_subgenres(raw_subgenres)
    if not unique_subgenres:
        return

    text = normalize_text(text.lower())
    found_words = re.findall(r"[a-z']+", text)

    for word in found_words:
        # Remove apostrophe at the start of a word ('til -> til)
        if word.startswith("'"):
            word = word[1:]
        # Remove apostrophe at the end of a word (brothers' -> brothers)
        if word.endswith("'"):
            word = word[:-1]
        if word.endswith("'s") and word not in contractions:
            word = word[:-2]
        if word.endswith("in'"):
            word = word[:-1] + "g"
        if not word: #removing empty strings ("") that I created with the previous checks
            continue

        for subgenre in unique_subgenres:
            cursor.execute("""
                INSERT INTO words (word, {subgenre})
                VALUES (?, 1)
                ON CONFLICT(word) DO UPDATE SET {subgenre} = {subgenre} + 1
            """.format(subgenre=subgenre), (word,))


def print_words():
    cursor.execute("SELECT * FROM words")
    for row in cursor.fetchall():
        print(row)

# Loop through folder structure: BandName -> Year_AlbumName -> song.txt
for band_name in os.listdir(lyrics_folder_path):
    band_path = os.path.join(lyrics_folder_path, band_name)
    if not os.path.isdir(band_path):
        continue

    json_path = os.path.join(band_path, f"{band_name}.json")
    if not os.path.exists(json_path):
        print(f"No JSON found for {band_name}, skipping")
        continue
    
    with open(json_path, "r") as bandJSON:
        try:
            band_data = json.load(bandJSON)
        except json.JSONDecodeError as e:
            print(f"JSON error in {band_name}: {e}")
            continue

    country = band_data["country"]
    artist_id = get_or_create_artist(band_name, country)

    clean_path = os.path.join(band_path, "clean")
    if not os.path.exists(clean_path):
        print(f"No clean folder found for {band_name}, skipping")
        continue

    for release_folder in os.listdir(clean_path):
        release_path = os.path.join(clean_path, release_folder)
        if not os.path.isdir(release_path):
            continue

        if release_folder not in band_data["releases"]:
            print(f"No data found for {release_folder}, skipping")
            continue

        release_data = band_data["releases"][release_folder]
        year = release_data["year"]
        subgenres = release_data["subgenres"]
        release_type = release_data["type"]
        vocalist_gender = release_data.get("vocalist_gender", band_data.get("vocalist_gender", "unknown"))
        release_id = get_or_create_release(artist_id, band_name, release_folder, year, release_type, vocalist_gender)

        for subgenre in subgenres:
            get_or_create_release_subgenre(release_id, artist_id, release_folder, band_name, subgenre, vocalist_gender)

        for filename in os.listdir(release_path):
            if filename.endswith(".txt"):
                song_title = filename.replace(".txt", "")
                song_id = get_or_create_song(release_id, band_name, song_title)

                with open(os.path.join(release_path, filename), "r") as songLyrics:
                    lyrics = songLyrics.read()

                print(f"Processing: {band_name} - {release_folder} ({year}, {release_type}) - {song_title}")
                get_words(release_data["subgenres"], lyrics)
                conn.commit()


#print_words()


elapsed = time.time() - start_time
minutes = int(elapsed // 60)
seconds = int(elapsed % 60)
print(f"Total runtime: {minutes}m {seconds}s")

conn.close()
