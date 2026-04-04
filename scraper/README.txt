╔══════════════════════════════════════════════════════════════════╗
║         MERCEDES-BENZ UK SCRAPER — PROJECT GUIDE                ║
╚══════════════════════════════════════════════════════════════════╝

────────────────────────────────────────────────────────────────────
FOLDER STRUCTURE
────────────────────────────────────────────────────────────────────

mercedes-benz-uk-scraper/
│
├── scrape_used.py          ← Scraper for USED cars (auto token refresh)
├── scrape_new.py           ← Scraper for NEW cars (auto token refresh)
├── to_excel.py             ← Converts JSON → Excel (works for both)
├── README.txt              ← This file
├── requirements.txt        ← Python dependencies
├── .gitignore              ← Tells GitHub what to ignore
├── .env.example            ← Template (safe to commit)
├── setup.bat               ← First-time setup (Windows)
├── setup.sh                ← First-time setup (Mac/Linux)
│
├── .env                    ← NOT in GitHub (gitignored)
├── venv/                   ← NOT in GitHub (gitignored)
│
├── data/
│   ├── used/
│   │   ├── chunks/                        ← Auto-filled while scraping
│   │   ├── progress.json                  ← Auto-created (resume tracking)
│   │   ├── credentials.json               ← Auto-created (refresh token)
│   │   └── mercedes_used_cars_FULL.json   ← Final output
│   └── new/
│       ├── chunks/                        ← Auto-filled while scraping
│       ├── progress.json                  ← Auto-created (resume tracking)
│       ├── credentials.json               ← Auto-created (refresh token)
│       └── mercedes_new_cars_FULL.json    ← Final output
│
└── output/
    ├── mercedes_used_vehicles.xlsx
    └── mercedes_new_vehicles.xlsx


────────────────────────────────────────────────────────────────────
HOW TO RUN
────────────────────────────────────────────────────────────────────

FIRST TIME ONLY:
  1. Open scrape_used.py (or scrape_new.py)
  2. Get a fresh token from DevTools (F12 → Network tab)
  3. Paste it into TOKEN = "Bearer ..."
  4. Run the script

AFTER FIRST RUN:
  The script saves credentials.json with the refresh token.
  Future runs auto-refresh the token — no DevTools needed.

COMMANDS:
  python scrape_used.py        ← scrape used cars
  python scrape_new.py         ← scrape new cars
  python to_excel.py used      ← convert used → Excel
  python to_excel.py new       ← convert new  → Excel


────────────────────────────────────────────────────────────────────
REQUIREMENTS
────────────────────────────────────────────────────────────────────

  pip install requests openpyxl


────────────────────────────────────────────────────────────────────
TROUBLESHOOTING
────────────────────────────────────────────────────────────────────

  Auto-refresh fails  →  Paste a fresh TOKEN into the script once
  403 on first run    →  Token expired before pasting — get a new one
  Missing JSON file   →  Run scraper first before to_excel.py