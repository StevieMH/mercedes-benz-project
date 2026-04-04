# Mercedes-Benz Market Analysis 🚗

A web scraping and data visualization project that collects 
Mercedes-Benz vehicle listings (new and used) and analyzes 
them through an interactive Power BI dashboard.

---

## 📁 Project Structure

mercedes-benz-project/
├── scraper/
│   ├── scrape_new.py       # Scrapes new Mercedes-Benz listings
│   ├── scrape_used.py      # Scrapes used Mercedes-Benz listings
│   ├── to_excel.py         # Exports scraped data to Excel
│   ├── to_powerbi.py       # Formats data for Power BI
│   ├── requirements.txt    # Python dependencies
│   ├── setup.bat           # Windows setup script
│   └── setup.sh            # Linux/Mac setup script
└── dashboard/
    ├── mercedes_luxury_dashboard.pbix  # Power BI report
    ├── mercedes_luxury_theme.json      # Custom Power BI theme
    └── data/
        └── mercedes_powerbi.xlsx       # Processed data

---

## 🛠️ Technologies Used

- **Python** — core programming language
- **Scrapy** — web scraping framework
- **Pandas** — data processing
- **Power BI** — data visualization
- **Excel** — data storage

---

## 🚀 How to Run

**1. Clone the repository:**
git clone https://github.com/StevieMH/mercedes-benz-project.git

**2. Navigate to the scraper folder:**
cd mercedes-benz-project/scraper

**3. Install dependencies:**
pip install -r requirements.txt

**4. Run the scrapers:**
python scrape_new.py
python scrape_used.py

**5. Export data for Power BI:**
python to_powerbi.py

**6. Open the dashboard:**
Open dashboard/mercedes_luxury_dashboard.pbix in Power BI Desktop

---

## 📊 Dashboard Preview

> Open the .pbix file in Power BI Desktop to explore 
> the full interactive report.

---

## 👤 Author

**StevieMH**  
GitHub: https://github.com/StevieMH

---

## 📌 Notes

- Make sure to check .env.example for any required 
  environment variables before running the scrapers.
