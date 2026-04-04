#!/bin/bash
# setup.sh — Run once to set up the project on Mac/Linux
# Usage: bash setup.sh

echo ""
echo "🔧 Setting up mercedes-benz-uk-scraper..."
echo ""

# Create virtual environment
python3 -m venv venv
echo "✅ Virtual environment created (venv/)"

# Activate and install dependencies
source venv/bin/activate
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
echo "✅ Dependencies installed"

# Create folder structure
mkdir -p data/used/chunks data/new/chunks output
echo "✅ Folder structure created"

# Create .env from example if it doesn't exist
if [ ! -f .env ]; then
    cp .env.example .env
    echo "✅ .env file created — open it and paste your TOKEN and COOKIE"
else
    echo "ℹ️  .env already exists — skipping"
fi

echo ""
echo "─────────────────────────────────────────"
echo "✅ Setup complete!"
echo ""
echo "Next steps:"
echo "  1. Open .env and paste your MB_TOKEN and MB_COOKIE"
echo "  2. Activate the venv:  source venv/bin/activate"
echo "  3. Run:                python scrape_used.py"
echo "─────────────────────────────────────────"
echo ""