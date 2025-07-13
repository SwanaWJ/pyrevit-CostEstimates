import os
import requests
import pdfplumber
import csv
from bs4 import BeautifulSoup

# --- Step 1: Fetch ZPPA page and extract latest PDF link ---
zppa_url = "https://www.zppa.org.zm/news-item/-/journal_content/56/20182/98470"
response = requests.get(zppa_url)
soup = BeautifulSoup(response.text, "html.parser")

pdf_links = [
    a['href'] for a in soup.find_all('a', href=True)
    if a['href'].lower().endswith('.pdf')
]

if not pdf_links:
    print("❌ No PDF links found.")
    exit()

# Use the most recent PDF link (assuming it's the first one)
latest_pdf_url = pdf_links[0]
if not latest_pdf_url.startswith("http"):
    latest_pdf_url = "https://www.zppa.org.zm" + latest_pdf_url

print(f"📄 Downloading: {latest_pdf_url}")

pdf_response = requests.get(latest_pdf_url)
pdf_path = "latest_zppa_prices.pdf"
with open(pdf_path, "wb") as f:
    f.write(pdf_response.content)

# --- Step 2: Extract table data from the PDF ---
data = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        try:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                # Basic cleaning
                if len(row) < 3:
                    continue
                item = row[0].strip()
                unit = row[1].strip()
                price_str = row[2].strip().replace(",", "")
                try:
                    price = float(price_str)
                except:
                    continue
                data.append({"Item": item, "Unit": unit, "UnitCost": price})
        except Exception as e:
            print("⚠️ Error reading page:", e)
            continue

if not data:
    print("❌ No valid material-price data extracted.")
    exit()

# --- Step 3: Save to material_unit_costs.csv ---
output_csv = "material_unit_costs.csv"
with open(output_csv, "w", newline="") as csvfile:
    fieldnames = ["Item", "UnitCost"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for row in data:
        writer.writerow({"Item": row["Item"], "UnitCost": row["UnitCost"]})

print(f"✅ Extracted {len(data)} items and saved to {output_csv}")