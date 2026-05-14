SOURCE_FILE = r"C:\Users\maxbi\SafeFields Technologies\Communication site - Documents\Product\Product Management\רמזור חשיפה שנתי ממוצע (for Amnon).xlsx"
SOURCE_SHEET = "Simplified"
OUTPUT_FILE = r"C:\Users\maxbi\OneDrive\Documents\GitHub\My-Agents\car-sales-agent\output\cars_tam.xlsx"

YEARS = [2020, 2021, 2022, 2023, 2024, 2025, 2026]

# MHEV is treated as HEV
ELECTRIFIED_TYPES = {"HEV", "PHEV", "BEV", "MHEV"}
TYPE_DISPLAY_MAP = {"MHEV": "HEV"}  # MHEV shown as HEV in output

SCRAPE_DELAY_SECONDS = 1.5  # polite delay between requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "he-IL,he;q=0.9,en-US;q=0.8,en;q=0.7",
}

CARZONE_BASE = "https://www.carzone.co.il"
CARZONE_SEARCH = "https://www.carzone.co.il/search"
STAT_ORG_BASE = "https://www.stat.org.il"
