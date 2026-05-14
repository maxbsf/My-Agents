"""
Scraper for Israeli car sales data from carzone.co.il.
URL pattern:
  - Current year (2026): https://www.carzone.co.il{model_path}
  - Historical years:     https://www.carzone.co.il{model_path}/{year}
Sales total is extracted from Hebrew text: "N ... נמכרו בשנת YYYY"
"""
import re
import time
import logging
from typing import Optional

import requests
from bs4 import BeautifulSoup

from config import HEADERS, SCRAPE_DELAY_SECONDS, CARZONE_BASE, YEARS

logger = logging.getLogger(__name__)

CURRENT_YEAR = max(YEARS)   # 2026

# ---------------------------------------------------------------------------
# Known URL paths keyed by row_num (matches source Excel column '#')
# Format: row_num → carzone path (without base URL or year suffix)
# ---------------------------------------------------------------------------
MODEL_PATHS: dict[int, str] = {
    1:  "/chery/Tiggo-8-Pro",             # Chery Tiggo8pro PHEV
    2:  "/Jeep/Wrangler",                 # Chrysler Wrangler 4xe (sold as Jeep in IL)
    3:  "/Omoda-Jaecoo/Jaecoo-8",         # Jaecoo 8 PHEV
    4:  "/MG/HS",                         # MG HS HYBRID+
    5:  "/Omoda-Jaecoo/Omoda-9",          # Omoda 9 PHEV
    6:  "/Hyundai/Elantra",               # Hyundai Elantra HEV
    7:  "/Omoda-Jaecoo/Omoda-7",          # Omoda 7 PHEV
    8:  "/Omoda-Jaecoo/Jaecoo-7",         # Jaecoo 7 PHEV
    9:  "/Omoda-Jaecoo/Jaecoo-5",         # Jaecoo 5 HEV
    10: "/MG/HS/EHS",                     # MG EHS PHEV
    11: "/Tesla/Model-3",                 # Tesla 3 BEV
    12: "/Hyundai/IONIQ",                # Hyundai Ioniq 4 HEV (original Ioniq HEV sold in IL)
    13: "/BYDAuto/Seal-U/DMi",            # BYD Seal U DM-I PHEV
    14: "/Porsche/Panamera",             # Porsche Panamera PHEV
    15: "/Suzuki/Vitara",                # Suzuki Vitara MHEV→HEV
    16: "/Suzuki/Swift",                 # Suzuki Swift micro-hybrid
    17: "/leapmotor/C10",               # LeapMotor C10 PHEV
    18: "/Audi/Q8/e-tron-Sportback",     # Audi Q8 Sportback e-tron BEV
    19: "/Kia/Niro/Plus",                # KIA Niro Gen 2 PHEV
    20: "/Kia/Niro/Plus",                # KIA Niro Plus Gen1 PHEV
    21: "/Toyota/Camry",                 # Toyota Camry HEV
    22: "/Chevrolet/Silverado",          # Chevrolet Silverado RST BEV
    23: "/Mitsubishi/Outlander",         # Mitsubishi Outlander PHEV
    24: "/Kia/Niro",                     # KIA Niro Hybrid Gen 1 HEV (same page)
    25: "/VOYAH/Courage",                # VOYAH Courage BEV
    26: "/Toyota/Corolla/4-doors",       # Toyota Corolla sedan HEV
    27: "/Toyota/RAV4",                  # Toyota RAV4 HEV
    28: "/Tesla/Model-Y",                # Tesla Y BEV
    29: "/BMW/X5",                       # BMW X5 xDrive45e PHEV
    30: "/MG/4",                         # MG 4 X-Range BEV
    31: "/SERES/5/M5",                   # Seres M5 BEV
    32: "/Hyundai/Kona",                 # Hyundai Kona HEV
    33: "/SERES/5/M5",                   # Seres M5 awd+rwd BEV (same page)
    34: "/Hyundai/Kona/EV",              # Hyundai Kona BEV
    35: "/BMW/5-Series",                 # BMW 530e PHEV
    36: "/Volvo/XC40",                   # Volvo XC40 PHEV
    37: "/Nissan/Juke",                  # Nissan Juke HEV
    38: "/LynkCo/01",                    # Lynk&Co 01 PHEV
    39: "/Geely/EX5",                    # Geely EX5 BEV
    40: "/Hyundai/Sonata",               # Hyundai Sonata HEV
    41: "/MG/Marvel-R",                  # MG Marvel R RWD BEV
    42: "/Toyota/Yaris",                 # Toyota Yaris HEV
    43: "/Toyota/Yaris-Cross",           # Toyota Yaris Cross HEV
    44: "/XPENG/G6",                     # Xpeng G6 BEV
    45: "/BYDAuto/Sea-Lion-07",          # BYD Sealion 7 BEV
    46: "/ZEEKR/X",                      # Zeekr X BEV
    47: "/Hyundai/Ioniq-5",              # Hyundai Ioniq 5 BEV
    48: "/BMW/X3/iX3",                   # BMW iX3 BEV
    49: "/BYDAuto/Atto-3",               # BYD Atto 3 BEV
    50: "/Honda/Accord",                 # Honda Accord HEV
    51: "/Skoda/ENYAQ-iv",               # Skoda Enyaq BEV
    52: "/Geely/Geometry-C",             # Geely Geometry C BEV
    53: "/Peugeot/208/e-208",            # Peugeot E208 BEV
    54: "/leapmotor/C10",               # LeapMotor C10 BEV (same page as PHEV)
    55: "/Volkswagen/ID.4",              # VW ID4 BEV
    56: "/MG/5",                         # MG 5 BEV
    57: "/MG/ZS-SUV",                    # MG ZS BEV
    58: "/EVeasy/Limo",                  # Eveasy Limo BEV
    59: "/SERES/E3",                     # Seres 3 (E3) BEV
    60: "/Toyota/C-HR",                  # Toyota CHR HEV
    61: "/Volvo/EX30",                   # Volvo EX30 BEV
    62: "/ZEEKR/001",                    # Zeekr 001 BEV
    63: "/ZEEKR/7X",                     # Zeekr 7x BEV
}

def _build_url(model_path: str, year: int) -> str:
    if year == CURRENT_YEAR:
        return f"{CARZONE_BASE}{model_path}"
    return f"{CARZONE_BASE}{model_path}/{year}"


def _fetch_page(url: str) -> Optional[str]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        resp.encoding = "utf-8"
        return resp.text
    except Exception as e:
        logger.warning(f"GET {url} → {e}")
        return None


def _extract_total(html: str, year: int) -> Optional[int]:
    """
    Find the sales total for the given year from a carzone page.

    Carzone page text contains a sentence like:
      "594 רכבי טויוטה קאמרי נמכרו בשנת 2020"
      "466 כ״מ טויוטה קאמרי נמכרו בשנת 2023"

    The sales count ALWAYS immediately precedes the Hebrew word for "vehicles"
    (רכבי / כ״מ / כ"מ).  We locate "נמכרו בשנת YEAR", then search backward
    in a 300-char window for the last occurrence of
      (NUMBER)(whitespace)(vehicle-word)
    This avoids false matches on model numbers embedded in the name
    (e.g. "Tiggo 8", "Tesla 3", "RAV4", "Ioniq 5").
    """
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(" ", strip=True)

    needle = f"נמכרו בשנת {year}"
    pos = text.find(needle)
    if pos == -1:
        return None

    window = text[max(0, pos - 300): pos]

    # רכבי = "vehicles of" (construct pl.)  כ"מ / כ״מ = abbreviation
    vehicle_words = r'(?:רכבי|כ["״]מ|מכוניות|רכב\s)'
    matches = list(re.finditer(rf'(\d[\d,]*)\s+{vehicle_words}', window))
    if matches:
        raw = matches[-1].group(1).replace(",", "")
        val = int(raw)
        return val if val > 0 else None

    # Fallback: last number in a short 60-char window (avoids model numbers
    # that tend to sit right before "נמכרו" in longer windows)
    short_window = text[max(0, pos - 60): pos]
    numbers = re.findall(r'\b(\d[\d,]*)\b', short_window)
    if numbers:
        raw = numbers[-1].replace(",", "")
        val = int(raw)
        return val if val > 0 else None

    return None


def get_sales_data(row_num: int) -> dict:
    """
    Fetch yearly sales {year: units|None} for a model identified by its row number.
    Returns dict covering all years in YEARS.
    """
    model_path = MODEL_PATHS.get(row_num)
    if not model_path:
        logger.warning(f"Row {row_num}: no URL mapping defined — skipping")
        return {year: None for year in YEARS}

    results = {}
    for year in YEARS:
        url = _build_url(model_path, year)
        time.sleep(SCRAPE_DELAY_SECONDS)
        html = _fetch_page(url)
        if html is None:
            logger.debug(f"  row {row_num} year {year}: 404/error → None")
            results[year] = None
            continue
        total = _extract_total(html, year)
        if total is not None:
            logger.debug(f"  row {row_num} year {year}: {total:,}")
        else:
            logger.debug(f"  row {row_num} year {year}: parsed None (no match in page)")
        results[year] = total

    found = sum(1 for v in results.values() if v is not None)
    logger.info(f"  → {found}/{len(YEARS)} years found  [{model_path}]")
    return results
