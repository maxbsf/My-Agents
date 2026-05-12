import requests
import schedule
import time
from datetime import datetime
from pathlib import Path
import openpyxl

CITIES = {
    "Hadera":   {"lat": 32.4342, "lon": 34.9196},
    "Netanya":  {"lat": 32.3337, "lon": 34.8601},
    "Tel Aviv": {"lat": 32.0853, "lon": 34.7818},
}

EXCEL_FILE = "temperatures.xlsx"


def fetch_temperature(lat: float, lon: float) -> float | None:
    url = (
        "https://api.open-meteo.com/v1/forecast"
        f"?latitude={lat}&longitude={lon}&current_weather=true"
    )
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return response.json()["current_weather"]["temperature"]
    except Exception as e:
        print(f"  Error fetching ({lat}, {lon}): {e}")
        return None


def get_or_create_workbook() -> tuple[openpyxl.Workbook, openpyxl.worksheet.worksheet.Worksheet]:
    path = Path(EXCEL_FILE)
    if path.exists():
        wb = openpyxl.load_workbook(path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Temperatures"
        headers = ["Timestamp"] + list(CITIES.keys())
        ws.append(headers)
        # Bold headers
        from openpyxl.styles import Font
        for cell in ws[1]:
            cell.font = Font(bold=True)
    return wb, ws


def collect_and_save():
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Collecting temperatures...")
    row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    for city, coords in CITIES.items():
        temp = fetch_temperature(coords["lat"], coords["lon"])
        display = f"{temp}°C" if temp is not None else "N/A"
        print(f"  {city}: {display}")
        row.append(temp)

    wb, ws = get_or_create_workbook()
    ws.append(row)
    wb.save(EXCEL_FILE)
    print(f"  Saved to {EXCEL_FILE}")


if __name__ == "__main__":
    print("Temperature monitoring agent started.")
    print(f"Cities: {', '.join(CITIES)}")
    print("Collecting every 30 minutes. Press Ctrl+C to stop.\n")

    collect_and_save()  # Run immediately on start

    schedule.every(30).minutes.do(collect_and_save)

    while True:
        schedule.run_pending()
        time.sleep(1)
