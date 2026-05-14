# Cars TAM Agent

Collects Israeli annual car sales data (2020–2026) for all electrified models listed in the SafeFields source file and produces a formatted Excel report.

## Models covered

- **HEV** – traditional full hybrids (+ MHEV mild hybrids treated as HEV)
- **PHEV** – plug-in hybrids
- **BEV** – battery electric vehicles

63 models total from the "Simplified" sheet of the source file.

## Setup

```bash
cd car-sales-agent
pip install -r requirements.txt
```

## Usage

### Full run (scrape all models)
```bash
python agent.py
```

### Dry run (build Excel without scraping — empty sales columns)
```bash
python agent.py --dry-run
```

### Single model test
```bash
python agent.py --model "Toyota Camry"
python agent.py --model "Tesla"
```

### Restart from scratch (ignore saved checkpoint)
```bash
python agent.py --clear-checkpoint
```

## Output

`output/cars_tam.xlsx` — copy of source sheet with new columns:

| Column | Content |
|--------|---------|
| Overall TAM | Excel SUM formula across 2020–2026 |
| 2026 (YTD) | Units sold Jan–present |
| 2025 | Annual units |
| 2024 | Annual units |
| 2023 | Annual units |
| 2022 | Annual units |
| 2021 | Annual units |
| 2020 | Annual units |

**Color coding:**
- Blue text = scraped sales figures (hardcoded inputs)
- Black text = Overall TAM formula

## Data sources

1. **Primary:** [carzone.co.il](https://www.carzone.co.il) — Israeli car marketplace with registration stats
2. **Fallback:** [stat.org.il](https://www.stat.org.il) — Israel Automobile Importers Association official data

## Checkpoint / Resume

The agent saves progress to `output/checkpoint.json` after each model. If the run is interrupted, re-running picks up where it left off automatically.

## Notes

- The original source file is **never modified**.
- Models with no Israeli sales data are left blank (not `0`).
- 2026 column is marked "YTD" (year-to-date).
- Scraping uses a ~1.5 s polite delay between requests.
