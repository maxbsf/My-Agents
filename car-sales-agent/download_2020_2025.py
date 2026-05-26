import sys, os
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout.reconfigure(encoding='utf-8')

import json
import time
import requests

RESOURCE_ID = '602ac32d-19c0-4b41-88e0-e3ce8a7e80b7'
BASE_URL = 'https://data.gov.il/api/3/action/datastore_search'
LIMIT = 10000
OUT_FILE = 'output/vehicle_2020_2025_raw.json'

# Binary search established: offset ~121,000 is near Jan 2020
# We'll start at 120,000 and collect everything through 202512

headers = {
    'User-Agent': 'Mozilla/5.0 (compatible; research-bot/1.0)',
    'Accept': 'application/json',
}

records = []
offset = 120000  # start just before Jan 2020
total_fetched = 0
found_start = False

print('Downloading passenger car records 2020-2025 from data.gov.il...')

while True:
    try:
        resp = requests.get(
            BASE_URL,
            params={
                'resource_id': RESOURCE_ID,
                'limit': LIMIT,
                'offset': offset,
            },
            headers=headers,
            timeout=60
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        print(f'  Error at offset {offset}: {e}')
        time.sleep(5)
        continue

    batch = data['result']['records']
    if not batch:
        print('No more records.')
        break

    first_month = batch[0].get('sgira_month', 0)
    last_month  = batch[-1].get('sgira_month', 0)

    # Filter: passenger cars only, months 202001-202512
    kept = [
        r for r in batch
        if r.get('sug_degem') == 'P'
        and 202001 <= int(r.get('sgira_month', 0)) <= 202512
    ]
    records.extend(kept)
    total_fetched += len(batch)

    print(f'  offset={offset:7d} | batch={len(batch):5d} | months {first_month}-{last_month} | kept={len(kept)} | total_kept={len(records)}')

    # Stop once we've passed December 2025
    if int(last_month) > 202512:
        print('Passed Dec 2025, stopping.')
        break

    offset += LIMIT
    time.sleep(0.3)

print(f'\nTotal passenger car records 2020-2025: {len(records)}')

with open(OUT_FILE, 'w', encoding='utf-8') as f:
    json.dump(records, f, ensure_ascii=False)
print(f'Saved: {OUT_FILE}')
