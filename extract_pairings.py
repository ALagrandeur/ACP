import re
import fitz
import json
import sys

MONTH_MAP = {
    'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
    'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
}

def format_time(t):
    """Format time string like '1340' to '13h40'"""
    if not t or len(t) < 3:
        return t
    t = t.zfill(4)
    return f"{t[:2]}h{t[2:]}"

def time_to_minutes(t):
    """Convert time string like '1340' to total minutes (13*60+40=840)"""
    if not t or not t.strip():
        return 0
    t = t.strip().zfill(4)
    try:
        h = int(t[:-2])
        m = int(t[-2:])
        return h * 60 + m
    except ValueError:
        return 0

def parse_block(block_text):
    lines = block_text.strip().split('\n')
    lines = [l for l in lines if l.strip()]

    if not lines:
        return None

    # Find pairing ID and dates
    pairing_id = None
    date_start = None
    date_end = None
    for line in lines:
        m = re.search(r'(M\d{4})\s+OPERATES/OPER-\s+(\d{2})([A-Z]{3})\s+-\s+(\d{2})([A-Z]{3})', line)
        if m:
            pairing_id = m.group(1)
            date_start = f"{m.group(2)}{m.group(3)}"
            date_end = f"{m.group(4)}{m.group(5)}"
            break

    if not pairing_id:
        return None

    # Parse crew and languages
    crew = ""
    languages = ""
    for line in lines:
        if re.search(r'(?:P\s+\d{2}|FA\d{2})', line) and 'OPERATES' not in line and 'FREQ' not in line:
            parts = re.split(r'\s{5,}', line.strip())
            crew_parts = []
            lang_parts = []
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                if re.match(r'^[A-Z]{2}\d{2}(\s+[A-Z]{2}\d{2})*$', part):
                    lang_parts.append(part)
                elif re.search(r'(?:P\s+\d{2}|FA\d{2}|GJ\d{2}|GY\d{2})', part):
                    crew_parts.append(part)
            crew = ' '.join(crew_parts)
            languages = ' '.join(lang_parts)
            break

    # Parse flight legs
    legs = []
    hotels = []
    destinations = set()
    equipment_set = set()
    layover_cities = []

    for i, line in enumerate(lines):
        # Skip header/metadata lines
        if 'OPERATES' in line or 'FREQ' in line or 'BLOCK/H-VOL' in line or 'TAFB/PTEB' in line:
            continue
        if 'TOTAL ALLOWANCE' in line:
            continue
        if re.match(r'^\s*-{5,}', line):
            continue
        if 'Su Mo Tu' in line or 'Di Lu Ma' in line:
            continue

        # Try to match flight leg: look for CITY TIME CITY TIME pattern
        m = re.search(r'([A-Z]{3})\s+(\d{4})\s+([A-Z]{3})\s+(\d{4})\s+(\d{2,4})', line)
        if m and 'DEPART' not in line and 'ARRIVEE' not in line and 'ARRIVE' not in line:
            dep_city = m.group(1)
            dep_time = m.group(2)
            arr_city = m.group(3)
            arr_time = m.group(4)
            block = m.group(5)

            # Extract prefix: frequency + equipment + flight number
            prefix = line[:m.start()]
            freq = ""
            equip = ""
            is_dh = False
            flt = ""

            # Parse prefix
            prefix_m = re.match(r'^\s*([\d ]{1,10}?)\s+(\w{2,7}?)(DHD)?\s+(\d{1,5})\s*$', prefix)
            if prefix_m:
                freq = prefix_m.group(1).strip()
                equip = prefix_m.group(2)
                is_dh = prefix_m.group(3) == 'DHD'
                flt = prefix_m.group(4)
            else:
                # Try alternate parsing
                prefix_m2 = re.match(r'^\s*([\d ]*?)\s+(\w+?)(DHD)\s+(\d{1,5})\s*$', prefix)
                if prefix_m2:
                    freq = prefix_m2.group(1).strip()
                    equip = prefix_m2.group(2)
                    is_dh = True
                    flt = prefix_m2.group(4)
                else:
                    prefix_m3 = re.match(r'^\s*([\d ]*)\s+(\d{3}[A-Z]?|[A-Z0-9]{2,4})\s+(\d{1,5})\s*$', prefix)
                    if prefix_m3:
                        freq = prefix_m3.group(1).strip()
                        equip = prefix_m3.group(2)
                        flt = prefix_m3.group(3)

            # Extract suffix: duty time, layover, meals
            suffix = line[m.end():]
            duty = ""
            layover = ""
            meals = ""

            # Parse duty time and layover from suffix
            suffix_nums = re.findall(r'(\d{2,5})', suffix)
            if suffix_nums:
                duty = suffix_nums[0]
                if len(suffix_nums) > 1:
                    layover = suffix_nums[1]

            # Extract meals from end of line
            meals_m = re.findall(r'\b(HB|HD|HL|SS|PP|[BLD])\b', suffix)
            meals = ' '.join(meals_m)

            leg = {
                'frequency': freq,
                'equipment': equip,
                'deadhead': is_dh,
                'flight': flt,
                'dep_city': dep_city,
                'dep_time': dep_time,
                'arr_city': arr_city,
                'arr_time': arr_time,
                'block_time': block,
                'duty_time': duty,
                'layover': layover,
                'meals': meals,
            }
            legs.append(leg)

            if dep_city != 'YUL':
                destinations.add(dep_city)
            if arr_city != 'YUL':
                destinations.add(arr_city)
            if equip:
                equipment_set.add(equip)

            # Track layover cities (arrival city when there's a layover)
            if layover:
                layover_cities.append(arr_city)

        # Check for hotel lines
        hotel_patterns = [
            'Hilton', 'Marriott', 'Sheraton', 'Westin', 'Hyatt', 'Novotel',
            'Radisson', 'Delta Hotels', 'Delta Calgary', 'Alt Hotel', 'Grand Hyatt',
            'Hoxton', 'The Hoxton', 'Sandman', 'NH Collection', 'Hotel NH',
            'The Westin', 'Hampton', 'Holiday Inn', 'DoubleTree', 'Crowne Plaza',
            'InterContinental', 'Fairmont', 'Four Points', 'Courtyard',
            'Best Western', 'Quality', 'Comfort', 'Days Inn', 'Le Germain',
            'Sofitel', 'Pullman', 'Stanford Court', 'W Hotel', 'Renaissance',
            'Residence Inn', 'SpringHill', 'AC Hotel', 'JW Marriott',
            'Embassy Suites', 'Homewood', 'Canopy', 'Curio', 'Tapestry',
            'DoubleTree', 'Tru by', 'Moxy', 'Aloft', 'Element',
        ]
        for hp in hotel_patterns:
            if hp.lower() in line.lower():
                # Extract the hotel name (from the hotel keyword to end of meaningful text)
                hotel_m = re.search(
                    r'(' + re.escape(hp) + r'[\w\s&\'-/,.]*)',
                    line, re.IGNORECASE
                )
                if hotel_m:
                    hotel_name = hotel_m.group(1).strip()
                    # Clean up trailing codes
                    hotel_name = re.sub(r'\s+(HND|DT|[A-Z]{2,3}\s+DT)\s*$', '', hotel_name)
                    hotel_name = re.sub(r'\s{2,}.*$', '', hotel_name)
                    if hotel_name and hotel_name not in hotels:
                        hotels.append(hotel_name)
                break

        # Check for DPG line (deadhead passenger time with hotel)
        dpg_m = re.search(r'(\d{2,4})\s*-DPG', line)
        if dpg_m and not m:  # Don't double-count if already matched as flight leg
            # This line may also contain hotel info (already handled above)
            pass

    # Parse summary
    block_total = ""
    tafb = ""
    allowance = 0.0
    dpg_total = ""
    cico = ""

    for line in lines:
        m = re.search(r'BLOCK/H-VOL\s+(\d+)', line)
        if m:
            block_total = m.group(1)
        m = re.search(r'TAFB/PTEB\s+(\d+)', line)
        if m:
            tafb = m.group(1)
        m = re.search(r'TOTAL ALLOWANCE -\$\s*([\d,.]+)', line)
        if m:
            allowance = float(m.group(1).replace(',', ''))
        m = re.search(r'DPG\s*-\s*(\d+)', line)
        if m:
            dpg_total = m.group(1)
        m = re.search(r'THG\s*-\s*(\d+)', line)
        if m:
            dpg_total = m.group(1)  # THG is similar concept
        m = re.search(r'TOTAL\s*-\s*(\d+)', line)
        if m and 'ALLOWANCE' not in line:
            pass  # This is the total flight time including DPG

    # Parse calendar operating dates
    operating_dates = []
    in_calendar = False
    calendar_line_count = 0
    for line in lines:
        if 'Su Mo Tu We Th Fr Sa' in line:
            in_calendar = True
            calendar_line_count = 0
            continue
        if 'Di Lu Ma Me Je Ve Sa' in line:
            continue
        if in_calendar:
            calendar_line_count += 1
            if calendar_line_count > 6:
                in_calendar = False
                continue
            # Extract date numbers from calendar row
            # Each cell is either '--' or a number (1-31)
            cells = re.findall(r'(\d{1,2}|--)', line)
            for cell in cells:
                if cell != '--':
                    try:
                        d = int(cell)
                        if 1 <= d <= 31:
                            operating_dates.append(d)
                    except ValueError:
                        pass

    # Calculate number of days (total calendar days away from base)
    # Use frequency digits: first digit of first leg = departure day,
    # first digit of last leg = return day. Difference + 1 = calendar days.
    def get_first_freq_digit(freq_str):
        for c in str(freq_str):
            if c.isdigit():
                return int(c)
        return None

    num_days = 1
    if legs:
        first_day = get_first_freq_digit(legs[0].get('frequency', ''))
        last_day = get_first_freq_digit(legs[-1].get('frequency', ''))
        if first_day and last_day:
            if last_day >= first_day:
                num_days = last_day - first_day + 1
            else:
                # Wraps around the week (e.g., Sat=6 -> Mon=1)
                num_days = (7 - first_day + last_day) + 1

    # Determine base city (always YUL for this document but let's extract it)
    base = 'YUL'
    if legs:
        base = legs[0].get('dep_city', 'YUL')

    pairing = {
        'id': pairing_id,
        'dateStart': date_start,
        'dateEnd': date_end,
        'crew': crew,
        'languages': languages,
        'legs': legs,
        'hotels': hotels,
        'destinations': sorted(list(destinations)),
        'equipment': sorted(list(equipment_set)),
        'blockTotal': block_total,
        'blockMinutes': time_to_minutes(block_total),
        'tafb': tafb,
        'tafbMinutes': time_to_minutes(tafb),
        'allowance': allowance,
        'operatingDates': sorted(list(set(operating_dates))),
        'numDays': num_days,
        'base': base,
        'layoverCities': layover_cities,
    }

    return pairing


def extract_pairings(pdf_path):
    doc = fitz.open(pdf_path)

    # Concatenate all pages, removing headers
    full_text = ""
    for page in doc:
        text = page.get_text()
        lines = text.split('\n')
        filtered = []
        for line in lines:
            if 'Produced' in line and 'Page No' in line:
                continue
            if re.match(r'^\s*FREQ\s+(?:APP|EQP)', line):
                continue
            if re.match(r'^\s*----\s+---', line):
                continue
            filtered.append(line)
        full_text += '\n'.join(filtered) + '\n'

    # Split into blocks by the separator
    blocks = re.split(r'={30,}', full_text)

    pairings = []
    errors = []
    for i, block in enumerate(blocks):
        if not block.strip():
            continue
        try:
            pairing = parse_block(block)
            if pairing:
                pairings.append(pairing)
        except Exception as e:
            errors.append(f"Block {i}: {str(e)}")

    return pairings, errors


def detect_month_from_pdf(pdf_path):
    """Detect the month name and year from the PDF filename or content."""
    import os
    filename = os.path.splitext(os.path.basename(pdf_path))[0]

    MONTH_NAMES_FR = {
        'janvier': ('Janvier', '01'), 'fevrier': ('Fevrier', '02'), 'mars': ('Mars', '03'),
        'avril': ('Avril', '04'), 'mai': ('Mai', '05'), 'juin': ('Juin', '06'),
        'juillet': ('Juillet', '07'), 'aout': ('Aout', '08'), 'septembre': ('Septembre', '09'),
        'octobre': ('Octobre', '10'), 'novembre': ('Novembre', '11'), 'decembre': ('Decembre', '12'),
    }
    MONTH_NAMES_EN = {
        'january': ('January', '01'), 'february': ('February', '02'), 'march': ('March', '03'),
        'april': ('April', '04'), 'may': ('May', '05'), 'june': ('June', '06'),
        'july': ('July', '07'), 'august': ('August', '08'), 'september': ('September', '09'),
        'october': ('October', '10'), 'november': ('November', '11'), 'december': ('December', '12'),
    }

    all_months = {**MONTH_NAMES_FR, **MONTH_NAMES_EN}
    fn_lower = filename.lower()

    month_name = None
    month_code = None
    year = None

    for key, (name, code) in all_months.items():
        if key in fn_lower:
            month_name = name
            month_code = code
            break

    year_match = re.search(r'(20\d{2})', filename)
    if year_match:
        year = year_match.group(1)

    if not year:
        # Try to detect from PDF content
        doc = fitz.open(pdf_path)
        first_page = doc[0].get_text()
        year_match = re.search(r'(\d{2})/(\d{2})/(\d{2})', first_page)
        if year_match:
            year = '20' + year_match.group(3)
        doc.close()

    if not month_name:
        month_name = filename
        month_code = '01'
    if not year:
        year = '2026'

    return month_name, month_code, year


def find_latest_pdf(pairing_dir):
    """Find the most recently modified PDF in the Pairing directory."""
    import os
    import glob
    pdfs = glob.glob(os.path.join(pairing_dir, '*.pdf'))
    if not pdfs:
        return None
    return max(pdfs, key=os.path.getmtime)


if __name__ == '__main__':
    import os
    import glob

    base_dir = os.path.dirname(os.path.abspath(__file__))
    pairing_dir = os.path.join(base_dir, 'Pairing')
    data_dir = os.path.join(base_dir, 'data')

    # Accept PDF path as argument, or find the latest PDF
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = find_latest_pdf(pairing_dir)
        if not pdf_path:
            print("ERREUR: Aucun fichier PDF trouve dans le dossier Pairing/")
            print(f"Placez votre PDF de pairing dans: {pairing_dir}")
            sys.exit(1)

    print(f"Extraction des pairings depuis: {pdf_path}")

    # Detect month
    month_name, month_code, year = detect_month_from_pdf(pdf_path)
    print(f"Mois detecte: {month_name} {year}")

    pairings, errors = extract_pairings(pdf_path)

    print(f"Pairings extraits: {len(pairings)}")
    if errors:
        print(f"Erreurs: {len(errors)}")
        for e in errors[:10]:
            print(f"  {e}")

    # Get unique destinations, equipment, languages for filter metadata
    all_destinations = set()
    all_equipment = set()
    all_languages = set()
    for p in pairings:
        all_destinations.update(p['destinations'])
        all_equipment.update(p['equipment'])
        if p['languages']:
            for lang in p['languages'].split():
                code = re.match(r'([A-Z]{2})', lang)
                if code:
                    all_languages.add(code.group(1))

    metadata = {
        'month': f'{month_name} {year}',
        'monthCode': f'{year}-{month_code}',
        'base': 'YUL',
        'totalPairings': len(pairings),
        'destinations': sorted(list(all_destinations)),
        'equipment': sorted(list(all_equipment)),
        'languages': sorted(list(all_languages)),
        'pdfFile': os.path.basename(pdf_path),
    }

    # Write as JavaScript module (active month)
    os.makedirs(data_dir, exist_ok=True)
    output_path = os.path.join(data_dir, 'pairings.js')

    output = f"var PAIRINGS_METADATA = {json.dumps(metadata, indent=2)};\n\n"
    output += f"var PAIRINGS_DATA = {json.dumps(pairings, indent=2)};\n"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(output)

    # Save per-month JSON file
    month_slug = f"{metadata['monthCode']}"  # e.g. "2026-04"
    json_path = os.path.join(data_dir, f"{month_slug}.json")
    json_data = {
        'metadata': metadata,
        'pairings': pairings
    }
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False)
    print(f"Fichier JSON mois: {json_path}")

    # Update months manifest
    manifest_path = os.path.join(data_dir, 'months.json')
    manifest = []
    if os.path.exists(manifest_path):
        try:
            with open(manifest_path, 'r', encoding='utf-8') as f:
                manifest = json.load(f)
        except Exception:
            manifest = []

    # Update or add this month
    existing = [m for m in manifest if m['monthCode'] != metadata['monthCode']]
    existing.append({
        'monthCode': metadata['monthCode'],
        'label': metadata['month'],
        'file': f"{month_slug}.json",
        'totalPairings': len(pairings)
    })
    existing.sort(key=lambda m: m['monthCode'])
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)
    print(f"Manifeste mis a jour: {manifest_path}")

    print(f"\nFichier genere: {output_path}")
    print(f"Destinations: {sorted(list(all_destinations))}")
    print(f"Equipment: {sorted(list(all_equipment))}")
    print(f"Langues: {sorted(list(all_languages))}")
    print(f"\nMise a jour terminee avec succes!")
