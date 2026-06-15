import csv
import io
import re
from thefuzz import fuzz
from collections import defaultdict


current_data = {
    "poles": [],
    "all_poles": [],
    "pending_duplicates": [],
    "confirmed_duplicates": [],
    "spelling_issues": [],
    "barangays": [],
    "municipalities": [],
    "poles_by_barangay": {},
    "poles_by_municipality": {},
    "total_poles": 0,
}


def parse_csv(file_bytes):
    content = file_bytes.decode('utf-8-sig')
    reader  = csv.DictReader(io.StringIO(content))
    return [
        row for row in reader
        if any(v.strip() for v in row.values() if v)
    ]


def is_close_gps(lat1, lon1, lat2, lon2, threshold=0.0005):
    try:
        return (
            abs(float(lat1) - float(lat2)) <= threshold and
            abs(float(lon1) - float(lon2)) <= threshold
        )
    except:
        return False


# Placeholder pole numbers that should be ignored
PLACEHOLDER_POLE_NUMBERS = {
    'npt', 'notb', 'none', 'n/a', 'null', 'na', 'n.a', 'n.a.', '-', '',
    'npy', 'nbt', 'ntp', 'nptt', 'not', 'notag', 'no tag', 'notag',
    'unknown', 'unk', 'tbd', 'n.t', 'nt'
}

# Valid placeholder values that are accepted as-is
VALID_PLACEHOLDERS = {'npt', 'notb', 'none', 'n/a', 'null', 'na', 'npy', 'nbt'}

# Common placeholder typos to detect
PLACEHOLDER_TYPO_THRESHOLD = 80  # fuzzy match score

def detect_pole_number_typos(poles):
    """
    Detect pole numbers that look like placeholder typos.
    Returns list of suspected typos with suggestions.
    """
    from thefuzz import fuzz
    
    typos    = []
    seen     = set()
    
    for pole in poles:
        raw = pole.get('pole_number', '').strip()
        if not raw:
            continue
            
        norm = normalize_pole_number(raw)
        
        # Skip if already a valid placeholder
        if norm in VALID_PLACEHOLDERS:
            continue
            
        # Skip if already seen
        if norm in seen:
            continue
        seen.add(norm)
        
        # Check similarity to each valid placeholder
        for valid in VALID_PLACEHOLDERS:
            score = fuzz.ratio(norm, valid)
            if score >= PLACEHOLDER_TYPO_THRESHOLD and norm != valid:
                # Count how many poles have this value
                count = sum(
                    1 for p in poles
                    if normalize_pole_number(p.get('pole_number', '')) == norm
                )
                typos.append({
                    'original':    raw,
                    'normalized':  norm,
                    'similar_to':  valid.upper(),
                    'score':       score,
                    'count':       count,
                    'suggestions': [valid.upper(), 'Keep as entered']
                })
                break
                
    return typos

def normalize_pole_number(pole_number):
    """Normalize pole number for comparison."""
    if not pole_number:
        return ''
    # lowercase, remove spaces and dashes
    normalized = pole_number.lower().strip()
    normalized = re.sub(r'[\s\-_]+', '', normalized)
    return normalized

def is_placeholder(pole_number):
    """Check if pole number is a placeholder value."""
    normalized = normalize_pole_number(pole_number)
    return normalized in PLACEHOLDER_POLE_NUMBERS

def find_duplicates(poles):
    """
    A pole is a duplicate if ANY of these are true:
    1. Same normalized pole number (excluding placeholders)
    2. Same loc_lat AND loc_long (exact match)
    3. Same photoAttachments URL

    Keeps FIRST occurrence (earliest date after sorting).
    Returns tuple: (clean_poles, duplicates)
    """
    seen_pole_numbers  = {}  # normalized pole number -> first pole
    seen_gps           = {}  # (lat, lon) -> first pole
    seen_photo         = {}  # photoAttachments URL -> first pole

    clean_poles = []
    duplicates  = []

    def get_field(pole, *field_names):
        """Try multiple field name variants."""
        for field in field_names:
            val = pole.get(field, '').strip()
            if val:
                return val
        return ''

    for pole in poles:
        # Get pole number — try multiple field name variants
        raw_pole_number = get_field(pole, 'pole_number', 'pole_num', 
                                    'polenumber', 'polenum')
        norm_pole       = normalize_pole_number(raw_pole_number)

        # Get GPS
        lat   = get_field(pole, 'loc_lat', 'latitude', 'lat')
        lon   = get_field(pole, 'loc_long', 'loc_lng', 'longitude', 'lon', 'lng')
        gps   = (lat, lon) if lat and lon else None

        # Get photo URL
        photo = get_field(pole, 'photoAttachments', 'photo_attachments',
                         'photoattachments')

        is_dup        = False
        matched_with  = None
        matched_reason = ''

        # Check 1: Same normalized pole number (skip placeholders)
        if norm_pole and not is_placeholder(raw_pole_number):
            if norm_pole in seen_pole_numbers:
                is_dup         = True
                matched_with   = seen_pole_numbers[norm_pole]
                matched_reason = f"Same pole number: {raw_pole_number}"

        # Check 2: Same GPS coordinates
        if not is_dup and gps and gps[0] and gps[1]:
            if gps in seen_gps:
                is_dup         = True
                matched_with   = seen_gps[gps]
                matched_reason = f"Same GPS: {lat}, {lon}"

        # Check 3: Same photoAttachments URL
        if not is_dup and photo:
            if photo in seen_photo:
                is_dup         = True
                matched_with   = seen_photo[photo]
                matched_reason = f"Same photo URL"

        if is_dup and matched_with:
            dup_entry = dict(pole)
            dup_entry['_matched_with']        = matched_with.get('poleid', '')
            dup_entry['_matched_pole_number'] = matched_with.get('pole_number', '')
            dup_entry['_matched_date']        = matched_with.get('created_at', '')
            dup_entry['_matched_reason']      = matched_reason
            duplicates.append(dup_entry)
        else:
            # Add to seen only if not duplicate
            if norm_pole and not is_placeholder(raw_pole_number):
                if norm_pole not in seen_pole_numbers:
                    seen_pole_numbers[norm_pole] = pole

            if gps and gps[0] and gps[1]:
                if gps not in seen_gps:
                    seen_gps[gps] = pole

            if photo:
                if photo not in seen_photo:
                    seen_photo[photo] = pole

            clean_poles.append(pole)

    return clean_poles, duplicates

def find_spelling_issues(poles):
    issues = []

    for field in ['barangay', 'municipality']:
        values = list(set(
            p.get(field, '').strip()
            for p in poles
            if p.get(field, '').strip()
        ))

        counts = defaultdict(int)
        for p in poles:
            v = p.get(field, '').strip()
            if v:
                counts[v] += 1

        checked = set()
        for i, val1 in enumerate(values):
            for val2 in values[i+1:]:
                pair = tuple(sorted([val1, val2]))
                if pair in checked:
                    continue
                checked.add(pair)

                score = fuzz.ratio(val1.lower(), val2.lower())

                DIRECTIONS = {
                    'east', 'west', 'north', 'south',
                    'upper', 'lower', 'inner', 'outer',
                    'norte', 'sur', 'silangan', 'kanluran',
                    'itaas', 'ibaba', 'proper', 'central',
                    'old', 'new', 'i', 'ii', 'iii', 'iv', 'v',
                    '1', '2', '3', '4', '5',
                }

                def extract_qualifiers(name):
                    words = set(name.lower().split())
                    return words & DIRECTIONS

                if score >= 75 and val1.lower() != val2.lower():
                    if extract_qualifiers(val1) != extract_qualifiers(val2):
                        continue
                    if counts[val1] >= counts[val2]:
                        suspect    = val2
                        suggestion = val1
                    else:
                        suspect    = val1
                        suggestion = val2

                    already = any(
                        x['value'] == suspect and x['field'] == field
                        for x in issues
                    )
                    if not already:
                        issues.append({
                            'field':       field,
                            'value':       suspect,
                            'count':       counts[suspect],
                            'suggestions': [suggestion, suspect],
                            'score':       score,
                        })

    return issues


def sort_poles(poles):
    def parse_date(pole):
        try:
            from datetime import datetime
            date_str = pole.get('created_at', '')
            for fmt in ['%d/%m/%Y %H:%M', '%Y-%m-%d %H:%M:%S',
                        '%m/%d/%Y %H:%M', '%d/%m/%Y %H:%M:%S']:
                try:
                    return datetime.strptime(date_str.strip(), fmt)
                except:
                    continue
            return date_str
        except:
            return ''
    return sorted(poles, key=parse_date)


def group_by_field(poles, field):
    groups = defaultdict(list)
    for pole in poles:
        key = pole.get(field, '').strip()
        if key:
            groups[key].append(pole)
    return dict(groups)


def regroup(poles):
    by_barangay     = group_by_field(poles, 'barangay')
    by_municipality = group_by_field(poles, 'municipality')
    return {
        "barangays":              sorted(by_barangay.keys()),
        "municipalities":         sorted(by_municipality.keys()),
        "poles_by_barangay":      {k: len(v) for k, v in by_barangay.items()},
        "poles_by_municipality":  {k: len(v) for k, v in by_municipality.items()},
    }


def analyze_csv(file_bytes):
    global current_data

    poles = parse_csv(file_bytes)
    poles = sort_poles(poles)

    # Don't remove duplicates yet — just find them
    spelling    = find_spelling_issues(poles)
    pole_typos  = detect_pole_number_typos(poles)

    # Find duplicates for review only — do NOT remove them yet
    _, pending_duplicates = find_duplicates(poles)

    # All poles stay in clean data until user explicitly confirms removal
    groups = regroup(poles)

    current_data = {
        "poles":                  poles,
        "clean_poles":            poles,
        "all_poles":              poles,
        "pending_duplicates":     pending_duplicates,
        "confirmed_duplicates":   [],
        "spelling_issues":        spelling,
        "pole_typos":             pole_typos,
        "total_poles":            len(poles),
        **groups,
    }

    # Add auto-suggest switch flags to duplicates
    orig_lookup = { p.get('poleid'): p for p in poles }
    for dup in pending_duplicates:
        matched_poleid = dup.get('_matched_with', '')
        orig_pole      = orig_lookup.get(matched_poleid, {})
        dup['_suggest_switch'] = suggest_switch(orig_pole, dup)

    return {
        "total_poles":        len(poles),
        "poles":              poles,
        "duplicates":         pending_duplicates,
        "spelling_issues":    spelling,
        "pole_typos":         pole_typos,
        **groups,
    }


def apply_fixes(fixes):
    global current_data

    # Apply spelling fixes to ALL poles
    for fix in fixes:
        field        = fix['field']
        original     = fix['original']
        replacement  = fix['replacement']
        municipality = fix.get('municipality')  # optional — move barangay to different municipality

        if original == replacement and not municipality:
            continue

        for pole in current_data['poles']:
            if pole.get(field, '').strip() == original:
                if original != replacement:
                    pole[field] = replacement
                # If municipality override provided, reassign it
                if municipality:
                    pole['municipality'] = municipality

    # Re-find duplicates after spelling fix — review only, no removal
    _, pending_duplicates = find_duplicates(current_data['poles'])
    groups = regroup(current_data['poles'])

    current_data.update({
        "clean_poles":        current_data['poles'],
        "pending_duplicates": pending_duplicates,
        "spelling_issues":    [],
        **groups,
    })

    orig_lookup = { p.get('poleid'): p for p in current_data['poles'] }
    for dup in pending_duplicates:
        matched_poleid = dup.get('_matched_with', '')
        orig_pole      = orig_lookup.get(matched_poleid, {})
        dup['_suggest_switch'] = suggest_switch(orig_pole, dup)

    return {
        "total_poles":     len(current_data['poles']),
        "poles":           current_data['poles'],
        "duplicates":      pending_duplicates,
        "spelling_issues": current_data.get('spelling_issues', []),
        **groups,
    }

def apply_pole_typo_fixes(fixes):
    """Apply pole number typo fixes."""
    global current_data

    for fix in fixes:
        original    = fix['original']
        replacement = fix['replacement']

        if replacement == 'Keep as entered' or original == replacement:
            continue

        for pole in current_data['poles']:
            if pole.get('pole_number', '').strip() == original:
                pole['pole_number'] = replacement

    # Re-run duplicate detection after fixes — review only, no removal
    _, pending_duplicates = find_duplicates(current_data['poles'])
    groups = regroup(current_data['poles'])

    current_data.update({
        "clean_poles":        current_data['poles'],
        "pending_duplicates": pending_duplicates,
        "pole_typos":         [],
        **groups,
    })

    orig_lookup = { p.get('poleid'): p for p in current_data['poles'] }
    for dup in pending_duplicates:
        matched_poleid = dup.get('_matched_with', '')
        orig_pole      = orig_lookup.get(matched_poleid, {})
        dup['_suggest_switch'] = suggest_switch(orig_pole, dup)

    return {
        "total_poles":     len(current_data['poles']),
        "poles":           current_data['poles'],
        "duplicates":      pending_duplicates,
        "spelling_issues": current_data.get('spelling_issues', []),
        **groups,
    }

def apply_single_pole_fix(poleid, new_pole_number):
    """Fix pole number for a single pole by poleid only."""
    global current_data

    for pole in current_data['poles']:
        if pole.get('poleid', '') == poleid:
            pole['pole_number'] = new_pole_number

    _, pending_duplicates = find_duplicates(current_data['poles'])
    groups = regroup(current_data['poles'])

    current_data.update({
        "clean_poles":        current_data['poles'],
        "pending_duplicates": pending_duplicates,
        **groups,
    })

    return {
        "total_poles":     len(current_data['poles']),
        "poles":           current_data['poles'],
        "duplicates":      pending_duplicates,
        "spelling_issues": current_data.get('spelling_issues', []),
        "pole_typos":      [],
        **groups,
    }

def is_weak_pole_number(pole_number):
    """Returns True if pole number is blank, placeholder, or weak."""
    if not pole_number:
        return True
    normalized = normalize_pole_number(pole_number)
    return normalized in PLACEHOLDER_POLE_NUMBERS


def suggest_switch(original_pole, duplicate_pole):
    """
    Returns True if the duplicate looks better than the original.
    Triggered when original has weak pole number but duplicate has a real one.
    """
    orig_weak = is_weak_pole_number(original_pole.get('pole_number', ''))
    dup_weak  = is_weak_pole_number(duplicate_pole.get('pole_number', ''))
    return orig_weak and not dup_weak

def confirm_duplicates(poleid_list, switches=None):
    """
    poleid_list: poleids to remove
    switches: list of { remove_poleid, keep_poleid } overrides
    """
    global current_data

    if switches is None:
        switches = []

    # Build a set of poleids to remove from switches
    switch_removes = { s['remove_poleid'] for s in switches }

    # Combine: explicitly checked + switch overrides
    remove_ids = set(poleid_list) | switch_removes

    # But remove any that were switched TO keep
    switch_keeps = { s['keep_poleid'] for s in switches }
    remove_ids   = remove_ids - switch_keeps

    if remove_ids:
        clean_poles = [
            p for p in current_data['poles']
            if p.get('poleid', '') not in remove_ids
        ]
        confirmed = [
            p for p in current_data['poles']
            if p.get('poleid', '') in remove_ids
        ]
    else:
        clean_poles = list(current_data['poles'])
        confirmed   = []

    groups = regroup(clean_poles)

    current_data.update({
        "clean_poles":          clean_poles,
        "confirmed_duplicates": confirmed,
        "pending_duplicates":   [],
        **groups,
    })

    return {
        "total_poles":     len(current_data['poles']),
        "poles":           clean_poles,
        "duplicates":      confirmed,
        "spelling_issues": [],
        **groups,
    }


def generate_kml():
    poles = current_data.get('clean_poles', current_data.get('poles', []))
    lines = ['<?xml version="1.0" encoding="UTF-8"?>']
    lines.append('<kml xmlns="http://www.opengis.net/kml/2.2">')
    lines.append('<Document>')
    lines.append('<name>Pole Survey</name>')

    for pole in poles:
        lat      = pole.get('loc_lat', '')
        lon      = pole.get('loc_long', '')
        name     = pole.get('pole_number', '')
        remarks  = pole.get('remarks', '')
        barangay = pole.get('barangay', '')
        owner    = pole.get('pole_owner', '')

        if not lat or not lon:
            continue

        description = (
            f"Pole Number: {name}\n"
            f"Owner: {owner}\n"
            f"Barangay: {barangay}\n"
            f"Remarks: {remarks}"
        )

        lines.append(f'''  <Placemark>
    <name>{name}</name>
    <description>{description}</description>
    <Point>
      <coordinates>{lon},{lat},0</coordinates>
    </Point>
  </Placemark>''')

    lines.append('</Document>')
    lines.append('</kml>')
    return '\n'.join(lines)


def get_download_csv(download_type, value=''):
    poles = current_data.get('clean_poles', current_data.get('poles', []))

    # Sort by item number if present, otherwise keep existing order
    def sort_key(pole):
        remarks = pole.get('remarks', '') or ''
        m = _re.match(r'^#(\d+)', remarks.strip())
        return int(m.group(1)) if m else float('inf')

    poles = sorted(poles, key=sort_key)

    if download_type == 'all':
        rows = poles
    elif download_type == 'duplicates':
        rows = current_data.get('confirmed_duplicates',
               current_data.get('pending_duplicates', []))
    elif download_type == 'barangay':
        rows = [p for p in poles if p.get('barangay', '').strip() == value]
    elif download_type == 'municipality':
        rows = [p for p in poles if p.get('municipality', '').strip() == value]
    else:
        rows = poles

    if not rows:
        return ''

    # Remove internal tracking fields
    clean_rows = []
    for row in rows:
        clean_row = {k: v for k, v in row.items() if not k.startswith('_')}
        clean_rows.append(clean_row)

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=clean_rows[0].keys())
    writer.writeheader()
    writer.writerows(clean_rows)
    return output.getvalue()

def get_municipality_zip(municipality_name):
    import zipfile
    import io as _io

    poles = current_data.get('clean_poles', current_data.get('poles', []))

    # Filter poles belonging to this municipality
    muni_poles = [
        p for p in poles
        if p.get('municipality', '').strip() == municipality_name
    ]

    # Group by barangay
    barangay_groups = {}
    for pole in muni_poles:
        brgy = pole.get('barangay', '').strip() or 'Unknown'
        if brgy not in barangay_groups:
            barangay_groups[brgy] = []
        barangay_groups[brgy].append(pole)

    # Build ZIP in memory
    zip_buffer = _io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for brgy, brgy_poles in sorted(barangay_groups.items()):
            clean_rows = [
                {k: v for k, v in p.items() if not k.startswith('_')}
                for p in brgy_poles
            ]
            csv_buffer = _io.StringIO()
            writer = csv.DictWriter(csv_buffer, fieldnames=clean_rows[0].keys())
            writer.writeheader()
            writer.writerows(clean_rows)
            filename = f'{brgy}.csv'.replace(' ', '_')
            zf.writestr(filename, csv_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer.read()

import re as _re

def extract_item_number(remarks):
    """Extract leading #N from remarks. Returns (number, rest) or (None, remarks)."""
    if not remarks:
        return None, remarks
    m = _re.match(r'^#(\d+)\s*(.*)', remarks.strip())
    if m:
        return int(m.group(1)), m.group(2).strip()
    return None, remarks.strip()


def build_sequence(poles):
    """
    Build sequence list from poles.
    Each entry: { poleid, pole_number, remarks, item_number, rest_of_remarks, barangay, municipality, ... all fields }
    Poles with item numbers come first sorted by item number.
    Poles without come after sorted by timestamp.
    """
    with_num    = []
    without_num = []

    for pole in poles:
        remarks = pole.get('remarks', '') or ''
        num, rest = extract_item_number(remarks)
        entry = dict(pole)
        entry['_item_number']     = num
        entry['_rest_of_remarks'] = rest
        if num is not None:
            with_num.append(entry)
        else:
            without_num.append(entry)

    with_num.sort(key=lambda p: p['_item_number'])
    return with_num + without_num


def apply_sequence(sequence_data):
    """
    Write back corrected item numbers + remarks into current_data poles.
    sequence_data: list of { poleid, item_number, remarks }
    """
    global current_data

    lookup = { s['poleid']: s for s in sequence_data }

    for pole in current_data['poles']:
        pid = pole.get('poleid', '')
        if pid in lookup:
            pole['remarks'] = lookup[pid]['remarks']

    clean_poles = current_data.get('clean_poles', current_data['poles'])
    for pole in clean_poles:
        pid = pole.get('poleid', '')
        if pid in lookup:
            pole['remarks'] = lookup[pid]['remarks']

    current_data['clean_poles'] = clean_poles
    current_data['poles']       = current_data['poles']

    groups = regroup(clean_poles)
    current_data.update(groups)

    return {
        "status": "ok",
        "poles":  clean_poles,
        **groups,
    }


def get_sequence():
    """Return the current sequence for the UI."""
    poles = current_data.get('clean_poles', current_data.get('poles', []))
    return build_sequence(poles)

def reset_data():
    """Reset all current session data."""
    global current_data
    current_data = {
        "poles": [],
        "all_poles": [],
        "clean_poles": [],
        "pending_duplicates": [],
        "confirmed_duplicates": [],
        "spelling_issues": [],
        "barangays": [],
        "municipalities": [],
        "poles_by_barangay": {},
        "poles_by_municipality": {},
        "total_poles": 0,
    }