import time
import re
from geopy.geocoders import Nominatim


geolocator = Nominatim(user_agent="pole-survey-tool", timeout=10)


def clean_remarks(remarks):
    """Remove leading #number from remarks."""
    cleaned = re.sub(r'^#\d+\s*', '', remarks.strip())
    return cleaned.strip()


def get_street_name(lat, lon):
    """Get most specific street name from coordinates."""
    try:
        location = geolocator.reverse(
            f"{lat}, {lon}",
            language='en',
            zoom=19,
            addressdetails=True
        )
        if not location:
            return None

        address = location.raw.get('address', {})
        print(f"    Full address data: {address}")

        street = (
            address.get('road') or
            address.get('pedestrian') or
            address.get('footway') or
            address.get('path') or
            address.get('residential') or
            address.get('suburb') or
            None
        )

        return street.upper() if street else None

    except Exception as e:
        print(f"  Warning: Could not get street for {lat},{lon}: {e}")
        return None


def clean_landmark(landmark, street):
    """Deeply clean landmark text."""
    if not landmark:
        return None

    original = landmark

    # Step 1: Remove leading number patterns like: #1, 1), 1., (1), 9), 41).
    landmark = re.sub(r'^[\#\(\s]*\d+[\.\)\s]+', '', landmark).strip()

    # Step 2: Remove ALL road/street references anywhere in text
    # This catches: "PASO DE BLAS RD", "PASO DE BLAS ROAD", "ITC ROAD", etc.
    landmark = re.sub(
        r'\b[\w\s]+(?:road|rd|street|st|ave|avenue|blvd|boulevard|highway|hwy|drive|dr)\b[\,\s]*',
        '', landmark, flags=re.IGNORECASE
    ).strip()

    # Step 3: Remove "ALONG" keyword anywhere
    landmark = re.sub(r'\balong\b[\,\s]*', '', landmark, flags=re.IGNORECASE).strip()

    # Step 4: Remove position words at start
    landmark = re.sub(
        r'^\s*(left|right|front\s*of\.?|back\s*of\.?|beside|near|corner|brgy\.?|brngy\.?)\s*',
        '', landmark, flags=re.IGNORECASE
    ).strip()

    # Step 5: Remove leftover punctuation at start/end
    landmark = re.sub(r'^[\s\,\.\)\(\#\d]+', '', landmark).strip()
    landmark = re.sub(r'[\s\,\.]+$', '', landmark).strip()

    # Step 6: If only short leftover like "RD" or "ST" alone — return None
    if re.match(r'^(rd|st|road|street|ave|blvd|dr|of|the|a|an)$', 
                landmark, flags=re.IGNORECASE):
        return None

    # Step 7: If only numbers/punctuation left — return None
    if re.match(r'^[\d\s\.\,\)\(\#\&N]+$', landmark):
        return None

    # Step 8: If empty return None
    return landmark.strip() if landmark.strip() else None


def format_address(street, position, landmark):
    """Format the final address string."""
    landmark = clean_landmark(landmark, street)

    if street and landmark:
        return f"ALONG {street}, {landmark.upper()}"
    elif street and not landmark:
        return f"ALONG {street}"
    elif landmark and not street:
        return f"{landmark.upper()}"
    else:
        return ""


def resolve_addresses(poles, progress_callback=None):
    """Resolve addresses for all poles."""
    total   = len(poles)
    updated = []

    for i, pole in enumerate(poles):
        if progress_callback:
            percent = int((i / total) * 100)
            progress_callback(percent, f"Resolving address {i+1} of {total}...")

        lat      = pole.get('loc_lat', '').strip()
        lon      = pole.get('loc_long', '').strip()
        remarks  = pole.get('remarks', '').strip()
        position = pole.get('position', '').strip()

        # Validate coordinates
        try:
            lat_f = float(lat)
            lon_f = float(lon)
            if not (3 <= lat_f <= 22 and 114 <= lon_f <= 127):
                print(f"  Skipping pole {pole.get('pole_number')} - invalid coordinates")
                updated.append(pole)
                continue
        except:
            updated.append(pole)
            continue

        print(f"  Processing pole {pole.get('pole_number', '')} ({i+1}/{total})")

        # Step 1: Get street name from coordinates
        street = get_street_name(lat, lon)
        time.sleep(1)  # Rate limit

        # Step 2: Clean remarks for landmark
        landmark = clean_remarks(remarks) if remarks else None
        print(f"    Landmark from remarks: '{landmark}'")

        # Step 3: Format address
        formatted = format_address(street, position, landmark)

        # Preserve leading #N item number if present
        import re as _re
        item_match = _re.match(r'^(#\d+)\s*', remarks)
        prefix = (item_match.group(1) + ' ') if item_match else ''

        new_pole = dict(pole)
        if formatted and formatted.strip():
            new_pole['remarks'] = prefix + formatted
            print(f"    ✅ {prefix + formatted}")
        else:
            new_pole['remarks'] = prefix.strip()
            print(f"    ⚠️ No address resolved — keeping prefix only")

        updated.append(new_pole)

    if progress_callback:
        progress_callback(100, "Address resolution complete!")

    return updated