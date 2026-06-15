import csv
import math
import os
import re
import requests
import win32com.client
from io import BytesIO
from PIL import Image as PILImage, ImageOps


# ── Block layout ──────────────────────────────────────────────────────────────
BLOCKS = [
    {
        "local_loop_cell": "K5",
        "tag_cell":        "K6",
        "item_cell":       "K7",
        "left_cell":       "A5",
        "center_cell":     "E5",
        "right_cell":      "I8",
    },
    {
        "local_loop_cell": "K24",
        "tag_cell":        "K25",
        "item_cell":       "K26",
        "left_cell":       "A24",
        "center_cell":     "E24",
        "right_cell":      "I27",
    },
    {
        "local_loop_cell": "K43",
        "tag_cell":        "K44",
        "item_cell":       "K45",
        "left_cell":       "A43",
        "center_cell":     "E43",
        "right_cell":      "I46",
    },
    {
        "local_loop_cell": "K62",
        "tag_cell":        "K63",
        "item_cell":       "K64",
        "left_cell":       "A62",
        "center_cell":     "E62",
        "right_cell":      "I65",
    },
]

def format_date_range(dates):
    from datetime import datetime
    parsed = []
    for d in dates:
        d = re.sub(r'#', '', d).strip()
        for fmt in ['%d/%m/%Y %H:%M', '%Y-%m-%d %H:%M:%S',
                    '%m/%d/%Y %H:%M', '%d/%m/%Y %H:%M:%S']:
            try:
                parsed.append(datetime.strptime(d.strip(), fmt))
                break
            except:
                continue
    if not parsed:
        return ''
    parsed.sort()
    months = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE',
              'JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']
    earliest = parsed[0]
    latest   = parsed[-1]
    month    = months[earliest.month - 1]
    year     = earliest.year
    if earliest.day == latest.day and earliest.month == latest.month:
        return f"{month} {earliest.day}, {year}"
    elif earliest.month == latest.month:
        return f"{month} {earliest.day}-{latest.day}, {year}"
    else:
        month2 = months[latest.month - 1]
        return f"{month} {earliest.day} - {month2} {latest.day}, {year}"

def update_progress(progress, percent, message):
    """Update progress tracker."""
    if progress is not None:
        progress["percent"] = percent
        progress["message"] = message
    print(f"  [{percent}%] {message}")


def download_image(url, save_path):
    """Download image from URL with retry, fix orientation and color, save as JPEG."""
    max_retries = 3
    
    for attempt in range(max_retries):
        try:
            print(f"    Downloading (attempt {attempt + 1})...")
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            img = PILImage.open(BytesIO(response.content))

            # Fix EXIF rotation
            img = ImageOps.exif_transpose(img)

            # Fix red tint — convert RGBA/P to RGB properly
            if img.mode in ('RGBA', 'P', 'LA'):
                background = PILImage.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')

            # Resize to max 800px to reduce file size
            max_size = 800
            if img.width > max_size or img.height > max_size:
                img.thumbnail((max_size, max_size), PILImage.LANCZOS)

            # Save as JPEG
            jpeg_path = save_path.replace('.png', '.jpg')
            img.save(jpeg_path, format='JPEG', quality=75, optimize=True)
            return True

        except Exception as e:
            print(f"  Warning attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                import time
                time.sleep(2)  # wait 2 seconds before retrying
            else:
                print(f"  Failed after {max_retries} attempts - skipping")
                return False


def generate_excel(template_path, csv_path, output_path,
                   project_name, site_name, address, date_taken,
                   survey_name, progress=None):

    # ── Read CSV ──────────────────────────────────────────────────────────────
    update_progress(progress, 5, "Reading CSV...")
    poles = []
    with open(csv_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            poles.append(row)

    total_poles = len(poles)

    # Pack poles continuously 4 per page
    pages       = []
    page_labels = []
    for pole in poles:
        if not pages or len(pages[-1]) == 4:
            pages.append([])
            page_labels.append(pole.get('barangay', '').strip().upper() or 'UNKNOWN')
        pages[-1].append(pole)

    total_pages = len(pages)
    print(f"  Found {total_poles} poles → {total_pages} page(s).")

    # ── Download all images first ─────────────────────────────────────────────
    update_progress(progress, 10, "Downloading images...")
    img_dir = os.path.join(os.path.dirname(output_path), 'temp_images')
    os.makedirs(img_dir, exist_ok=True)

    for i, pole in enumerate(poles):
        for photo_type in ['photoWhole', 'photoAttachments', 'photoTags']:
            url       = pole.get(photo_type, '')
            save_path = os.path.join(img_dir, f'pole_{i}_{photo_type}.jpg')

            if url and url.startswith('http') and not os.path.exists(save_path):
                total_images  = total_poles * 3
                current_image = i * 3 + ['photoWhole', 'photoAttachments', 'photoTags'].index(photo_type) + 1
                dl_percent    = 10 + int((current_image / total_images) * 40)
                update_progress(progress, dl_percent, f"Downloading image {current_image} of {total_images}...")
                success = download_image(url, save_path)
                if not success:
                    print(f"    Failed - will skip this image")

    # ── Open Excel via win32com ───────────────────────────────────────────────
    update_progress(progress, 50, "Opening Excel...")
    import pythoncom
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False

    # When running as exe, extract template to a temp location
    import sys
    if getattr(sys, 'frozen', False):
        exe_dir      = os.path.dirname(sys.executable)
        abs_template = os.path.join(exe_dir, 'Template.xlsm')
    else:
        abs_template = os.path.abspath(template_path)
    
    print(f"  Template path: {abs_template}")
    print(f"  Template exists: {os.path.exists(abs_template)}")
    abs_output   = os.path.abspath(output_path)
    abs_img_dir  = os.path.abspath(img_dir)

    try:
        template_wb = excel.Workbooks.Open(abs_template)
        wb_name     = template_wb.Name
        # Store references to all page workbooks
        page_workbooks = []

        # ── Create pages ──────────────────────────────────────────────────────
        global_pole_index = 0
        for page_index, poles_on_this_page in enumerate(pages):
            page_number = page_index + 1

            fill_percent = 50 + int((page_index / total_pages) * 45)
            update_progress(progress, fill_percent, f"Filling Page {page_number} of {total_pages}...")

            wb = excel.Workbooks(wb_name)
            # Each page gets its own copy of template
            if page_index == 0:
                page_wb = template_wb
            else:
                template_wb.Sheets(1).Copy()
                page_wb = excel.ActiveWorkbook
                ws = page_wb.Sheets(1)
                ws.Activate()
                for shp in list(ws.Shapes):
                    if shp.Type == 13 or shp.Type == 11:
                        shp.Delete()
            page_workbooks.append(page_wb)
            wb = page_wb
            ws = page_wb.Sheets(1)

            # Sheet tab name — continuous numbering across all pages
            if survey_name:
                desired_name = f"{survey_name.upper()} {page_number}"
            else:
                desired_name = f"{page_labels[page_index]} {page_number}"

            desired_name   = re.sub(r'[\\/*?[\]:]', '', desired_name)[:28]
            existing_names = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
            final_name     = desired_name
            counter        = 2
            while final_name in existing_names:
                final_name = f"{desired_name[:25]}_{counter}"
                counter   += 1
            ws.Name = final_name


            # Restore row heights from template
            for row_num in range(1, 80):
                try:
                    ws.Rows(row_num).RowHeight = wb.Sheets(1).Rows(row_num).RowHeight
                except:
                    pass

            # Page setup
            ws.PageSetup.Orientation     = 1
            ws.PageSetup.PaperSize       = 9
            ws.PageSetup.FitToPagesWide  = 1
            ws.PageSetup.FitToPagesTall  = 1
            ws.PageSetup.Zoom            = False
            ws.PageSetup.LeftMargin      = excel.InchesToPoints(0.3)
            ws.PageSetup.RightMargin     = excel.InchesToPoints(0.3)
            ws.PageSetup.TopMargin       = excel.InchesToPoints(0.3)
            ws.PageSetup.BottomMargin    = excel.InchesToPoints(0.3)
            ws.PageSetup.HeaderMargin    = excel.InchesToPoints(0)
            ws.PageSetup.FooterMargin    = excel.InchesToPoints(0)
            ws.PageSetup.PrintArea       = "A1:L79"

            # Fill headers
            first_barangay = poles_on_this_page[0].get('barangay', '').strip().upper()

            page_address = address if address else f"ALONG {first_barangay}" if first_barangay else ''

            page_dates = [
                re.sub(r'#', '', p.get('created_at', '')).strip()
                for p in poles_on_this_page
                if p.get('created_at', '').strip()
            ]
            page_date = date_taken if date_taken else format_date_range(page_dates)

            ws.Range("C1").Value = project_name
            ws.Range("C2").Value = site_name
            ws.Range("C3").Value = page_address
            ws.Range("C4").Value = page_date

            # Fill pole blocks
            for block_index, pole in enumerate(poles_on_this_page):
                block = BLOCKS[block_index]
                print(f"  Block {block_index + 1}: Pole {pole.get('pole_number', '')}")

                # Text data — use pole_owner from CSV
                ws.Range(block['local_loop_cell']).Value = pole.get('pole_owner', 'MERALCO').upper()

                tag_text = (
                    f"{pole.get('pole_number', '').upper()}\n"
                    f"{pole.get('loc_lat', '')}, {pole.get('loc_long', '')}"
                )
                ws.Range(block['tag_cell']).Value  = tag_text
                remarks    = pole.get('remarks', '') or ''
                item_match = re.match(r'^#(\d+)', remarks.strip())
                ws.Range(block['item_cell']).Value = int(item_match.group(1)) if item_match else ''

                # Place images
                pole_index = global_pole_index

                for photo_type, cell_key in [
                    ('photoWhole',       'left_cell'),
                    ('photoAttachments', 'center_cell'),
                    ('photoTags',        'right_cell'),
                ]:
                    img_path = os.path.join(abs_img_dir, f'pole_{pole_index}_{photo_type}.jpg')

                    url = pole.get(photo_type, '')
                    if not url or not url.startswith('http'):
                        print(f"    Skipping {photo_type} - no URL")
                        continue

                    if not os.path.exists(img_path):
                        print(f"    Skipping {photo_type} - image not found")
                        continue

                    if os.path.getsize(img_path) == 0:
                        print(f"    Skipping {photo_type} - empty file")
                        continue

                    try:
                        test = PILImage.open(img_path)
                        test.verify()
                    except Exception:
                        print(f"    Skipping {photo_type} - corrupt image")
                        continue

                    print(f"    Placing {photo_type} into {block[cell_key]}...")

                    ws.Activate()
                    excel.ActiveWindow.ScrollRow = 1
                    ws.Shapes.AddPicture(
                        Filename        = img_path,
                        LinkToFile      = False,
                        SaveWithDocument= True,
                        Left            = 0,
                        Top             = 0,
                        Width           = 100,
                        Height          = 100
                    )

                    ws.Activate()
                    ws.Range(block[cell_key]).Select()

                    import time
                    time.sleep(0.3)

                    excel.Run(f"'{wb_name}'!FitImageToSelectedCell")

                    time.sleep(0.3)
                global_pole_index += 1

        # ── Save ──────────────────────────────────────────────────────────────
        update_progress(progress, 96, "Saving file...")
        # Save each page workbook to temp file first
        update_progress(progress, 93, "Combining sheets...")
        temp_paths = []
        for i, page_wb in enumerate(page_workbooks):
            temp_path = os.path.join(os.path.dirname(abs_output), f'_temp_page_{i}.xlsm')
            page_wb.SaveAs(temp_path, FileFormat=52)
            page_wb.Close(SaveChanges=False)
            temp_paths.append(temp_path)

        # Re-open template so macro is available
        macro_wb = excel.Workbooks.Open(abs_template)
        macro_wb_name = macro_wb.Name

        # Debug: show all temp files
        print(f"  Temp files to combine: {len(temp_paths)}")
        for tp in temp_paths:
            print(f"    - {tp} exists: {os.path.exists(tp)}")

        # Open all temp workbooks then use VBA macro to copy sheets
        opened_wbs   = []
        for tp in temp_paths:
            opened_wbs.append(excel.Workbooks.Open(tp))

        base_wb      = opened_wbs[0]
        base_wb_name = base_wb.Name
        print(f"  Opened base_wb: {base_wb_name}, sheets: {base_wb.Sheets.Count}")

        for src_wb in opened_wbs[1:]:
            sheet_name = src_wb.Sheets(1).Name
            print(f"  Copying sheet {sheet_name} from {src_wb.Name} to {base_wb_name}...")
            excel.Run(
                f"'{macro_wb_name}'!CopySheetToWorkbook",
                src_wb.Name,
                sheet_name,
                base_wb_name
            )
            src_wb.Close(SaveChanges=False)
            base_wb = excel.Workbooks(base_wb_name)
            print(f"  base_wb sheets now: {base_wb.Sheets.Count}")

        # Re-open base_wb in case it got disconnected
        try:
            sheet_count = base_wb.Sheets.Count
        except:
            base_wb = excel.Workbooks.Open(temp_paths[0])

        print(f"  Total sheets before save: {base_wb.Sheets.Count}")
        update_progress(progress, 96, "Saving file...")
        base_wb.SaveAs(abs_output, FileFormat=52)
        base_wb.Close(SaveChanges=False)
        try:
            macro_wb.Close(SaveChanges=False)
        except:
            pass

        # Cleanup temp files
        for temp_path in temp_paths:
            try:
                os.remove(temp_path)
            except:
                pass

        update_progress(progress, 100, "Done! ✅")
        print("Done! ✅")

    except Exception as e:
        print(f"Error: {e}")
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        raise e

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    # ── Cleanup temp images ───────────────────────────────────────────────────
    print("Cleaning up temp images...")
    for f in os.listdir(img_dir):
        os.remove(os.path.join(img_dir, f))
    os.rmdir(img_dir)
