import csv
import math
import os
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


def update_progress(progress, percent, message):
    """Update progress tracker."""
    if progress is not None:
        progress["percent"] = percent
        progress["message"] = message
    print(f"  [{percent}%] {message}")


def download_image(url, save_path):
    """Download image from URL, fix orientation and color, save as JPEG."""
    try:
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
        print(f"  Warning: Could not download {url}: {e}")
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
    total_pages = math.ceil(total_poles / 4)
    print(f"  Found {total_poles} poles → {total_pages} sheet(s).")

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

    abs_template = os.path.abspath(template_path)
    abs_output   = os.path.abspath(output_path)
    abs_img_dir  = os.path.abspath(img_dir)

    try:
        wb          = excel.Workbooks.Open(abs_template)
        template_ws = wb.Sheets(1)

        # ── Create pages ──────────────────────────────────────────────────────
        for page_index in range(total_pages):
            page_number        = page_index + 1
            poles_on_this_page = poles[page_index * 4 : page_index * 4 + 4]

            fill_percent = 50 + int((page_index / total_pages) * 45)
            update_progress(progress, fill_percent, f"Filling Page {page_number} of {total_pages}...")

            if page_index == 0:
                ws      = template_ws
                ws.Name = f"{survey_name.upper()} {page_number}"
            else:
                template_ws.Copy(After=wb.Sheets(wb.Sheets.Count))
                ws      = wb.Sheets(wb.Sheets.Count)
                ws.Name = f"{survey_name.upper()} {page_number}"

                # Delete all images copied from template
                for shp in list(ws.Shapes):
                    if shp.Type == 13 or shp.Type == 11:
                        shp.Delete()

            # Restore row heights from template
            for row_num in range(1, 80):
                ws.Rows(row_num).RowHeight = template_ws.Rows(row_num).RowHeight

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
            ws.Range("C1").Value = project_name
            ws.Range("C2").Value = site_name
            ws.Range("C3").Value = address
            ws.Range("C4").Value = date_taken

            # Fill pole blocks
            for block_index, pole in enumerate(poles_on_this_page):
                block = BLOCKS[block_index]
                print(f"  Block {block_index + 1}: Pole {pole.get('pole_number', '')}")

                # Text data
                ws.Range(block['local_loop_cell']).Value = 'MERALCO'

                tag_text = (
                    f"{pole.get('pole_number', '').upper()}\n"
                    f"{pole.get('loc_lat', '')}, {pole.get('loc_long', '')}"
                )
                ws.Range(block['tag_cell']).Value  = tag_text
                ws.Range(block['item_cell']).Value = ''

                # Place images
                pole_index = page_index * 4 + block_index

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
                    ws.Shapes.AddPicture(
                        Filename        = img_path,
                        LinkToFile      = False,
                        SaveWithDocument= True,
                        Left            = 0,
                        Top             = 0,
                        Width           = 100,
                        Height          = 100
                    )

                    ws.Range(block[cell_key]).Select()

                    import time
                    time.sleep(0.3)

                    excel.Run("FitImageToSelectedCell")

                    time.sleep(0.3)

        # ── Save ──────────────────────────────────────────────────────────────
        update_progress(progress, 96, "Saving file...")
        wb.SaveAs(abs_output, FileFormat=52)
        wb.Close(SaveChanges=False)
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