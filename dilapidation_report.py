import sys
import zipfile
import tempfile
import argparse
from pathlib import Path

from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# ── Unit conversion helpers ─────────────────────────────────────────────────────
# openpyxl column width  : Excel "character units"  (1 unit ≈ 7 px at default font)
# openpyxl row height    : points                   (1 pt = 96/72 px ≈ 1.333 px)
# openpyxl image size    : pixels (EMU / 9144)  — we set width/height in px directly

def px_to_col_width(px: float) -> float:
    """Convert pixels to Excel column-width character units."""
    return px / 7.0

def px_to_row_height(px: float) -> float:
    """Convert pixels to Excel row height in points."""
    return px * 72.0 / 96.0

# ── Layout constants ────────────────────────────────────────────────────────────
COLS_PER_PAGE   = 3
ROWS_PER_PAGE   = 2
IMAGES_PER_PAGE = COLS_PER_PAGE * ROWS_PER_PAGE  # 6

# Target image display size in pixels (3:4 portrait ratio).
# These drive EVERYTHING — column width, row height, and the embedded image size
# all derive from these two numbers so they always stay in sync.
IMG_W_PX = 220   # display width  in px  (3 units)
IMG_H_PX = 293   # display height in px  (4 units → 220 * 4/3 ≈ 293)

# Description row below each image
DESC_H_PX = 40   # pixels

# Small padding added to column width so image doesn't touch the border
COL_PADDING_PX = 6

# ── Derived cell dimensions ─────────────────────────────────────────────────────
COL_WIDTH_UNITS  = px_to_col_width(IMG_W_PX + COL_PADDING_PX)
IMG_ROW_HEIGHT   = px_to_row_height(IMG_H_PX)          # points
DESC_ROW_HEIGHT  = px_to_row_height(DESC_H_PX)         # points

# ── Page margins (inches) ───────────────────────────────────────────────────────
MARGIN_LEFT_IN   = 2.0 / 2.54
MARGIN_RIGHT_IN  = 1.0 / 2.54
MARGIN_TOP_IN    = 1.0 / 2.54
MARGIN_BOTTOM_IN = 1.0 / 2.54

# ── Styles ──────────────────────────────────────────────────────────────────────
THIN        = Side(style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DESC_FONT   = Font(name="Calibri", size=11)
DESC_ALIGN  = Alignment(horizontal="center", vertical="center", wrap_text=True)
HDR_FONT    = Font(name="Calibri", size=13, bold=True)

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp"}


# ── Archive extraction ──────────────────────────────────────────────────────────

def extract_archive(archive_path: Path, dest_dir: Path) -> list:
    suffix = archive_path.suffix.lower()

    if suffix == ".zip":
        with zipfile.ZipFile(archive_path, "r") as zf:
            zf.extractall(dest_dir)

    elif suffix == ".rar":
        try:
            import rarfile
            rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"
        except ImportError:
            sys.exit(
                "ERROR: 'rarfile' package not found.\n"
                "Install it with:  pip install rarfile\n"
                "Also install WinRAR: winget install rarlab.winrar"
            )
        with rarfile.RarFile(archive_path, "r") as rf:
            rf.extractall(dest_dir)
    else:
        sys.exit(f"ERROR: Unsupported archive format '{suffix}'. Use .zip or .rar.")

    images = sorted(
        [p for p in dest_dir.rglob("*") if p.suffix.lower() in SUPPORTED_EXTS],
        key=lambda p: p.name
    )
    if not images:
        sys.exit("ERROR: No supported image files found inside the archive.")

    print(f"  Found {len(images)} image(s).")
    return images


# ── Excel helpers ───────────────────────────────────────────────────────────────

def configure_sheet(ws):
    """A4 landscape page setup and margins."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0

    ws.page_margins = PageMargins(
        left=MARGIN_LEFT_IN, right=MARGIN_RIGHT_IN,
        top=MARGIN_TOP_IN,   bottom=MARGIN_BOTTOM_IN,
        header=0, footer=0,
    )

    # All three image columns get the same width
    for col_idx in range(1, COLS_PER_PAGE + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTH_UNITS


def add_title_row(ws, sheet_num: int, total_sheets: int):
    ws.row_dimensions[1].height = 20
    cell = ws.cell(row=1, column=1,
                   value=f"Dilapidation Survey — Photo Report   (Page {sheet_num} of {total_sheets})")
    cell.font      = HDR_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=COLS_PER_PAGE)


def apply_row_heights(ws):
    """
    Row layout (Excel rows):
      1        → title
      2, 4     → image rows
      3, 5     → description rows
    """
    for grid_row in range(ROWS_PER_PAGE):
        img_row  = 2 + grid_row * 2
        desc_row = img_row + 1
        ws.row_dimensions[img_row].height  = IMG_ROW_HEIGHT
        ws.row_dimensions[desc_row].height = DESC_ROW_HEIGHT


def place_image_and_desc(ws, img_path: Path, grid_row: int, grid_col: int, img_index: int):
    """Embed one image and its description cell at the given grid slot."""
    excel_row_img  = 2 + grid_row * 2
    excel_row_desc = excel_row_img + 1
    excel_col      = grid_col + 1

    # Border on the image cell
    img_cell = ws.cell(row=excel_row_img, column=excel_col)
    img_cell.border    = THIN_BORDER
    img_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Embed the image — width & height in PIXELS to match the cell exactly
    xl_img        = XLImage(str(img_path))
    xl_img.width  = IMG_W_PX
    xl_img.height = IMG_H_PX
    xl_img.anchor = f"{get_column_letter(excel_col)}{excel_row_img}"
    ws.add_image(xl_img)

    # Description cell
    desc_cell           = ws.cell(row=excel_row_desc, column=excel_col,
                                  value=f"Photo {img_index + 1}")
    desc_cell.font      = DESC_FONT
    desc_cell.alignment = DESC_ALIGN
    desc_cell.border    = THIN_BORDER


# ── Main ────────────────────────────────────────────────────────────────────────

def build_report(archive_path: Path, output_path: Path):
    print(f"\n[1/4] Extracting archive: {archive_path.name}")
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        images  = extract_archive(archive_path, tmp_dir)
        total   = len(images)
        n_pages = (total + IMAGES_PER_PAGE - 1) // IMAGES_PER_PAGE

        print(f"[2/4] Building workbook ({n_pages} page(s), {total} image(s)) ...")
        wb = Workbook()
        wb.remove(wb.active)

        for page_idx in range(n_pages):
            sheet_num = page_idx + 1
            ws = wb.create_sheet(title=f"Page {sheet_num}")
            configure_sheet(ws)
            apply_row_heights(ws)
            add_title_row(ws, sheet_num, n_pages)

            start = page_idx * IMAGES_PER_PAGE
            end   = min(start + IMAGES_PER_PAGE, total)

            for slot, img_idx in enumerate(range(start, end)):
                grid_row = slot // COLS_PER_PAGE
                grid_col = slot  % COLS_PER_PAGE
                place_image_and_desc(ws, images[img_idx], grid_row, grid_col, img_idx)

            print(f"  Page {sheet_num}/{n_pages}: images {start+1}-{end}")

        print(f"[3/4] Saving to: {output_path}")
        wb.save(output_path)

    print(f"[4/4] Done! Report saved -> {output_path}\n")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate a Dilapidation Survey photo report in Excel."
    )
    parser.add_argument("archive", help="Path to .zip or .rar file.")
    parser.add_argument("output",  nargs="?", default="dilapidation_report.xlsx",
                        help="Output .xlsx filename.")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    build_report(Path(args.archive), Path(args.output))
