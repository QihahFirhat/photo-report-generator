# Dilapidation Survey Photo Report Generator

A Python automation tool that generates a **Dilapidation Survey Photo Report** in Microsoft Excel (`.xlsx`) from photographs stored in a ZIP or RAR archive.

The script extracts images, arranges them into formatted worksheets, and adds editable photo labels to simplify survey documentation.

## Features

- Extracts images from `.zip` and `.rar` archives, including nested folders.
- Supports JPG, JPEG, PNG, BMP, GIF, TIFF, and WEBP images.
- Arranges **6 photos per worksheet** in a **3-column × 2-row** layout.
- Embeds photos at a fixed display size of **220 × 293 pixels**.
- Automatically labels photos (`Photo 1`, `Photo 2`, etc.).
- Creates additional worksheets as needed (`Page 1`, `Page 2`, etc.).
- Adds borders, centered descriptions, and report page headers.
- Configures **A4 landscape** print settings and page margins.
- Produces an editable Excel workbook for inspection remarks.

## Technologies and Techniques

| Technology | Purpose |
|---|---|
| Python | Core automation |
| `openpyxl` | Excel workbook generation, styling, and image embedding |
| `Pillow` | Image-format support for Excel embedding |
| `zipfile` | ZIP archive extraction |
| `rarfile` | Optional RAR archive extraction |
| `argparse` | Command-line inputs |
| `pathlib` and `tempfile` | File handling and temporary storage |

**Techniques:** file automation, archive extraction, image embedding and fixed-size scaling, grid-based document layout, and Excel report generation. **No machine learning is used.**

## Project Structure

```text
Dilapidation-Survey/
├── generate_report.py
├── README.md
├── requirements.txt
├── photos.zip                  # Example input (not included in repo)
└── dilapidation_report.xlsx    # Generated output (not included in repo)
```

> Replace `generate_report.py` in the examples with your actual Python script filename.

## Installation

Install Python 3.10+ (recommended), then run:

```bash
pip install openpyxl Pillow rarfile
```

Or create `requirements.txt`:

```text
openpyxl
Pillow
rarfile
```

And install it with:

```bash
pip install -r requirements.txt
```

`rarfile` is optional if you only process ZIP files. For RAR files, the current code expects **UnRAR.exe** at:

```text
C:\Program Files\WinRAR\UnRAR.exe
```

Install WinRAR if necessary, and update `rarfile.UNRAR_TOOL` in the script if its location differs.

## Usage

### 1. Prepare the images

Create a ZIP or RAR archive containing your photographs:

```text
photos.zip
├── IMG_001.jpg
├── IMG_002.jpg
├── IMG_003.jpg
├── IMG_004.jpg
├── IMG_005.jpg
├── IMG_006.jpg
└── IMG_007.jpg
```

Images are sorted **alphabetically by filename**, so sequential filenames help maintain the intended order.

### 2. Generate the report

Use the default output name:

```bash
python generate_report.py photos.zip
```

Or choose an output name:

```bash
python generate_report.py photos.zip survey_report.xlsx
```

For RAR archives:

```bash
python generate_report.py photos.rar survey_report.xlsx
```

### 3. Edit the report

Open the generated `.xlsx` file and replace the default `Photo N` labels with details such as defect type, location, structural condition, and inspection remarks.

## Excel Output

Each worksheet includes a report title and up to six photos:

| Column A | Column B | Column C |
|---|---|---|
| Photo 1 image | Photo 2 image | Photo 3 image |
| Photo 1 description | Photo 2 description | Photo 3 description |
| Photo 4 image | Photo 5 image | Photo 6 image |
| Photo 4 description | Photo 5 description | Photo 6 description |

| Setting | Value |
|---|---|
| Paper | A4 landscape |
| Photos per worksheet | 6 |
| Layout | 3 columns × 2 rows |
| Image display size | 220 × 293 px |
| Description row height | 40 px |
| Worksheet names | `Page 1`, `Page 2`, ... |
| Output | `.xlsx` |

For example, **20 photos produce 4 worksheets**.

## Customization

Edit the configuration constants in the script:

```python
COLS_PER_PAGE = 3
ROWS_PER_PAGE = 2
IMG_W_PX = 220
IMG_H_PX = 293
DESC_H_PX = 40
COL_PADDING_PX = 6
```

You can also modify the report title in `add_title_row()` and the default photo label in `place_image_and_desc()`.

> Changing the grid dimensions may require further adjustments to the Excel row layout and page setup, since parts of the current script assume a 3 × 2 grid.

## Processing Workflow

1. **Extract:** Unpack the archive into a temporary directory and find supported images.
2. **Organize:** Sort images by filename and divide them into groups of six.
3. **Format:** Create Excel sheets, place images, add descriptions, borders, and page headings.
4. **Export:** Save the `.xlsx` workbook and automatically delete temporary files.

## Limitations

- Only ZIP and RAR archives are accepted as input.
- The program does **not** analyze photos or automatically detect defects.
- Images are assigned fixed dimensions, which can distort photos with other aspect ratios.
- Ordering is based on filename, not capture time or folder hierarchy.
- Print settings target A4 landscape, but users should verify print preview before exporting or printing; one worksheet is not guaranteed to equal one physical printed page.
- The output directory must already exist.

## Purpose

The tool reduces repetitive work when preparing photographic documentation for building condition assessments and dilapidation surveys. Its input images, filename, report title, and layout can be adapted to different inspection projects.
