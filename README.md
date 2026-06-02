# dilapidation-report-Digimap.Sdn.Bhd

## Dilapidation Survey Photo Report Generator

Generates a professional Excel photo report from images in a .zip or .rar archive.

Layout: A4 Landscape, 2 rows x 3 columns per sheet (6 images per page).
Each image is followed by a formatted description cell.

Dependencies:
    pip install openpyxl pillow rarfile
    # For .rar support, also install WinRAR (Windows) or unrar (Linux/macOS)
    # Windows:       winget install rarlab.winrar
    # Ubuntu/Debian: sudo apt-get install unrar
    # macOS:         brew install unrar

Usage:
    python dilapidation_report.py <archive_path> [output_path]

Example:
    python dilapidation_report.py survey_photos.rar report.xlsx
