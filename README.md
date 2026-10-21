# TxDOT Aviation Bid Summaries

This script takes a TxDOT aviation project PDF and turns the bids into an Excel comparison, a PDF summary, and a master sheet of prices from the lowest bidder.

## Getting started

You'll need **Python 3.10 or newer** and an internet connection. Open a terminal in the project folder and run:

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install -r requirements.txt
```

On Windows, use `venv\Scripts\activate` to activate the environment.

The packages are listed in `requirements.txt`. The first run may take longer because Docling needs to download its OCR models. After that, it reads the PDFs locally on your computer.

## Required packages

| Package | What it does |
| --- | --- |
| `beautifulsoup4` | Reads the project names and PDF links from the TxDOT website. |
| `requests` | Downloads the website pages and project PDFs. |
| `pypdf` | Reads the PDF's editable summary fields. |
| `docling` | Uses local OCR to extract text and tables from scanned pages. |
| `openpyxl` | Creates and updates the Excel workbooks. |
| `reportlab` | Creates the PDF comparison report. |
| `rich` | Displays the progress bar and console messages. |

Install these together using `python -m pip install -r requirements.txt`. The version requirements are listed in that file.

## Running the script

To choose a project from the TxDOT website:

```bash
python main.py
```

Enter the number next to the project you want. The script downloads the PDF and creates the reports. A progress bar shows which step it's working on.

You can also give it a PDF link directly. For example, to use the Lockhart project:

```bash
python main.py "https://www.dot.state.tx.us/AVN/avninfo/status/construct/ntb2614lhart00034br.pdf" outputs
```

`outputs` is the folder where the files are saved. You can change it to another folder name. When the script finishes, it prints the full paths to the three reports.

## Output files

| File | What's in it |
| --- | --- |
| `<CSJ>_summary.xlsx` | Each contractor's unit prices and totals shown side by side, with the apparent lowest bidder marked. A second tab keeps the extracted records and their PDF page numbers. |
| `<CSJ>_summary.pdf` | A printable version of the bid comparison. |
| `master_bid_prices.xlsx` | The lowest bidder's item prices, along with the project details. If bidders tie for the lowest total, both are included. |
| `<CSJ>_source.pdf` | A copy of the original PDF. |

If you run the same project again, its summary files are replaced. The master sheet updates that project's rows without adding duplicates and keeps the other projects. Use the same output folder if you want to keep adding to one master sheet. Keep any custom analysis in a separate workbook, since the script rebuilds the master in its own format.

## Project files

- `main.py` handles project selection, downloading, and the progress bar.
- `pdf_extractor.py` reads the PDF and matches the bid tables to contractors.
- `excel_export.py` creates the Excel files.
- `pdf_export.py` creates the PDF summary.

## Limitations

Some pages are scanned images, so the script uses OCR to read them. **The extraction won't always be perfect.** Blurry scans or unusual formatting can cause it to misread numbers, item codes, or descriptions.

The script checks the price calculations and stops if it finds a mismatch or can't tell which contractor a table belongs to. These checks can't catch every OCR mistake, so the original PDF is saved for reference.

This version is built around the TxDOT aviation format shown in the assignment. Different layouts, unclear scans, or alternate bid items may need additional work. The lowest bid shown in the reports is an apparent low bid, not confirmation of an award.
