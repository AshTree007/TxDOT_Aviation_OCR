from __future__ import annotations

from bs4 import BeautifulSoup as bs
import requests as re
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TimeElapsedColumn
from pathlib import Path
from urllib.parse import urljoin
import sys

from pdf_extractor import get_summary, get_bid_details
from excel_export import save_summary, save_master
from pdf_export import save_pdf


def getData(url: str) -> bs:
    response = re.get(url, timeout=30)
    response.raise_for_status()
    return bs(response.content, "html.parser")


def get_name_to_links(page: bs) -> dict[str, str]:
    names_to_link = {}
    for row in page.select("table tr"):
        cells = row.find_all("td")
        if len(cells) != 3:
            continue
        link = cells[0].find("a", href=True)
        if link is None:
            continue
        name = " ".join(cells[0].stripped_strings)
        # Include project ID so two entries at the same airport remain selectable.
        project_id = " ".join(cells[1].stripped_strings)
        names_to_link[f"{name} ({project_id})"] = str(link["href"])
    return names_to_link


def process_pdf(url: str, output_folder: str = "outputs") -> dict[str, str]:
    """Download, OCR, validate and export; return absolute Excel and PDF paths.

    Raises an exception on failure. A failure never returns success paths.
    """
    console = Console(file=sys.stdout)
    output = Path(output_folder).expanduser().resolve()

    with Progress(
        SpinnerColumn(),
        "[progress.description]{task.description}",
        BarColumn(),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Downloading PDF...", total=None)
        response = re.get(url, timeout=(15, 120))
        response.raise_for_status()
        if not response.content[:1024].lstrip().startswith(b"%PDF-"):
            raise ValueError("The address did not return a PDF.")

        progress.update(task, description="Reading project summary...")
        summary = get_summary(response)

        progress.update(task, description="Reading scanned bid tables...")
        contractors = get_bid_details(response, summary)

        output.mkdir(parents=True, exist_ok=True)
        prefix = "".join(character if character.isalnum() else "_" for character in summary["project"]["TxDOT CSJ#"])
        (output / f"{prefix}_source.pdf").write_bytes(response.content)

        progress.update(task, description="Creating comparison Excel...")
        summary_path = save_summary(summary, contractors, url, output / f"{prefix}_summary.xlsx")

        progress.update(task, description="Updating master Excel...")
        master_path = save_master(summary, contractors, url, output / "master_bid_prices.xlsx")

        progress.update(task, description="Creating comparison PDF...")
        pdf_path = save_pdf(summary, contractors, url, output / f"{prefix}_summary.pdf")

        progress.update(task, description="Complete", total=1, completed=1)

    return {"summary": summary_path, "master": master_path, "pdf": pdf_path}


def main() -> dict[str, str]:
    console = Console(file=sys.stdout)
    if len(sys.argv) > 1:
        pdf_url = sys.argv[1]
    else:
        index_url = "https://www.dot.state.tx.us/AVN/avninfo/status/construct/index.htm"
        with console.status("Loading projects..."):
            names_to_link = get_name_to_links(getData(index_url))
        project_names = list(names_to_link)
        if not project_names:
            raise ValueError("No project PDF links were found on the TxDOT page.")
        for number, name in enumerate(project_names, start=1):
            console.print(f"{number}: {name}", markup=False)
        while True:
            try:
                choice = int(input("Enter the project number: "))
                if 1 <= choice <= len(project_names):
                    break
            except ValueError:
                pass
            console.print(f"Enter a number from 1 to {len(project_names)}.")
        pdf_url = urljoin(index_url, names_to_link[project_names[choice - 1]])

    output_folder = sys.argv[2] if len(sys.argv) > 2 else "outputs"
    files = process_pdf(pdf_url, output_folder)
    console.print("Summary Excel:", files["summary"], markup=False)
    console.print("Master Excel:", files["master"], markup=False)
    console.print("Summary PDF:", files["pdf"], markup=False)
    return files


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        Console().print("Cancelled.")
        sys.exit(130)
    except Exception as error:
        Console(stderr=True).print(f"Could not finish: {error}", markup=False)
        sys.exit(1)
