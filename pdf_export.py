from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from xml.sax.saxutils import escape


def save_pdf(summary: dict, contractors: list[dict], source_url: str, path: Path) -> str:
    """Create the printable bid comparison and return its absolute path."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from pdf_extractor import clean_name, validate_bids

    validate_bids(contractors)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp.pdf")
    page_size = landscape(A3)
    width = page_size[0] - 64
    document = SimpleDocTemplate(
        str(temporary), pagesize=page_size, rightMargin=32, leftMargin=32,
        topMargin=30, bottomMargin=35, title="TxDOT aviation bid comparison",
    )
    styles = getSampleStyleSheet()
    text_style = ParagraphStyle("Cell", fontName="Helvetica", fontSize=9, leading=12, alignment=TA_LEFT)
    heading_style = ParagraphStyle("Heading cell", parent=text_style, fontName="Helvetica-Bold", textColor=colors.white)
    number_style = ParagraphStyle("Number cell", parent=text_style, alignment=TA_RIGHT)
    small_style = ParagraphStyle("Source", parent=text_style, fontSize=8, leading=11)

    def paragraph(value: object, heading: bool = False, number: bool = False) -> Paragraph:
        style = heading_style if heading else number_style if number else text_style
        return Paragraph(escape(str(value)), style)

    def money(value: Decimal) -> Paragraph:
        return paragraph(f"${value:,.2f}", number=True)

    def footer(canvas: object, doc: object) -> None:
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#536572"))
        canvas.drawString(32, 18, "TxDOT aviation - Bid comparison")
        canvas.drawRightString(page_size[0] - 32, 18, f"Page {doc.page}")
        canvas.restoreState()

    project = summary["project"]
    lowest = min(contractor["base_bid"] for contractor in contractors)
    winners = [contractor["name"] for contractor in contractors if contractor["base_bid"] == lowest]
    story = []

    # Repeat shared item columns on additional sheets if there are many bidders.
    for start in range(0, len(contractors), 3):
        group = contractors[start:start + 3]
        if start:
            story.append(PageBreak())
        story.append(Paragraph(escape(project["Airport name"]) + " - Bid comparison", styles["Title"]))
        story.append(paragraph(f"CSJ: {project['TxDOT CSJ#']}    |    Bid date: {project['Bid Opening Date']}    |    Scope: {project['Project Scope']}"))
        label = "Apparent lowest bidders (tie)" if len(winners) > 1 else "Apparent lowest bidder"
        story.append(paragraph(f"{label}: {', '.join(winners)} - ${lowest:,.2f}"))
        story.append(paragraph("Lowest bid is not a confirmed award. Amounts below are reported bid prices."))
        if len(contractors) > 3:
            story.append(paragraph(f"Bidders {start + 1}-{start + len(group)} of {len(contractors)}; shared item columns repeat."))
        story.append(Spacer(1, 14))

        first_header = [paragraph(value, heading=True) for value in ("Item", "Description", "Qty", "Unit")]
        second_header = ["", "", "", ""]
        commands = [
            ("BACKGROUND", (0, 0), (-1, 1), colors.HexColor("#234E70")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LINEBELOW", (0, 1), (-1, 1), 0.6, colors.HexColor("#234E70")),
        ]
        for column in range(4):
            commands.append(("SPAN", (column, 0), (column, 1)))
        for index, contractor in enumerate(group):
            column = 4 + 2 * index
            first_header.extend([paragraph(contractor["name"], heading=True), ""])
            second_header.extend([paragraph("Unit price", heading=True), paragraph("Total price", heading=True)])
            commands.append(("SPAN", (column, 0), (column + 1, 0)))
            if contractor["base_bid"] == lowest:
                commands.append(("BACKGROUND", (column, 0), (column + 1, 1), colors.HexColor("#27745C")))

        rows = [first_header, second_header]
        lookup = [{clean_name(item["Item"]): item for item in contractor["items"]} for contractor in group]
        for item in contractors[0]["items"]:
            row = [paragraph(item["Item"]), paragraph(item["Description"]), paragraph(f"{item['Qty']:,}", number=True), paragraph(item["Unit"])]
            for items in lookup:
                bid = items[clean_name(item["Item"])]
                row.extend([money(bid["Unit Price"]), money(bid["Total Price"])])
            rows.append(row)
            if len(rows) % 2:
                commands.append(("BACKGROUND", (0, len(rows) - 1), (-1, len(rows) - 1), colors.HexColor("#F2F6F9")))

        for label in ("Calculated item totals", "Reported base bids", "Difference"):
            row = [paragraph(label), "", "", ""]
            for contractor in group:
                calculated = sum((item["Total Price"] for item in contractor["items"]), Decimal("0"))
                value = calculated if label == "Calculated item totals" else contractor["base_bid"]
                if label == "Difference":
                    value = calculated - contractor["base_bid"]
                row.extend(["", money(value)])
            rows.append(row)
            commands.append(("SPAN", (0, len(rows) - 1), (3, len(rows) - 1)))
        for index, contractor in enumerate(group):
            if contractor["base_bid"] == lowest:
                commands.append(("BACKGROUND", (4 + index * 2, len(rows) - 2), (5 + index * 2, len(rows) - 2), colors.HexColor("#DDEFE5")))
        commands.append(("LINEABOVE", (0, len(rows) - 3), (-1, len(rows) - 3), 0.8, colors.HexColor("#234E70")))
        fixed_widths = [65, 265, 65, 40]
        price_width = (width - sum(fixed_widths)) / (2 * len(group))
        table = Table(rows, colWidths=fixed_widths + [price_width] * (2 * len(group)), repeatRows=2, hAlign="LEFT")
        table.setStyle(TableStyle(commands))
        story.extend([table, Spacer(1, 12)])
        story.append(Paragraph("Source PDF: " + escape(source_url), small_style))
    try:
        document.build(story, onFirstPage=footer, onLaterPages=footer)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
    return str(path.resolve())
