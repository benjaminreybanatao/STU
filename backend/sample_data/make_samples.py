"""Generates a synthetic OM PDF + underwriting model for end-to-end testing.
Not part of the app itself -- just test fixtures."""
import fitz
from PIL import Image, ImageDraw
import openpyxl
from pathlib import Path

OUT = Path(__file__).parent

# --- Fake "hero photo" (large, photographic aspect ratio) ---
hero = Image.new("RGB", (1600, 1000), (90, 110, 140))
d = ImageDraw.Draw(hero)
for i in range(0, 1600, 40):
    d.line([(i, 0), (i, 1000)], fill=(80 + i % 60, 100, 130), width=20)
d.rectangle([100, 600, 1500, 950], fill=(60, 60, 65))
hero.save(OUT / "hero.jpg", quality=85)

# --- Fake logo (thin banner, should be filtered out) ---
logo = Image.new("RGB", (400, 60), (255, 255, 255))
d2 = ImageDraw.Draw(logo)
d2.text((10, 20), "ACME BROKERAGE LOGO", fill=(0, 0, 0))
logo.save(OUT / "logo.png")

# --- Fake floor plan (small, near-square line art) ---
plan = Image.new("RGB", (250, 250), (245, 245, 245))
d3 = ImageDraw.Draw(plan)
d3.rectangle([20, 20, 230, 230], outline=(0, 0, 0), width=2)
plan.save(OUT / "floorplan.png")

doc = fitz.open()
page = doc.new_page(width=612, height=792)

text = (
    "1200 Market Street, Austin, TX 78701\n\n"
    "Offering Memorandum\n\n"
    "1200 Market Street is a 145,000 square feet office property built in 2016, "
    "located in the Austin CBD submarket, a high-growth technology submarket with "
    "strong absorption over the trailing 24 months.\n\n"
    "The property is 94% leased. Tenants include a mix of technology and professional "
    "services firms anchored by a 12-year investment-grade tenant occupying 40% of the NRA.\n\n"
)
page.insert_textbox(fitz.Rect(50, 50, 560, 300), text, fontsize=11)

page.insert_image(fitz.Rect(50, 320, 560, 620), filename=str(OUT / "hero.jpg"))
page.insert_image(fitz.Rect(50, 630, 200, 660), filename=str(OUT / "logo.png"))

page2 = doc.new_page(width=612, height=792)
page2.insert_textbox(fitz.Rect(50, 50, 560, 100), "Site Plan", fontsize=14)
page2.insert_image(fitz.Rect(180, 120, 430, 370), filename=str(OUT / "floorplan.png"))

doc.save(OUT / "sample_om.pdf")
doc.close()

# --- Underwriting model ---
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Base Case"

rows = [
    ("Purchase Price", 42_500_000),
    ("Total Purchase Price", 42_500_000),
    ("Price / SF", 293.10),
    ("Going-In Cap Rate", 0.0525),
    ("Exit Cap Rate", 0.0575),
    ("Leverage", 0.60),
    ("Interest Rate", 0.055),
    ("Initial Equity", 17_000_000),
    ("Peak Equity", 17_850_000),
    ("Hold Period", 5),
    ("Unlevered IRR", 0.079),
    ("Levered IRR", 0.134),
    ("Unlevered Equity Multiple", 1.42),
    ("Levered Equity Multiple", 1.71),
    ("Cash-on-Cash", 0.061),
    ("Lease Term Assumption", "10-year weighted average"),
    ("Downtime Assumption", "9 months between leases"),
    ("Exit Assumption", "Sale in Year 5 at stabilized NOI"),
]
for i, (label, value) in enumerate(rows, start=2):
    ws.cell(row=i, column=1, value=label)
    ws.cell(row=i, column=2, value=value)

ws2 = wb.create_sheet("Downside Case")
downside_rows = [
    ("Exit Cap Rate", 0.065),
    ("Levered IRR", 0.081),
    ("Unlevered IRR", 0.052),
    ("Levered Equity Multiple", 1.32),
    ("Unlevered Equity Multiple", 1.21),
    ("Cash-on-Cash", 0.041),
]
for i, (label, value) in enumerate(downside_rows, start=2):
    ws2.cell(row=i, column=1, value=label)
    ws2.cell(row=i, column=2, value=value)

wb.save(OUT / "sample_model.xlsx")

print("Wrote sample_om.pdf and sample_model.xlsx to", OUT)
