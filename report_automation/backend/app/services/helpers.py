# backend/app/helpers.py
import os
import time
import pandas as pd

REPORTS_DIR = "reports"
os.makedirs(REPORTS_DIR, exist_ok=True)

def timestamp():
    return time.strftime("%Y%m%d_%H%M%S")

def load_csv(path: str) -> pd.DataFrame:
    """Load a CSV file into a pandas DataFrame."""
    return pd.read_csv(path)

def compute_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Compute basic statistical summary of the dataset."""
    return df.describe(include="all")

def save_csv(df: pd.DataFrame, filename: str | None = None) -> str:
    """Save DataFrame as CSV and return file path."""
    if filename is None:
        filename = f"report_{timestamp()}.csv"
    filepath = os.path.join(REPORTS_DIR, filename)
    df.to_csv(filepath, index=False)
    return filepath

def save_excel(df: pd.DataFrame, filename: str | None = None) -> str:
    """Save DataFrame as Excel and return file path."""
    if filename is None:
        filename = f"report_{timestamp()}.xlsx"
    filepath = os.path.join(REPORTS_DIR, filename)
    # pandas will use openpyxl engine for xlsx by default if installed
    df.to_excel(filepath, index=False)
    return filepath

def save_pdf(df: pd.DataFrame, filename: str | None = None) -> str:
    """Save DataFrame summary as a simple PDF and return file path."""
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    if filename is None:
        filename = f"report_{timestamp()}.pdf"
    filepath = os.path.join(REPORTS_DIR, filename)

    # Get summary text (use compute_summary)
    summary_df = compute_summary(df)
    summary = summary_df.to_string()

    # Create PDF
    c = canvas.Canvas(filepath, pagesize=letter)
    textobject = c.beginText(40, 750)  # x, y starting point
    textobject.setFont("Courier", 9)   # monospace to align columns
    for line in summary.split("\n"):
        # If you have very long lines, you might need to wrap them; keep it simple here
        textobject.textLine(line)
        # If the page gets long, add new page (very simple check)
        if textobject.getY() < 40:
            c.drawText(textobject)
            c.showPage()
            textobject = c.beginText(40, 750)
            textobject.setFont("Courier", 9)
    c.drawText(textobject)
    c.showPage()
    c.save()

    return filepath
