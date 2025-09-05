from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime
from .reports import generate_csv, generate_excel, generate_pdf
import pandas as pd
import os

DATA_DIR = "e:/report_generation/report_automation/backend/generated_reports"
SAMPLE_FILE = "e:/report_generation/report_automation/data/sample_data.csv"

def scheduled_report(report_type):
    class DummyUploadFile:
        def __init__(self, file_path):
            self.file = open(file_path, "rb")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_obj = DummyUploadFile(SAMPLE_FILE)
    if report_type == "csv":
        output_file = os.path.join(DATA_DIR, f"report_{timestamp}.csv")
        df = pd.read_csv(file_obj.file)
        df.to_csv(output_file, index=False)
    elif report_type == "excel":
        output_file = os.path.join(DATA_DIR, f"report_{timestamp}.xlsx")
        df = pd.read_csv(file_obj.file)
        df.to_excel(output_file, index=False)
    elif report_type == "pdf":
        output_file = os.path.join(DATA_DIR, f"report_{timestamp}.pdf")
        # You may need to modify generate_pdf to accept output_file as a parameter
        generate_pdf(file_obj)  # If needed, adjust to save with timestamp
    file_obj.file.close()
    print(f"{report_type.upper()} report generated: {output_file}")

def start_scheduler():
    scheduler = BackgroundScheduler()
    scheduler.add_job(lambda: scheduled_report("csv"), 'interval', days=1, id='daily_csv')
    scheduler.add_job(lambda: scheduled_report("excel"), 'interval', weeks=1, id='weekly_excel')
    scheduler.add_job(lambda: scheduled_report("pdf"), 'interval', weeks=4, id='monthly_pdf')
    scheduler.start()
    print("Scheduler started.")