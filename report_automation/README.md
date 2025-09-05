# Report Automation System

## Overview
This project is a full-stack automated report generation and analytics system. It allows users to upload data, generate reports in multiple formats (CSV, Excel, PDF), visualize analytics, and send reports via email. The system is built with a Python backend (FastAPI), a React frontend, and supports dynamic charting and statistics.

---

## Features
- **Upload Data:** Upload CSV files for report generation.
- **Generate Reports:** Create reports in CSV, Excel, and PDF formats.
- **Analytics Dashboard:** View trends, categories, and format distribution using interactive charts.
- **Email Reports:** Send generated reports to specified recipients.
- **Statistics Tracking:** All report actions update a shared `report_stats.json` for analytics.

---

## Project Structure
```
report_automation/
├── backend/
│   ├── app/
│   │   ├── api/         # API endpoints (FastAPI)
│   │   ├── core/        # Core logic and utilities
│   │   ├── services/    # Report generation and email services
│   │   ├── main.py      # FastAPI app entry point
│   ├── generated_reports/ # Output reports and assets
│   ├── requirements.txt # Python dependencies
│   ├── Dockerfile       # Backend containerization
├── frontend/
│   └── report-frontend/
│       ├── public/      # Static assets (logo, report_stats.json)
│       ├── src/         # React source code
│       ├── package.json # Frontend dependencies
├── configs/             # Configuration files
├── data/                # Sample data
├── reports/             # (Optional) Additional reports
```

---

## Setup Instructions

### Backend (Python/FastAPI)
1. **Install dependencies:**
   ```sh
   cd backend
   pip install -r requirements.txt
   ```
2. **Run the server:**
   ```sh
   uvicorn app.main:app --reload
   ```
3. **Environment variables:**
   - Configure `.env` for email and other secrets.

### Frontend (React)
1. **Install dependencies:**
   ```sh
   cd frontend/report-frontend
   npm install
   ```
2. **Start the frontend:**
   ```sh
   npm start
   ```


---

## Usage
- **Upload a CSV file** via the frontend dashboard.
- **Generate a report** in your chosen format.
- **View analytics** (trends, categories, formats) on the dashboard.
- **Send reports by email** using the provided form.

---

## How Statistics Work
- Every report generation or email action updates `report_stats.json`.
- The frontend reads this file to display up-to-date analytics.

---

## Customization & Extensibility
- Add new report formats by extending backend services.
- Modify chart types or dashboard UI in the React frontend.
- Integrate with other data sources or authentication as needed.

---

## License
This project is for educational and internal use. Please review and update licensing as needed for your organization.

---

