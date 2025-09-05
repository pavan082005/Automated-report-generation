from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from app.services import reports

router = APIRouter()

# ✅ Test endpoint
@router.get("/test")
def test_reports():
    return {"message": "Reports router works!"}

# ✅ Generate CSV
@router.post("/generate/csv")
async def generate_csv(file: UploadFile = File(...)):
    output_file = reports.generate_csv(file)
    return FileResponse(output_file, media_type="text/csv", filename="report.csv")

# ✅ Generate Excel
@router.post("/generate/excel")
async def generate_excel(file: UploadFile = File(...)):
    output_file = reports.generate_excel(file)
    return FileResponse(output_file, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="report.xlsx")

# ✅ Generate PDF
@router.post("/generate/pdf")
async def generate_pdf(file: UploadFile = File(...)):
    output_file = reports.generate_pdf(file)
    return FileResponse(output_file, media_type="application/pdf", filename="report.pdf")

# ✅ Send CSV via email
@router.post("/send/csv")
async def send_csv(to_email: str):
    success = reports.send_email_with_report(to_email, "csv")
    if not success:
        raise HTTPException(status_code=500, detail="Failed to send CSV")
    return {"message": f"CSV sent to {to_email}"}

# ✅ Send Excel via email
@router.post("/send/excel")
async def send_excel(to_email: str):
    success = reports.send_email_with_report(to_email, "excel")
    if not success:
        raise HTTPException(status_code=500, detail="Failed to send Excel")
    return {"message": f"Excel sent to {to_email}"}

# ✅ Send PDF via email
@router.post("/send/pdf")
async def send_pdf(to_email: str):
    success = reports.send_email_with_report(to_email, "pdf")
    if not success:
        raise HTTPException(status_code=500, detail="Failed to send PDF")
    return {"message": f"PDF sent to {to_email}"}

@router.post("/send/sample")
async def send_sample(to_email: str):
    success = reports.send_sample_message(to_email)
    if not success:
        raise HTTPException(status_code=500, detail="Failed to send sample message")
    return {"message": f"Sample message sent to {to_email}"}