# backend/app/api/schemas.py
from pydantic import BaseModel
from typing import Optional, List

# Payload for scheduling a report
class ScheduleRequest(BaseModel):
    report_type: str        # e.g., "csv", "excel", "pdf"
    email: Optional[str]    # optional email to send report
    schedule_time: Optional[str]  # optional, e.g., "2025-08-30 10:00"

# Payload for running a report immediately
class RunRequest(BaseModel):
    report_type: str        # e.g., "csv", "excel", "pdf"

# Response for history (optional, if needed)
class ReportHistoryItem(BaseModel):
    report_id: int
    report_type: str
    created_at: str
    status: str

class HistoryResponse(BaseModel):
    history: List[ReportHistoryItem]
