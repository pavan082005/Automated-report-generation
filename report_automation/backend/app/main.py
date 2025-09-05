from app.services import reports         # for report generation logic
from app.api import routes               # for router
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.services.scheduler import start_scheduler
from contextlib import asynccontextmanager

# ✅ Lifespan function to start scheduler
@asynccontextmanager
async def lifespan(app: FastAPI):
    start_scheduler()
    yield

# ✅ Create FastAPI app with lifespan
app = FastAPI(title="Report Automation API", lifespan=lifespan)

# ✅ Add CORS middleware to allow frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],      # For development, allow all origins; replace with frontend URL in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Include reports router at /reports
app.include_router(routes.router, prefix="/reports", tags=["reports"])
