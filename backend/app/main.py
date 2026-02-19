from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routers.reportes import router as reportes_router

app = FastAPI(title="Backend Call Center", version="1.0.0")

# ==========================================
# CORS - PERMITIR TODOS LOS ORÍGENES
# ==========================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # ← CAMBIAR ESTO
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(reportes_router, prefix="/api")

@app.get("/")
def root():
    return {"status": "Backend OK", "cors_enabled": True}