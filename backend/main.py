import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from api.process import router
from dotenv import load_dotenv
load_dotenv()

app = FastAPI(title="Webenoid AI Query Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router)

# Serve icon files directly (bypasses ngrok interstitial)
@app.get("/icon/{filename}")
def serve_icon(filename: str):
    icon_path = os.path.join(os.path.dirname(__file__), "addin", "assets", filename)
    if os.path.exists(icon_path):
        return FileResponse(icon_path, media_type="image/png")
    return {"error": "Icon not found"}

# Serve the Excel add-in static files (HTML, JS, CSS, assets)
addin_path = os.path.join(os.path.dirname(__file__), "addin")
if os.path.exists(addin_path):
    app.mount("/addin", StaticFiles(directory=addin_path, html=True), name="addin")

@app.get("/")
def root():
    return {"message": "Webenoid AI Running 🚀"}
