from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from pathlib import Path
from .utils import process_pdf_to_excel
import os

# Initialize FastAPI app
app = FastAPI()

# Set directories for uploads and outputs
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploaded_files"
OUTPUT_FOLDER = BASE_DIR / "output_files"
TEMPLATES_FOLDER = BASE_DIR / "templates"

# Ensure directories exist
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# Set up templates
templates = Jinja2Templates(directory=TEMPLATES_FOLDER)

@app.get("/", response_class=HTMLResponse)
async def home():
    """
    Serve the homepage.
    """
    return templates.TemplateResponse("index.html", {"request": {}})

@app.post("/upload/")
async def upload_file(file: UploadFile):
    """
    Handle PDF file upload and convert it to Excel.
    """
    # Save the uploaded file
    file_path = UPLOAD_FOLDER / file.filename
    with file_path.open("wb") as f:
        f.write(await file.read())

    # Generate the Excel file
    output_filename = f"{file.filename.rsplit('.', 1)[0]}.xlsx"  # Remove .pdf extension
    output_excel = OUTPUT_FOLDER / output_filename
    process_pdf_to_excel(file_path, output_excel)

    # Return the download URL
    return {"download_url": f"/download/{output_excel.name}"}

@app.get("/download/{filename}")
async def download_file(filename: str):
    """
    Serve the generated Excel file.
    """
    file_path = OUTPUT_FOLDER / filename
    print(file_path)
    if file_path.exists():
        return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return {"error": "File not found"}
