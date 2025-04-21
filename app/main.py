from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from pathlib import Path
from .utils import process_pdf_to_excel

app = FastAPI()

# Set directories
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploaded_files"
OUTPUT_FOLDER = BASE_DIR / "output_files"
TEMPLATES_FOLDER = BASE_DIR / "templates"

# Ensure directories exist
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

templates = Jinja2Templates(directory=TEMPLATES_FOLDER)

@app.get("/", response_class=HTMLResponse)
async def home(request: Request, success: str = "", filename: str = ""):
    """
    Serve homepage with optional success message.
    """
    return templates.TemplateResponse("index.html", {
        "request": request,
        "success": success,
        "filename": filename
    })

@app.head("/")
async def head_home():
    """
    Handle HEAD request to root and return 200 OK
    """
    return HTMLResponse(status_code=200)

@app.post("/upload/")
async def upload_file(file: UploadFile):
    """
    Handle PDF upload and convert to Excel.
    """
    if not file.filename.endswith(".pdf"):
        return {"error": "Only PDF files are allowed."}

    file_path = UPLOAD_FOLDER / file.filename
    with file_path.open("wb") as f:
        f.write(await file.read())

    output_filename = f"{file.filename.rsplit('.', 1)[0]}.xlsx"
    output_excel = OUTPUT_FOLDER / output_filename

    process_pdf_to_excel(file_path, output_excel)

    return {"download_url": f"/download/{output_filename}"}

@app.get("/download/{filename}")
async def download_file(filename: str):
    """
    Serve the generated Excel file.
    """
    file_path = OUTPUT_FOLDER / filename
    if file_path.exists():
        headers = {
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
        }
        return FileResponse(
            path=file_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
            headers=headers
        )
    return {"error": "File not found"}
