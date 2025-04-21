from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, JSONResponse
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
    Handle HEAD request to root.
    """
    return HTMLResponse(status_code=200)

@app.post("/upload/")
async def upload_file(file: UploadFile):
    """
    Receive PDF, process, and redirect to download.
    """
    if not file.filename.endswith(".pdf"):
        return JSONResponse(content={"error": "Only PDF files are allowed."}, status_code=400)

    file_path = UPLOAD_FOLDER / file.filename
    with file_path.open("wb") as f:
        f.write(await file.read())

    output_filename = f"{file.filename.rsplit('.', 1)[0]}.xlsx"
    output_excel = OUTPUT_FOLDER / output_filename

    # Log file paths for debugging
    print(f"Uploaded PDF path: {file_path}")
    print(f"Output Excel path: {output_excel}")

    try:
        print(f"Processing {file_path} and saving to {output_excel}")
        process_pdf_to_excel(file_path, output_excel)
    except Exception as e:
        print(f"Error during file processing: {e}")
        return JSONResponse(content={"error": f"File generation failed: {str(e)}"}, status_code=500)

    # Confirm file was created
    if output_excel.exists():
        print(f"File successfully created at: {output_excel}")
        return RedirectResponse(url=f"/download/{output_filename}", status_code=303)
    else:
        print("File was not created.")
        return JSONResponse(content={"error": "File generation failed"}, status_code=500)

@app.get("/download/{filename}")
async def download_file(filename: str):
    """
    Serve the generated Excel file.
    """
    file_path = OUTPUT_FOLDER / filename
    print(f"Looking for file: {file_path}")
    
    if file_path.is_file():
        return FileResponse(
            path=str(file_path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
            headers={
                "Cache-Control": "no-cache, no-store, must-revalidate",
                "Pragma": "no-cache",
                "Expires": "0",
            }
        )
    else:
        print("File not found.")
        return JSONResponse(content={"error": "File not found"}, status_code=404)
