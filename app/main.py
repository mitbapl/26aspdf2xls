from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from utils import process_pdf_to_excel
import os

app = FastAPI()

UPLOAD_FOLDER = "uploaded_files"
OUTPUT_FOLDER = "output_files"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.post("/upload/")
async def upload_file(file: UploadFile):
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    with open(file_path, "wb") as f:
        f.write(await file.read())

    # Process the uploaded PDF to extract data and generate Excel
    output_excel = os.path.join(OUTPUT_FOLDER, f"{file.filename}.xlsx")
    process_pdf_to_excel(file_path, output_excel)

    return {"download_url": f"/download/{os.path.basename(output_excel)}"}

@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return {"error": "File not found"}
