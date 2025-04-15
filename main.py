
from fastapi import FastAPI, File, UploadFile, Form, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
from io import BytesIO

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("upload_form.html", {"request": request})

@app.post("/upload/")
async def upload_files(
    request: Request,
    file_source: UploadFile = File(...),
    file_target: UploadFile = File(...),
    sheet_name: str = Form(...)
):
    source_df = pd.read_excel(await file_source.read())
    target_excel = pd.ExcelFile(await file_target.read())

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    for sheet in target_excel.sheet_names:
        df = target_excel.parse(sheet)
        if sheet == sheet_name:
            df = pd.concat([df, source_df], ignore_index=True)
        df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()
    output.seek(0)

    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": "attachment; filename=modified_target.xlsx"}
    )
