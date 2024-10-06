from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"], # Adjust this to match your Vue.js dev server
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    contents = await file.read()
    try:
        if file.filename.endswith('.xlsx'):
            xls = pd.ExcelFile(BytesIO(contents))
            sheets_data = {}
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                sheets_data[sheet_name] = {
                    "column_names": df.columns.tolist(),
                    "data": df.to_dict(orient='records')
                }
            response = {
                "sheets": sheets_data,
                "sheet_names": xls.sheet_names
            }
        elif file.filename.endswith('.csv'):
            df = pd.read_csv(BytesIO(contents))
            response = {
                "sheets": {
                    "Sheet1": {
                        "column_names": df.columns.tolist(),
                        "data": df.to_dict(orient='records')
                    }
                },
                "sheet_names": ["Sheet1"]
            }
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format")
        print(f"response = {response}")
        return response
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not process file: {str(e)}")

@app.get("/")
async def root():
    return {"message": "Hello World"}