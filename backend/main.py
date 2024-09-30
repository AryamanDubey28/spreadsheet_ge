from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],  #TODO:Verify this is the frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    contents = await file.read()
    df = pd.read_excel(BytesIO(contents))
    
    # Perform analysis on the DataFrame
    analysis_result = {
        "num_rows": len(df),
        "num_columns": len(df.columns),
        "column_names": df.columns.tolist(),
        "summary_stats": df.describe().to_dict(),
    }
    
    return analysis_result

@app.get("/")
async def root():
    return {"message": "Hello World"}
