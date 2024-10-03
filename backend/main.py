from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    print(f"---> in upload_excel")
    contents = await file.read()
    try:
        if file.filename.endswith('.xlsx'):
            df = pd.read_excel(BytesIO(contents), header=None)
        elif file.filename.endswith('.csv'):
            df = pd.read_csv(BytesIO(contents))
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format")
    except Exception as e:
        print(f"Error: {str(e)}")
        raise HTTPException(status_code=400, detail=f"Could not process file: {str(e)}")
    
    # Perform analysis on the DataFrame
    print(f"df = {df}")
    analysis_result = {
        "num_rows": len(df),
        "num_columns": len(df.columns),
        "column_names": df.columns.tolist(),
        "data": df.to_dict(orient='records')
    }
    print(f"Analysis = {analysis_result}")
    
    return analysis_result

@app.get("/")
async def root():
    return {"message": "Hello World"}