from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],  # Adjust this to match your Vue.js dev server
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    contents = await file.read()
    
    try:
        if file.filename.endswith('.xlsx'):
            df = pd.read_excel(BytesIO(contents))
        elif file.filename.endswith('.csv'):
            df = pd.read_csv(BytesIO(contents))
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format")
        
        # Convert DataFrame to list of dictionaries
        data = df.to_dict(orient='records')
        
        # Prepare the response
        response = {
            "column_names": df.columns.tolist(),
            "data": data
        }
        
        return response
    
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not process file: {str(e)}")

@app.get("/")
async def root():
    return {"message": "Hello World"}