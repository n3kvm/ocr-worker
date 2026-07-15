from fastapi import FastAPI
from pydantic import BaseModel

from parser import leer_pdf

app = FastAPI()


class PDFRequest(BaseModel):
    pdf: str


@app.get("/")
def home():

    return {
        "status": "online"
    }


from fastapi import HTTPException

@app.post("/leerpdf")
def procesar_pdf(data: PDFRequest):
    try:
        return leer_pdf(data.pdf)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))