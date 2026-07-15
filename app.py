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


@app.post("/leerpdf")
def procesar_pdf(data: PDFRequest):

    return leer_pdf(data.pdf)