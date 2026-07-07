from fastapi import FastAPI
from pydantic import BaseModel
from typing import List

from ppt_generator import generate_ppt

app = FastAPI(
    title="BLIPS PowerPoint Generator",
    description="Accepts BLIPS JSON and generates a PowerPoint.",
    version="1.0.0"
)


class Action(BaseModel):
    owner: str
    action: str


class GeneratePptRequest(BaseModel):
    title: str
    subtitle: str
    summary: str
    primaryFailure: str
    recommendation: str
    keyFocusAreas: List[str]
    actions: List[Action]


@app.get("/")
def home():
    return {
        "status": "running",
        "message": "Use POST /generate-ppt or open /docs to test the API."
    }

@app.post("/generate-ppt")
async def generate(request: GeneratePptRequest):

    data = request.model_dump()

    output_path = "generated_blips.pptx"

    generate_ppt(data, output_path)

    return {
        "status": "success",
        "ppt_file": output_path
    }
