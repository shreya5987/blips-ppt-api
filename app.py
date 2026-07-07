from datetime import datetime
from pathlib import Path
import tempfile
import re

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

from ppt_generator import generate_ppt_from_json
from sharepoint_graph import upload_file_to_sharepoint

app = FastAPI(
    title="BLIPS PowerPoint Generator",
    description="Accepts BLIPS JSON, generates a PowerPoint, uploads it to SharePoint, and returns the SharePoint link.",
    version="1.0.0",
)


class GeneratePptRequest(BaseModel):
    title: str = Field(..., description="Title for the generated PowerPoint")
    problem_statement: str = Field(
        ..., description="Problem statement or BLIPS summary"
    )

    # Additional flexible fields
    data: dict = Field(default_factory=dict)


def safe_filename(text: str) -> str:
    text = re.sub(r"[^a-zA-Z0-9_-]+", "_", text)
    return text.strip("_")[:80]


@app.get("/")
def home():
    return {
        "status": "running",
        "message": "Use POST /generate-ppt or open /docs to test the API.",
    }


@app.post("/generate-ppt")
def generate_ppt(request: GeneratePptRequest):
    try:
        # Combine standard fields with additional JSON fields
        payload = {
            "title": request.title,
            "problem_statement": request.problem_statement,
            **request.data,
        }

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = safe_filename(request.title)
        output_file_name = f"{base_name}_{timestamp}.pptx"

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / output_file_name

            # Generate PowerPoint
            generate_ppt_from_json(
                payload,
                str(output_path)
            )

            # Upload to SharePoint
            uploaded_file = upload_file_to_sharepoint(
                local_file_path=str(output_path),
                target_file_name=output_file_name,
            )

        return {
            "status": "success",
            "message": "PowerPoint generated and uploaded to SharePoint.",
            "fileName": uploaded_file["name"],
            "sharePointUrl": uploaded_file["webUrl"],
            "driveItemId": uploaded_file["id"],
        }

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=str(e),
        )