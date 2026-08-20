"""
StampScan 2.0 - Hugging Face Space (Gradio)
Upload a photo -> Roboflow Serverless Workflow -> numbered preview + table + Excel
The Roboflow API key is read from the Space secret ROBOFLOW_API_KEY (never in the browser).
"""

import os
import io
import base64
import tempfile

import gradio as gr
from inference_sdk import InferenceHTTPClient

import openpyxl
from openpyxl.styles import Font, PatternFill

# ---- Config -----------------------------------------------------------------
API_KEY = os.environ.get("ROBOFLOW_API_KEY", "")
WORKSPACE = "vidcas567-gmail-com"
WORKFLOW_ID = "automated-stamp-scanner-1787234733118"
API_URL = "https://serverless.roboflow.com"

client = InferenceHTTPClient(api_url=API_URL, api_key=API_KEY) if API_KEY else None


# ---- Helpers ----------------------------------------------------------------
def _decode_output_image(value):
    """Roboflow may return the annotated image as base64, a data URL, or a dict."""
    if not value:
        return None
    try:
        if isinstance(value, dict):
            value = value.get("value") or value.get("image") or ""
        if not isinstance(value, str):
            return None
        if value.startswith("data:image"):
            value = value.split(",", 1)[1]
        raw = base64.b64decode(value)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.write(raw)
        tmp.close()
        return tmp.name
    except Exception:
        return None


def _text_from(entry):
    if isinstance(entry, dict):
        for k in ("text", "ocr", "value", "denomination"):
            if entry.get(k):
                return str(entry[k])
    elif isinstance(entry, str):
        return entry
    return ""


def _condition_from(entry):
    if isinstance(entry, dict):
        parts = []
        for k in ("condition", "damage", "cancellation", "centering"):
            if entry.get(k):
                parts.append(f"{k}: {entry[k]}")
        return ", ".join(parts)
    return ""


def _build_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stamps"
    headers = ["#", "Type", "Confidence", "Text / OCR", "Condition"]
    ws.append(headers)
    fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
    for r in rows:
        ws.append(r)
    for col in ws.columns:
        width = max(len(str(c.value or "")) for c in col) + 2
        ws.column_dimensions[col[0].column_letter].width = min(width, 50)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


# ---- Main scan --------------------------------------------------------------
def scan(image_path):
    if not API_KEY or client is None:
        return None, [], "⚠️ ROBOFLOW_API_KEY is not set in this Space's secrets.", None
    if not image_path:
        return None, [], "Please upload a photo first.", None

    try:
        result = client.run_workflow(
            workspace_name=WORKSPACE,
            workflow_id=WORKFLOW_ID,
            images={"image": image_path},
            use_cache=True,
        )
    except Exception as e:
        return None, [], f"Error calling the workflow: {e}", None

    scan_out = result[0] if result else {}

    predictions = scan_out.get("predictions", []) or []
    ocr_results = scan_out.get("ocr_results", []) or []
    analysis = scan_out.get("stamp_analysis", []) or []
    count = scan_out.get("stamp_count", None)

    n = max(len(predictions), len(ocr_results), len(analysis))
    if count is None:
        count = n

    rows = []
    for i in range(n):
        pred = predictions[i] if i < len(predictions) else {}
        ocr = ocr_results[i] if i < len(ocr_results) else {}
        cond = analysis[i] if i < len(analysis) else {}
        conf = pred.get("confidence", "") if isinstance(pred, dict) else ""
        if isinstance(conf, (int, float)):
            conf = f"{conf:.2f}"
        cls = pred.get("class", "stamp") if isinstance(pred, dict) else "stamp"
        rows.append([i + 1, cls, conf, _text_from(ocr), _condition_from(cond)])

    annotated = _decode_output_image(scan_out.get("output_image"))
    xlsx = _build_excel(rows) if rows else None
    summary = f"Found {count} stamp(s)."

    return annotated, rows, summary, xlsx


# ---- UI ---------------------------------------------------------------------
demo = gr.Interface(
    fn=scan,
    inputs=gr.Image(type="filepath", label="Upload an album-page photo"),
    outputs=[
        gr.Image(label="Detected stamps (numbered)"),
        gr.Dataframe(
            headers=["#", "Type", "Confidence", "Text / OCR", "Condition"],
            label="Results (review before trusting)",
            wrap=True,
        ),
        gr.Textbox(label="Summary"),
        gr.File(label="Download Excel"),
    ],
    title="StampScan 2.0",
    description=(
        "Upload a photo of a stamp album page. It detects each stamp, reads the "
        "printed text, suggests an identification, and checks visible condition. "
        "Counts on very dense pages are still approximate — always review the numbered preview."
    ),
    allow_flagging="never",
)

if __name__ == "__main__":
    demo.launch(
        server_name="0.0.0.0",
        server_port=int(os.environ.get("PORT", 7860)),
    )
