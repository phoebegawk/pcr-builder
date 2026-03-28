from io import BytesIO
from pathlib import Path
from datetime import datetime
import re

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from pptx import Presentation
import uvicorn

from rep_data import REP_DATA, DROPDOWN_REPS

app = FastAPI(title="PCR Report Builder")

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
TEMPLATES_DIR = BASE_DIR / "templates"
PPTX_TEMPLATE_PATH = TEMPLATES_DIR / "PCR - Template.pptx"

app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")


# -----------------------------
# Helpers
# -----------------------------
def clean_filename_part(value: str) -> str:
    value = value.strip()
    value = re.sub(r"[^A-Za-z0-9 _-]", "", value)
    value = re.sub(r"\s+", " ", value)
    return value or "Client"


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip())


def extract_month_year_from_excel(file_bytes: bytes) -> str:
    workbook = load_workbook(filename=BytesIO(file_bytes), data_only=True)

    for sheet in workbook.worksheets:
        ended_cell = None

        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip().upper() == "ENDED":
                    ended_cell = cell
                    break
            if ended_cell:
                break

        if not ended_cell:
            continue

        date_cell = sheet.cell(row=ended_cell.row + 1, column=ended_cell.column)
        date_value = date_cell.value

        if isinstance(date_value, datetime):
            return date_value.strftime("%B %Y")

        if hasattr(date_value, "strftime"):
            return date_value.strftime("%B %Y")

        if isinstance(date_value, str) and date_value.strip():
            parsed = _parse_date_string(date_value.strip())
            if parsed:
                return parsed.strftime("%B %Y")

        raise ValueError("I found the 'ENDED' header, but the cell below it wasn't a valid date.")

    raise ValueError("Couldn't find an 'ENDED' header in the uploaded Excel file.")


def _parse_date_string(value: str):
    formats = [
        "%d %B %Y",
        "%d %b %Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return None


def replace_text_on_slide(slide, replacements: dict):
    normalized_replacements = {
        normalize_text(key): value for key, value in replacements.items()
    }

    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or not shape.has_text_frame:
            continue

        current_text = normalize_text(shape.text)
        if current_text not in normalized_replacements:
            continue

        new_text = normalized_replacements[current_text]

        text_frame = shape.text_frame
        if not text_frame.paragraphs:
            text_frame.text = new_text
            continue

        first_paragraph = text_frame.paragraphs[0]
        if first_paragraph.runs:
            first_paragraph.runs[0].text = new_text
            for run in first_paragraph.runs[1:]:
                run.text = ""
        else:
            first_paragraph.text = new_text

        for paragraph in text_frame.paragraphs[1:]:
            for run in paragraph.runs:
                run.text = ""
            if not paragraph.runs:
                paragraph.text = ""


def build_pcr_pptx(client_name: str, month_year: str, sales_rep: str) -> BytesIO:
    if not PPTX_TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found at: {PPTX_TEMPLATE_PATH}")

    if sales_rep not in REP_DATA:
        raise ValueError("Selected sales rep wasn't recognised.")

    rep = REP_DATA[sales_rep]
    prs = Presentation(str(PPTX_TEMPLATE_PATH))

    if not prs.slides:
        raise ValueError("The PPTX template has no slides.")

    cover_slide = prs.slides[0]
    replace_text_on_slide(
        cover_slide,
        {
            "Client Name": client_name,
            "Month Year": month_year,
        },
    )

    rep_contact_line = f'{rep["phone"]} | {rep["email"]}'

    last_slide = prs.slides[-1]
    replace_text_on_slide(
        last_slide,
        {
            "Rep Name!": rep["display_name"],
            "Rep Number | Rep Email": rep_contact_line,
            "Rep Number   |   Rep Email": rep_contact_line,
            '"Rep Number" | Rep Email': rep_contact_line,
            '"Rep Number"   |   Rep Email': rep_contact_line,
        },
    )

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# -----------------------------
# Routes
# -----------------------------
@app.get("/", response_class=HTMLResponse)
async def home():
    options_html = "".join(
        f'<option value="{rep}">{rep}</option>' for rep in DROPDOWN_REPS
    )

    return f"""
    <html>
    <head>
        <title>PCR Builder</title>
        <link rel="icon" type="image/png" href="/assets/favicon.png">
        <link rel="stylesheet" href="/assets/styles.css">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            .client-name-input {{
                width: 100%;
                padding: 12px 14px;
                border-radius: 8px;
                border: 2px solid #D7DF23;
                background: #FFFFFF;
                color: #542D54;
                font-size: 16px;
                font-weight: 600;
            }}

            .field-note {{
                width: 100%;
                color: #542D54;
                font-size: 13px;
                font-weight: 600;
                opacity: 0.9;
                margin-top: 6px;
            }}

            .section-label {{
                width: 100%;
                color: #542D54;
                font-size: 24px;
                font-weight: 700;
                margin: 0 0 10px 0;
                line-height: 1.1;
            }}

            .gawk-button {{
                border: none;
                border-radius: 999px;
                padding: 14px 28px;
                font-size: 16px;
                font-weight: 700;
                cursor: pointer;
                transition: transform 0.15s ease, opacity 0.15s ease;
            }}

            .gawk-button:hover {{
                transform: translateY(-1px);
            }}

            .browse-btn {{
                background: #D7DF23;
                color: #542D54;
            }}

            .actions-row {{
                width: 100%;
                display: flex;
                justify-content: center;
                gap: 14px;
                margin: 8px auto 0 auto;
                flex-wrap: wrap;
            }}

            .build-button {{
                background: linear-gradient(90deg, #d9b3db 0%, #8b7f40 100%);
                color: #542D54;
                min-width: 170px;
            }}

            .reset-button {{
                display: block;
                margin: 24px auto 0 auto;
                background: #FFFFFF;
                color: #542D54;
                border: 3px solid #D7DF23;
                min-width: 140px;
            }}

            .upload-confirm-text {{
                color: #542D54;
                font-weight: 700;
            }}

            .page-subtext {{
                width: 100%;
                max-width: 1200px;
                margin: -18px auto 24px auto;
                text-align: center;
                color: #FFFFFF;
            }}

            .page-subtext-primary {{
                font-weight: 700;
                font-size: 16px;
                margin-bottom: 4px;
            }}

            .page-subtext-secondary {{
                font-weight: 600;
                font-size: 14px;
                opacity: 0.95;
            }}

            .status-message {{
                width: 100%;
                text-align: center;
                color: #542D54;
                font-weight: 700;
                font-size: 14px;
                min-height: 20px;
            }}

            .dropdown {{
                width: 100%;
                padding: 12px 14px;
                border-radius: 8px;
                border: 2px solid #D7DF23;
                background: #FFFFFF;
                color: #542D54;
                font-size: 16px;
                font-weight: 600;
            }}

            @media (max-width: 768px) {{
                body {{
                    padding: 20px;
                }}

                .spec-section {{
                    padding: 18px;
                }}

                .section-label {{
                    font-size: 20px;
                }}
            }}
        </style>
    </head>

    <body>
        <div class="header-container">
            <img src="/assets/Header-PCRBuilder.png" class="header-image" alt="PCR Builder">
        </div>

        <div class="page-subtext">
            <div class="page-subtext-primary">Build your PCR using the PCR Numbers Excel file.</div>
            <div class="page-subtext-secondary">Ready to send straight to the client once downloaded!</div>
        </div>

        <form id="pcrForm" class="spec-section">
            <div class="section-inner">
                <div style="width:100%;">
                    <div class="section-label">Client Name</div>
                    <input
                        type="text"
                        id="clientName"
                        name="client_name"
                        class="client-name-input"
                        placeholder="Client Name"
                        autocomplete="off"
                        required
                    />
                    <div class="field-note">Provide the Client Name above.</div>
                </div>

                <div style="width:100%;">
                    <div class="section-label">Sales Representative</div>
                    <select id="salesRep" name="sales_rep" class="dropdown" required>
                        <option value="" selected disabled>Select Sales Rep</option>
                        {options_html}
                    </select>
                    <div class="field-note">Select the Sales Representative above.</div>
                </div>

                <div style="width:100%;">
                    <div class="section-label">PCR Numbers File</div>
                    <div class="drop-area" id="dropArea">
                        <p>Drag & drop PCR Numbers Excel here!</p>
                        <button type="button" class="gawk-button browse-btn" id="browseBtn">Browse File</button>
                        <input
                            type="file"
                            id="excelFile"
                            class="file-input"
                            accept=".xlsx,.xlsm,.xltx,.xltm"
                            required
                        />
                    </div>
                </div>

                <div class="upload-confirm hidden" id="uploadConfirm">
                    <div class="upload-confirm-text" id="uploadConfirmText">File ready.</div>
                </div>

                <div class="status-message" id="statusMessage"></div>
            </div>
        </form>

        <div class="actions-row">
            <button id="buildBtn" class="gawk-button build-button">Build PCR</button>
        </div>

        <button id="resetBtn" class="gawk-button reset-button">Reset all</button>

        <script>
            const form = document.getElementById("pcrForm");
            const clientNameInput = document.getElementById("clientName");
            const salesRepSelect = document.getElementById("salesRep");
            const excelFileInput = document.getElementById("excelFile");
            const browseBtn = document.getElementById("browseBtn");
            const buildBtn = document.getElementById("buildBtn");
            const resetBtn = document.getElementById("resetBtn");
            const dropArea = document.getElementById("dropArea");
            const uploadConfirm = document.getElementById("uploadConfirm");
            const uploadConfirmText = document.getElementById("uploadConfirmText");
            const statusMessage = document.getElementById("statusMessage");

            browseBtn.addEventListener("click", () => excelFileInput.click());

            excelFileInput.addEventListener("change", () => {{
                updateSelectedFile();
            }});

            ["dragenter", "dragover"].forEach(eventName => {{
                dropArea.addEventListener(eventName, (e) => {{
                    e.preventDefault();
                    e.stopPropagation();
                    dropArea.classList.add("drag-over");
                }});
            }});

            ["dragleave", "drop"].forEach(eventName => {{
                dropArea.addEventListener(eventName, (e) => {{
                    e.preventDefault();
                    e.stopPropagation();
                    dropArea.classList.remove("drag-over");
                }});
            }});

            dropArea.addEventListener("drop", (e) => {{
                const files = e.dataTransfer.files;
                if (!files || !files.length) return;

                const dt = new DataTransfer();
                dt.items.add(files[0]);
                excelFileInput.files = dt.files;
                updateSelectedFile();
            }});

            function updateSelectedFile() {{
                if (excelFileInput.files && excelFileInput.files.length > 0) {{
                    uploadConfirm.classList.remove("hidden");
                    uploadConfirmText.textContent = `File ready: ${{excelFileInput.files[0].name}}`;
                    statusMessage.textContent = "";
                }} else {{
                    uploadConfirm.classList.add("hidden");
                }}
            }}

            buildBtn.addEventListener("click", async () => {{
                const clientName = clientNameInput.value.trim();
                const salesRep = salesRepSelect.value;
                const file = excelFileInput.files[0];

                if (!clientName) {{
                    statusMessage.textContent = "Please enter a client name.";
                    clientNameInput.focus();
                    return;
                }}

                if (!salesRep) {{
                    statusMessage.textContent = "Please select a sales rep.";
                    salesRepSelect.focus();
                    return;
                }}

                if (!file) {{
                    statusMessage.textContent = "Please upload a PCR Numbers Excel file.";
                    return;
                }}

                statusMessage.textContent = "Building PCR...";
                buildBtn.disabled = true;

                const formData = new FormData();
                formData.append("client_name", clientName);
                formData.append("sales_rep", salesRep);
                formData.append("excel_file", file);

                try {{
                    const response = await fetch("/build", {{
                        method: "POST",
                        body: formData
                    }});

                    if (!response.ok) {{
                        let errorText = "Something went wrong building the PCR.";
                        try {{
                            const errorJson = await response.json();
                            errorText = errorJson.detail || errorText;
                        }} catch {{
                            errorText = "Something went wrong building the PCR.";
                        }}
                        throw new Error(errorText);
                    }}

                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const anchor = document.createElement("a");
                    const disposition = response.headers.get("Content-Disposition") || "";
                    const match = disposition.match(/filename="([^"]+)"/);
                    anchor.href = url;
                    anchor.download = match ? match[1] : "PCR_Report.pptx";
                    document.body.appendChild(anchor);
                    anchor.click();
                    anchor.remove();
                    window.URL.revokeObjectURL(url);

                    statusMessage.textContent = "Done. Your PCR has been downloaded.";
                }} catch (error) {{
                    statusMessage.textContent = error.message || "Failed to build PCR.";
                }} finally {{
                    buildBtn.disabled = false;
                }}
            }});

            resetBtn.addEventListener("click", () => {{
                form.reset();
                excelFileInput.value = "";
                uploadConfirm.classList.add("hidden");
                statusMessage.textContent = "";
            }});
        </script>
    </body>
    </html>
    """


@app.post("/build")
async def build_pcr(
    client_name: str = Form(...),
    sales_rep: str = Form(...),
    excel_file: UploadFile = File(...),
):
    client_name = client_name.strip()
    sales_rep = sales_rep.strip()

    if not client_name:
        raise HTTPException(status_code=400, detail="Client name is required.")

    if not sales_rep or sales_rep not in REP_DATA:
        raise HTTPException(status_code=400, detail="Please select a valid sales rep.")

    if not excel_file.filename.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file.")

    try:
        file_bytes = await excel_file.read()
        month_year = extract_month_year_from_excel(file_bytes)
        pptx_stream = build_pcr_pptx(
            client_name=client_name,
            month_year=month_year,
            sales_rep=sales_rep,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    safe_client = clean_filename_part(client_name).replace(" ", "_")
    safe_month = month_year.replace(" ", "_")
    output_filename = f"PCR_Report_{safe_client}_{safe_month}.pptx"

    headers = {
        "Content-Disposition": f'attachment; filename="{output_filename}"'
    }

    return StreamingResponse(
        pptx_stream,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )


if __name__ == "__main__":
    uvicorn.run("pcr_app:app", host="0.0.0.0", port=8000, reload=True)
