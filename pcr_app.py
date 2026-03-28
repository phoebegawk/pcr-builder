from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

from rep_data import REP_DATA, DROPDOWN_REPS
from pcr_helpers import (
    clean_filename_part,
    extract_ado_preview_data,
    extract_ado_contract_fields,
    choose_contract_value,
    extract_month_year_from_excel,
    extract_board_rows,
    read_uploaded_images,
    build_pcr_pptx,
)

app = FastAPI(title="PCR Report Builder")

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")


@app.get("/", response_class=HTMLResponse)
async def home():
    options_html = "".join(
        f'<option value="{rep}"></option>' for rep in DROPDOWN_REPS
    )

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>PCR Builder</title>
        <link rel="icon" type="image/png" href="/assets/favicon.png">
        <link rel="stylesheet" href="/assets/styles.css">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body>
        <div class="header-container">
            <img src="/assets/Header-PCRBuilder.png" class="header-image" alt="PCR Builder">
        </div>

        <div class="page-subtext">
            <div class="page-subtext-primary">Build your PCR using the PCR Numbers Excel file.</div>
            <div class="page-subtext-secondary">Upload an ADO first, or use manual contract information entry.</div>
        </div>

        <form id="pcrForm" class="spec-section">
            <div class="section-inner">

                <div class="actions-row top-toggle-row">
                    <button type="button" id="manualToggleBtn" class="gawk-button secondary-button">Manual Contract Information Entry</button>
                </div>

                <div class="section-block">
                    <div class="section-label">Upload ADO</div>
                    <div class="drop-area" id="adoDropArea">
                        <p>Drag & drop PDF here!</p>
                        <button type="button" class="gawk-button browse-btn" id="adoBrowseBtn">Browse File</button>
                        <input
                            type="file"
                            id="adoFile"
                            class="file-input"
                            accept=".pdf"
                        />
                    </div>

                    <div class="upload-confirm hidden" id="adoUploadConfirm">
                        <div class="upload-confirm-text" id="adoUploadConfirmText">ADO ready.</div>
                    </div>

                    <div class="extracted-preview" id="adoPreview">
                        <div class="extracted-preview-title">Extracted Contract Info</div>
                        <div class="extracted-preview-grid">
                            <div class="extracted-preview-item">
                                <span class="extracted-preview-label">Client Name:</span>
                                <span id="adoPreviewClientName">—</span>
                            </div>
                            <div class="extracted-preview-item">
                                <span class="extracted-preview-label">Sales Representative:</span>
                                <span id="adoPreviewSalesRep">—</span>
                            </div>
                            <div class="extracted-preview-item">
                                <span class="extracted-preview-label">Contract Number:</span>
                                <span id="adoPreviewContractNumber">—</span>
                            </div>
                        </div>
                    </div>
                </div>

                <div id="manualSection" class="manual-section">
                    <div class="section-block">
                        <div class="section-label">Client Name</div>
                        <input
                            type="text"
                            id="clientName"
                            name="client_name"
                            class="client-name-input"
                            placeholder=""
                            autocomplete="off"
                        />
                    </div>

                    <div class="section-block-spaced">
                        <div class="section-label">Sales Representative</div>
                        <input
                            type="text"
                            id="salesRep"
                            name="sales_rep"
                            class="dropdown"
                            list="salesRepOptions"
                            placeholder=""
                            autocomplete="off"
                        />
                        <datalist id="salesRepOptions">
                            {options_html}
                        </datalist>
                    </div>

                    <div class="section-block-spaced">
                        <div class="section-label">Contract Number</div>
                        <input
                            type="text"
                            id="contractNumber"
                            name="contract_number"
                            class="client-name-input"
                            placeholder=""
                            autocomplete="off"
                        />
                    </div>
                </div>

                <div class="section-block-spaced">
                    <div class="section-label">Upload PCR Numbers</div>
                    <div class="drop-area" id="excelDropArea">
                        <p>Drag & drop Excel file here!</p>
                        <button type="button" class="gawk-button browse-btn" id="excelBrowseBtn">Browse File</button>
                        <input
                            type="file"
                            id="excelFile"
                            class="file-input"
                            accept=".xlsx,.xlsm,.xltx,.xltm"
                            required
                        />
                    </div>
                    <div class="upload-confirm hidden" id="excelUploadConfirm">
                        <div class="upload-confirm-text" id="excelUploadConfirmText">File ready.</div>
                    </div>
                </div>

                <div class="section-block-spaced">
                    <div class="section-label">Upload PoP's</div>
                    <div class="drop-area" id="imageDropArea">
                        <p>Drag & drop JPG and PNG here!</p>
                        <button type="button" class="gawk-button browse-btn" id="imageBrowseBtn">Browse Files</button>
                        <input
                            type="file"
                            id="imageFiles"
                            class="file-input"
                            accept=".jpg,.jpeg,.png"
                            multiple
                        />
                    </div>
                    <div class="upload-confirm hidden" id="imageUploadConfirm">
                        <div class="upload-confirm-text" id="imageUploadConfirmText">Images ready.</div>
                    </div>
                </div>

                <div class="status-message" id="statusMessage"></div>
            </div>
        </form>

        <div class="actions-row">
            <button id="buildBtn" class="gawk-button build-button">Build PCR</button>
        </div>

        <button id="resetBtn" class="gawk-button reset-button">Reset all</button>

        <script src="/assets/app.js"></script>
    </body>
    </html>
    """


@app.post("/extract-ado")
async def extract_ado(
    ado_file: UploadFile = File(...),
):
    if not ado_file.filename or not ado_file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a valid ADO PDF.")

    try:
        pdf_bytes = await ado_file.read()
        preview_data = extract_ado_preview_data(pdf_bytes)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return preview_data


@app.post("/build")
async def build_pcr(
    client_name: str = Form(""),
    sales_rep: str = Form(""),
    contract_number: str = Form(""),
    excel_file: UploadFile = File(...),
    ado_file: UploadFile | None = File(default=None),
    board_images: list[UploadFile] = File(default=[]),
):
    client_name = client_name.strip()
    sales_rep = sales_rep.strip()
    contract_number = contract_number.strip()

    if ado_file and ado_file.filename and not ado_file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a valid ADO PDF.")

    if not excel_file.filename.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file.")

    try:
        ado_values = {
            "client_name": "",
            "sales_rep": "",
            "contract_number": "",
        }

        if ado_file and ado_file.filename:
            ado_bytes = await ado_file.read()
            ado_values = extract_ado_contract_fields(ado_bytes)

        final_client_name = choose_contract_value(client_name, ado_values["client_name"], "Client Name")
        final_sales_rep = choose_contract_value(sales_rep, ado_values["sales_rep"], "Sales Representative")
        final_contract_number = choose_contract_value(contract_number, ado_values["contract_number"], "Contract Number")

        file_bytes = await excel_file.read()
        month_year = extract_month_year_from_excel(file_bytes)
        board_rows = extract_board_rows(file_bytes)
        uploaded_images = await read_uploaded_images(board_images)

        pptx_stream = build_pcr_pptx(
            client_name=final_client_name,
            month_year=month_year,
            contract_number=final_contract_number,
            sales_rep=final_sales_rep,
            board_rows=board_rows,
            uploaded_images=uploaded_images,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    safe_client = clean_filename_part(final_client_name).replace(" ", "_")
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
