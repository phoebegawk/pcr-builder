const form = document.getElementById("pcrForm");
const manualSection = document.getElementById("manualSection");
const manualToggleBtn = document.getElementById("manualToggleBtn");

const adoFileInput = document.getElementById("adoFile");
const adoBrowseBtn = document.getElementById("adoBrowseBtn");
const adoDropArea = document.getElementById("adoDropArea");
const adoUploadConfirm = document.getElementById("adoUploadConfirm");
const adoUploadConfirmText = document.getElementById("adoUploadConfirmText");

const adoPreview = document.getElementById("adoPreview");
const adoPreviewClientName = document.getElementById("adoPreviewClientName");
const adoPreviewSalesRep = document.getElementById("adoPreviewSalesRep");
const adoPreviewContractNumber = document.getElementById("adoPreviewContractNumber");

const clientNameInput = document.getElementById("clientName");
const salesRepSelect = document.getElementById("salesRep");
const contractNumberInput = document.getElementById("contractNumber");

const excelFileInput = document.getElementById("excelFile");
const excelBrowseBtn = document.getElementById("excelBrowseBtn");
const excelDropArea = document.getElementById("excelDropArea");
const excelUploadConfirm = document.getElementById("excelUploadConfirm");
const excelUploadConfirmText = document.getElementById("excelUploadConfirmText");

const imageFilesInput = document.getElementById("imageFiles");
const imageBrowseBtn = document.getElementById("imageBrowseBtn");
const imageDropArea = document.getElementById("imageDropArea");
const imageUploadConfirm = document.getElementById("imageUploadConfirm");
const imageUploadConfirmText = document.getElementById("imageUploadConfirmText");

const buildBtn = document.getElementById("buildBtn");
const resetBtn = document.getElementById("resetBtn");
const statusMessage = document.getElementById("statusMessage");

manualToggleBtn.addEventListener("click", () => {
    manualSection.classList.toggle("visible");
});

adoBrowseBtn.addEventListener("click", () => adoFileInput.click());
excelBrowseBtn.addEventListener("click", () => excelFileInput.click());
imageBrowseBtn.addEventListener("click", () => imageFilesInput.click());

adoFileInput.addEventListener("change", () => {
    updateSelectedAdoFile();
});

excelFileInput.addEventListener("change", () => {
    updateSelectedExcelFile();
});

imageFilesInput.addEventListener("change", () => {
    updateSelectedImageFiles();
});

["dragenter", "dragover"].forEach((eventName) => {
    [adoDropArea, excelDropArea, imageDropArea].forEach((area) => {
        area.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            area.classList.add("drag-over");
        });
    });
});

["dragleave", "drop"].forEach((eventName) => {
    [adoDropArea, excelDropArea, imageDropArea].forEach((area) => {
        area.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            area.classList.remove("drag-over");
        });
    });
});

adoDropArea.addEventListener("drop", (e) => {
    const files = e.dataTransfer.files;
    if (!files || !files.length) return;

    const dt = new DataTransfer();
    dt.items.add(files[0]);
    adoFileInput.files = dt.files;
    updateSelectedAdoFile();
});

excelDropArea.addEventListener("drop", (e) => {
    const files = e.dataTransfer.files;
    if (!files || !files.length) return;

    const dt = new DataTransfer();
    dt.items.add(files[0]);
    excelFileInput.files = dt.files;
    updateSelectedExcelFile();
});

imageDropArea.addEventListener("drop", (e) => {
    const files = e.dataTransfer.files;
    if (!files || !files.length) return;

    const dt = new DataTransfer();
    for (let i = 0; i < files.length; i++) {
        dt.items.add(files[i]);
    }
    imageFilesInput.files = dt.files;
    updateSelectedImageFiles();
});

async function updateSelectedAdoFile() {
    if (adoFileInput.files && adoFileInput.files.length > 0) {
        adoUploadConfirm.classList.remove("hidden");
        adoUploadConfirmText.textContent = `ADO uploaded ✅ ${adoFileInput.files[0].name}`;
        statusMessage.textContent = "";

        const formData = new FormData();
        formData.append("ado_file", adoFileInput.files[0]);

        try {
            const response = await fetch(`${window.location.origin}/extract-ado`, {
                method: "POST",
                body: formData
            });

            if (!response.ok) {
                let errorText = "Couldn't extract fields from the ADO.";
                try {
                    const errorJson = await response.json();
                    errorText = errorJson.detail || errorText;
                } catch {
                    errorText = "Couldn't extract fields from the ADO.";
                }
                throw new Error(errorText);
            }

            const data = await response.json();

            adoPreviewClientName.textContent = data.client_name || "—";
            adoPreviewSalesRep.textContent = data.sales_rep || "—";
            adoPreviewContractNumber.textContent = data.contract_number || "—";
            adoPreview.classList.add("visible");

            if (data.client_name && !clientNameInput.value.trim()) {
                clientNameInput.value = data.client_name;
            }
            if (data.sales_rep && !salesRepSelect.value.trim()) {
                salesRepSelect.value = data.sales_rep;
            }
            if (data.contract_number && !contractNumberInput.value.trim()) {
                contractNumberInput.value = data.contract_number;
            }
        } catch (error) {
            adoPreview.classList.remove("visible");
            adoPreviewClientName.textContent = "—";
            adoPreviewSalesRep.textContent = "—";
            adoPreviewContractNumber.textContent = "—";
            statusMessage.textContent = error.message || "Failed to read ADO.";
        }
    } else {
        adoUploadConfirm.classList.add("hidden");
        adoPreview.classList.remove("visible");
        adoPreviewClientName.textContent = "—";
        adoPreviewSalesRep.textContent = "—";
        adoPreviewContractNumber.textContent = "—";
    }
}

function updateSelectedExcelFile() {
    if (excelFileInput.files && excelFileInput.files.length > 0) {
        excelUploadConfirm.classList.remove("hidden");
        excelUploadConfirmText.textContent = `PCR Numbers uploaded ✅ ${excelFileInput.files[0].name}`;
        statusMessage.textContent = "";
    } else {
        excelUploadConfirm.classList.add("hidden");
    }
}

function updateSelectedImageFiles() {
    if (imageFilesInput.files && imageFilesInput.files.length > 0) {
        imageUploadConfirm.classList.remove("hidden");
        imageUploadConfirmText.textContent = `${imageFilesInput.files.length} PoP image(s) uploaded ✅`;
        statusMessage.textContent = "";
    } else {
        imageUploadConfirm.classList.add("hidden");
    }
}

function resizeImageFile(file, maxDimension = 2200, quality = 0.82) {
    return new Promise((resolve, reject) => {
        if (!file.type.startsWith("image/")) {
            resolve(file);
            return;
        }

        const img = new Image();
        const objectUrl = URL.createObjectURL(file);

        img.onload = () => {
            const width = img.width;
            const height = img.height;

            let targetWidth = width;
            let targetHeight = height;

            if (width > height && width > maxDimension) {
                targetWidth = maxDimension;
                targetHeight = Math.round((height * maxDimension) / width);
            } else if (height >= width && height > maxDimension) {
                targetHeight = maxDimension;
                targetWidth = Math.round((width * maxDimension) / height);
            }

            const canvas = document.createElement("canvas");
            canvas.width = targetWidth;
            canvas.height = targetHeight;

            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0, targetWidth, targetHeight);

            canvas.toBlob(
                (blob) => {
                    URL.revokeObjectURL(objectUrl);

                    if (!blob) {
                        reject(new Error(`Couldn't process image: ${file.name}`));
                        return;
                    }

                    const newName = file.name.replace(/\.(png|jpg|jpeg)$/i, ".jpg");
                    const resizedFile = new File([blob], newName, {
                        type: "image/jpeg",
                        lastModified: Date.now()
                    });

                    resolve(resizedFile);
                },
                "image/jpeg",
                quality
            );
        };

        img.onerror = () => {
            URL.revokeObjectURL(objectUrl);
            reject(new Error(`Couldn't read image: ${file.name}`));
        };

        img.src = objectUrl;
    });
}

async function prepareBoardImages(files) {
    const prepared = [];
    for (const file of files) {
        const resized = await resizeImageFile(file, 2200, 0.82);
        prepared.push(resized);
    }
    return prepared;
}

buildBtn.addEventListener("click", async () => {
    const clientName = clientNameInput ? clientNameInput.value.trim() : "";
    const salesRep = salesRepSelect ? salesRepSelect.value.trim() : "";
    const contractNumber = contractNumberInput ? contractNumberInput.value.trim() : "";
    const excelFile = excelFileInput.files[0];
    const adoFile = adoFileInput.files[0];

    if (!excelFile) {
        statusMessage.textContent = "Please upload a PCR Numbers Excel file.";
        return;
    }

    if (!adoFile && !clientName && !salesRep && !contractNumber) {
        statusMessage.textContent = "Upload an ADO PDF or use manual contract information entry.";
        return;
    }

    statusMessage.textContent = "Building PCR...";
    buildBtn.disabled = true;

    try {
        const formData = new FormData();
        formData.append("client_name", clientName);
        formData.append("sales_rep", salesRep);
        formData.append("contract_number", contractNumber);
        formData.append("excel_file", excelFile);

        if (adoFile) {
            formData.append("ado_file", adoFile);
        }

        const preparedImages = await prepareBoardImages(Array.from(imageFilesInput.files));

        for (const imageFile of preparedImages) {
            formData.append("board_images", imageFile);
        }

        const response = await fetch(`${window.location.origin}/build`, {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            let errorText = "Something went wrong building the PCR.";
            try {
                const errorJson = await response.json();
                errorText = errorJson.detail || errorText;
            } catch {
                errorText = "Something went wrong building the PCR.";
            }
            throw new Error(errorText);
        }

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

        statusMessage.textContent = "Done. Your PCR PPTX has been downloaded.";
    } catch (error) {
        statusMessage.textContent = error.message || "Failed to build PCR.";
    } finally {
        buildBtn.disabled = false;
    }
});

resetBtn.addEventListener("click", () => {
    form.reset();

    adoFileInput.value = "";
    excelFileInput.value = "";
    imageFilesInput.value = "";

    adoUploadConfirm.classList.add("hidden");
    excelUploadConfirm.classList.add("hidden");
    imageUploadConfirm.classList.add("hidden");

    adoPreview.classList.remove("visible");
    adoPreviewClientName.textContent = "—";
    adoPreviewSalesRep.textContent = "—";
    adoPreviewContractNumber.textContent = "—";

    statusMessage.textContent = "";
});
