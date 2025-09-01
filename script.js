pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

let files = [], currentIndex = 0, extractedData = [];

const pdfCanvas = document.getElementById("pdfCanvas");
const overlayCanvas = document.getElementById("overlayCanvas");
const ctx = pdfCanvas.getContext("2d");
const overlayCtx = overlayCanvas.getContext("2d");
const textOutput = document.getElementById("textOutput");
const tableBody = document.querySelector("#dataTable tbody");

const fileInput = document.getElementById("fileInput");
const prevBtn = document.getElementById("prevBtn");
const nextBtn = document.getElementById("nextBtn");
const downloadExcelBtn = document.getElementById("downloadExcel");
const cropBtn = document.getElementById("cropBtn");

// Overlay sizing sync
function resizeOverlay() {
  overlayCanvas.width = pdfCanvas.width;
  overlayCanvas.height = pdfCanvas.height;
  overlayCanvas.style.top = pdfCanvas.offsetTop + "px";
  overlayCanvas.style.left = pdfCanvas.offsetLeft + "px";
}
window.addEventListener('resize', () => { 
  resizeOverlay();
});
resizeOverlay();

function enableNavButtons() {
  prevBtn.disabled = currentIndex <= 0;
  nextBtn.disabled = currentIndex >= files.length - 1;
}

fileInput.addEventListener("change", async e => {
  files = Array.from(e.target.files);
  currentIndex = 0;
  extractedData = [];
  tableBody.innerHTML = "";
  if (files.length > 0) {
    prevBtn.disabled = false;
    nextBtn.disabled = false;
    await showPDF(currentIndex);
    enableNavButtons();
  } else {
    textOutput.textContent = "Upload a PDF to start...";
    prevBtn.disabled = true;
    nextBtn.disabled = true;
  }
});

prevBtn.addEventListener("click", async () => {
  saveTableEdits();
  if (currentIndex > 0) {
    currentIndex--;
    await showPDF(currentIndex);
    enableNavButtons();
  }
});

nextBtn.addEventListener("click", async () => {
  saveTableEdits();
  if (currentIndex < files.length - 1) {
    currentIndex++;
    await showPDF(currentIndex);
    enableNavButtons();
  }
});

downloadExcelBtn.addEventListener("click", () => {
  saveTableEdits();
  if (extractedData.length === 0 || extractedData.every(row => !row)) {
    alert("No data to export");
    return;
  }
  const filteredData = extractedData.filter(r => r);
  const ws = XLSX.utils.json_to_sheet(filteredData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Internships");
  XLSX.writeFile(wb, "Internship_Data.xlsx");
});

async function showPDF(index) {
  textOutput.textContent = "Loading PDF...";
  overlayCtx.clearRect(0, 0, overlayCanvas.width, overlayCanvas.height);
  const file = files[index];
  const url = URL.createObjectURL(file);
  try {
    const pdf = await pdfjsLib.getDocument(url).promise;
    const page = await pdf.getPage(1);
    const scale = 2.5;
    const viewport = page.getViewport({ scale });
    pdfCanvas.width = viewport.width;
    pdfCanvas.height = viewport.height;
    resizeOverlay();
    await page.render({ canvasContext: ctx, viewport }).promise;
    const textContent = await page.getTextContent();
    let extractedText = textContent.items.map(item => item.str).join(" ").trim();
    if (!extractedText) {
      textOutput.textContent = "Running OCR...";
      extractedText = await runImprovedOCR(pdfCanvas);
      textOutput.textContent = extractedText || "No text found (OCR failed)";
    } else {
      textOutput.textContent = extractedText;
    }
    saveExtracted(extractedText, index);
  } catch (err) {
    textOutput.textContent = "Error: " + err.message;
    extractedData[index] = { Name: "Error", Company: "Error", Duration: "Error" };
    renderTable();
  }
}

async function runImprovedOCR(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(async blob => {
      try {
        const result = await Tesseract.recognize(blob, 'eng', {
          tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789&.- ',
          tessedit_pageseg_mode: Tesseract.PSM.AUTO
        });
        resolve(result.data.text);
      } catch (e) {
        reject(e);
      }
    });
  });
}

// Crop OCR with nice overlay drawing and stable user experience
cropBtn.addEventListener("click", () => {
  let cropping = false, startX, startY, endX, endY;
  overlayCanvas.style.cursor = "crosshair";
  function draw() {
    overlayCtx.clearRect(0, 0, overlayCanvas.width, overlayCanvas.height);
    if (cropping) {
      const x = Math.min(startX, endX);
      const y = Math.min(startY, endY);
      const w = Math.abs(endX - startX);
      const h = Math.abs(endY - startY);
      overlayCtx.strokeStyle = "red";
      overlayCtx.lineWidth = 2;
      overlayCtx.strokeRect(x, y, w, h);
    }
  }
  function mouseDown(e) {
    cropping = true;
    startX = e.offsetX;
    startY = e.offsetY;
    endX = startX;
    endY = startY;
    draw();
    overlayCanvas.addEventListener("mousemove", mouseMove);
    overlayCanvas.addEventListener("mouseup", mouseUp);
  }
  function mouseMove(e) {
    endX = e.offsetX;
    endY = e.offsetY;
    draw();
  }
  async function mouseUp() {
    cropping = false;
    overlayCanvas.style.cursor = "default";
    overlayCtx.clearRect(0, 0, overlayCanvas.width, overlayCanvas.height);
    overlayCanvas.removeEventListener("mousemove", mouseMove);
    overlayCanvas.removeEventListener("mouseup", mouseUp);
    const x = Math.min(startX, endX);
    const y = Math.min(startY, endY);
    const w = Math.abs(endX - startX);
    const h = Math.abs(endY - startY);
    if (w > 10 && h > 10) {
      const tempCanvas = document.createElement("canvas");
      tempCanvas.width = w;
      tempCanvas.height = h;
      const tempCtx = tempCanvas.getContext("2d");
      tempCtx.drawImage(pdfCanvas, x, y, w, h, 0, 0, w, h);
      textOutput.textContent = "Running OCR on cropped area...";
      try {
        const croppedText = await runImprovedOCR(tempCanvas);
        textOutput.textContent = croppedText || "No text found in cropped area.";
        saveExtracted(croppedText, currentIndex);
      } catch (err) {
        textOutput.textContent = "OCR failed: " + err.message;
      }
    }
  }
  overlayCanvas.addEventListener("mousedown", mouseDown, { once: true });
});

function extractFields(text) {
  const cleaned = text.replace(/\r?\n|\r/g, ' ').replace(/\s+/g, ' ').trim();
  let name = "Not Found";
  let company = "Not Found";
  let duration = "Not Found";
  let nameMatch = cleaned.match(/Dear[, ]+\s*([A-Z][A-Z\s]*[A-Z])/i) ||
                  cleaned.match(/Dear[, ]+([A-Z][a-z]+(?:\s[A-Z][a-z]+){0,3})/i) ||
                  cleaned.match(/Dear\s*,?\s*([A-Z][^\s,]*)/i) ||
                  cleaned.match(/Dear\s*[,:]?\s*([A-Z][A-Za-z\s\.]{1,50})/i);
  if (nameMatch) name = nameMatch[1].trim();
  else {
    const fallbackName = cleaned.match(/([A-Z][a-z]+(?:\s[A-Z][a-z]+){1,3})/);
    if (fallbackName) name = fallbackName[1].trim();
  }
  let companyMatch = cleaned.match(/with\s+“?([A-Z][A-Za-z&\.\s]{1,50})”?/i) ||
                     cleaned.match(/at\s+“?([A-Z][A-Za-z&\.\s]{1,50})”?/i) ||
                     cleaned.match(/Founder\s*\(([^)]+)\)/i) ||
                     cleaned.match(/(?:with|at|from)\s+([A-Z][A-Za-z&\.\s\-]{3,50}(Foundation|Ltd|Pvt|Inc|Corp|Studio|Solutions|Technology|Technologies|Systems))/) ||
                     cleaned.match(/([A-Z][A-Za-z\s&\.\-]{3,50}(Foundation|Ltd|Pvt|Inc|Corp|Studio|Solutions|Technology|Technologies|Systems))/);
  if (companyMatch) {
    company = typeof companyMatch === 'string' ? companyMatch.trim() : companyMatch[1] ? companyMatch[1].trim() : companyMatch[0].trim();
  }
  let durationMatch = cleaned.match(/duration(?: of| will be of)?\s*(?:the)?\s*(\d+\s*(?:weeks?|months?))/i) ||
                      cleaned.match(/duration(?: of| will be of)?\s*(?:the)?\s*(one|two|three|four|five|six|seven|eight|nine|ten)\s*(weeks?|months?)/i);
  if (durationMatch) {
    duration = durationMatch[1].trim();
  } else {
    let genericDur = cleaned.match(/(\d+\s*(?:weeks?|months?))/i);
    if (genericDur) duration = genericDur[1].trim();
  }
  return { name, company, duration };
}

function saveExtracted(text, index = currentIndex) {
  const { name, company, duration } = extractFields(text);
  extractedData[index] = { Name: name, Company: company, Duration: duration };
  renderTable();
}

function renderTable() {
  tableBody.innerHTML = "";
  extractedData.forEach((row, i) => {
    const tr = document.createElement("tr");
    tr.dataset.index = i;
    tr.innerHTML = `
      <td contenteditable="true" aria-label="Name">${row?.Name || ""}</td>
      <td contenteditable="true" aria-label="Company">${row?.Company || ""}</td>
      <td contenteditable="true" aria-label="Duration">${row?.Duration || ""}</td>`;
    tableBody.appendChild(tr);
  });
}

function saveTableEdits() {
  const rows = tableBody.querySelectorAll("tr");
  rows.forEach(tr => {
    const idx = tr.dataset.index;
    const cells = tr.querySelectorAll("td");
    extractedData[idx] = {
      Name: cells[0].textContent.trim(),
      Company: cells[1].textContent.trim(),
      Duration: cells[2].textContent.trim()
    };
  });
}
