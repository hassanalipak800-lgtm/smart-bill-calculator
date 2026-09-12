/* ---------- Theme toggle ---------- */
const themeToggle = document.getElementById('themeToggle');
themeToggle.addEventListener('click', () => {
  document.body.classList.toggle('dark');
  themeToggle.textContent = document.body.classList.contains('dark') ? '☀️' : '🌙';
});

/* ---------- Tabs ---------- */
const tabBtns = document.querySelectorAll('.tab-btn');
const panels = document.querySelectorAll('.tab-panel');
tabBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    tabBtns.forEach(b => { b.classList.remove('active'); b.setAttribute('aria-selected','false'); });
    panels.forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    btn.setAttribute('aria-selected','true');
    document.getElementById('panel-' + btn.dataset.tab).classList.add('active');
  });
});

/* ---------- Helper: extract numbers from raw text ---------- */
function extractNumbers(rawText, minAmount = 0) {
  // Matches numbers like 1,200.50 or 1200 or 150
  const matches = rawText.match(/\d{1,3}(?:[,.\s]\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?/g) || [];
  return matches
    .map(m => parseFloat(m.replace(/,/g, '')))
    .filter(n => !isNaN(n) && n > 0 && n >= minAmount);
}

function formatNum(n) {
  return n.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

/* ---------- Hero receipt animated total (decorative) ---------- */
(function animateHeroTotal(){
  const el = document.getElementById('heroTotal');
  const target = 30200;
  let cur = 0;
  const step = Math.ceil(target / 40);
  const iv = setInterval(() => {
    cur += step;
    if (cur >= target) { cur = target; clearInterval(iv); }
    el.textContent = formatNum(cur);
  }, 30);
})();

/* ======================================================
   TOOL 1: IMAGE OCR
====================================================== */
const imageInput = document.getElementById('imageInput');
const imageDropzone = document.getElementById('imageDropzone');
const imagePreview = document.getElementById('imagePreview');
const processImageBtn = document.getElementById('processImageBtn');
const imageStatus = document.getElementById('imageStatus');
const imageResult = document.getElementById('imageResult');
const imageFoundNumbers = document.getElementById('imageFoundNumbers');
const imageTotalValue = document.getElementById('imageTotalValue');

let selectedImageFile = null;

['dragover','dragenter'].forEach(evt =>
  imageDropzone.addEventListener(evt, e => { e.preventDefault(); imageDropzone.classList.add('drag'); })
);
['dragleave','drop'].forEach(evt =>
  imageDropzone.addEventListener(evt, e => { e.preventDefault(); imageDropzone.classList.remove('drag'); })
);
imageDropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file) handleImageFile(file);
});
imageInput.addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) handleImageFile(file);
});

function handleImageFile(file) {
  selectedImageFile = file;
  const url = URL.createObjectURL(file);
  imagePreview.src = url;
  imagePreview.hidden = false;
  processImageBtn.disabled = false;
  imageResult.hidden = true;
}

processImageBtn.addEventListener('click', async () => {
  if (!selectedImageFile) return;
  imageStatus.hidden = false;
  imageStatus.textContent = 'Tasveer padhi ja rahi hai... thoda intezaar karein.';
  processImageBtn.disabled = true;
  imageResult.hidden = true;

  try {
    const result = await Tesseract.recognize(selectedImageFile, 'eng', {
      logger: m => {
        if (m.status === 'recognizing text') {
          imageStatus.textContent = `Padha ja raha hai... ${Math.round(m.progress * 100)}%`;
        }
      }
    });
    const minAmount = parseFloat(document.getElementById('minAmountInput').value) || 0;
    const numbers = extractNumbers(result.data.text, minAmount);
    showImageResults(numbers);
    imageStatus.hidden = true;
  } catch (err) {
    imageStatus.textContent = 'Kuch masla ho gaya, dobara koshish karein. (' + err.message + ')';
  } finally {
    processImageBtn.disabled = false;
  }
});

function showImageResults(numbers) {
  imageFoundNumbers.innerHTML = '';
  if (numbers.length === 0) {
    imageFoundNumbers.innerHTML = '<p class="note">Koi number nahi mila. Saaf tasveer try karein.</p>';
    imageResult.hidden = false;
    imageTotalValue.textContent = '0';
    return;
  }
  numbers.forEach((n, i) => {
    const label = document.createElement('label');
    label.innerHTML = `<input type="checkbox" checked data-value="${n}"> ${formatNum(n)}`;
    imageFoundNumbers.appendChild(label);
  });
  imageResult.hidden = false;
  recalcImageTotal();
  imageFoundNumbers.querySelectorAll('input').forEach(cb =>
    cb.addEventListener('change', recalcImageTotal)
  );
}

function recalcImageTotal() {
  let total = 0;
  imageFoundNumbers.querySelectorAll('input:checked').forEach(cb => {
    total += parseFloat(cb.dataset.value);
  });
  imageTotalValue.textContent = formatNum(total);
}

document.getElementById('imagePrintBtn').addEventListener('click', () => window.print());

/* ======================================================
   TOOL 2: EXCEL / CSV
====================================================== */
const sheetInput = document.getElementById('sheetInput');
const sheetDropzone = document.getElementById('sheetDropzone');
const sheetStatus = document.getElementById('sheetStatus');
const sheetResult = document.getElementById('sheetResult');
const sheetTable = document.getElementById('sheetTable');
const sheetColumnSelect = document.getElementById('sheetColumnSelect');
const sheetTotalValue = document.getElementById('sheetTotalValue');

let sheetRows = [];

['dragover','dragenter'].forEach(evt =>
  sheetDropzone.addEventListener(evt, e => { e.preventDefault(); sheetDropzone.classList.add('drag'); })
);
['dragleave','drop'].forEach(evt =>
  sheetDropzone.addEventListener(evt, e => { e.preventDefault(); sheetDropzone.classList.remove('drag'); })
);
sheetDropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file) handleSheetFile(file);
});
sheetInput.addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) handleSheetFile(file);
});

function handleSheetFile(file) {
  sheetStatus.hidden = false;
  sheetStatus.textContent = 'File padhi ja rahi hai...';
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      sheetRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      renderSheetPreview(sheetRows);
      sheetStatus.hidden = true;
    } catch (err) {
      sheetStatus.textContent = 'File nahi padhi ja saki. Sahi Excel/CSV file chunein.';
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderSheetPreview(rows) {
  if (!rows.length) return;
  const header = rows[0];
  const bodyRows = rows.slice(1, 11); // preview first 10 rows

  let html = '<thead><tr>';
  header.forEach(h => html += `<th>${h || ''}</th>`);
  html += '</tr></thead><tbody>';
  bodyRows.forEach(r => {
    html += '<tr>' + header.map((_, i) => `<td>${r[i] ?? ''}</td>`).join('') + '</tr>';
  });
  html += '</tbody>';
  sheetTable.innerHTML = html;

  sheetColumnSelect.innerHTML = '';
  header.forEach((h, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = h || `Column ${i + 1}`;
    sheetColumnSelect.appendChild(opt);
  });

  // Auto-pick the first column that looks numeric
  let bestCol = 0, bestCount = -1;
  header.forEach((_, colIdx) => {
    const count = rows.slice(1).filter(r => !isNaN(parseFloat(r[colIdx]))).length;
    if (count > bestCount) { bestCount = count; bestCol = colIdx; }
  });
  sheetColumnSelect.value = bestCol;

  recalcSheetTotal();
  sheetColumnSelect.addEventListener('change', recalcSheetTotal);
  sheetResult.hidden = false;
}

function recalcSheetTotal() {
  const colIdx = parseInt(sheetColumnSelect.value);
  let total = 0;
  sheetRows.slice(1).forEach(r => {
    const val = parseFloat(String(r[colIdx]).replace(/,/g, ''));
    if (!isNaN(val)) total += val;
  });
  sheetTotalValue.textContent = formatNum(total);
}

document.getElementById('sheetPrintBtn').addEventListener('click', () => window.print());

/* ======================================================
   TOOL 3: MANUAL TEXT
====================================================== */
const textInput = document.getElementById('textInput');
const processTextBtn = document.getElementById('processTextBtn');
const textResult = document.getElementById('textResult');
const textTotalValue = document.getElementById('textTotalValue');
const textAvgValue = document.getElementById('textAvgValue');
const textCountValue = document.getElementById('textCountValue');

processTextBtn.addEventListener('click', () => {
  const numbers = extractNumbers(textInput.value);
  const total = numbers.reduce((a, b) => a + b, 0);
  const avg = numbers.length ? total / numbers.length : 0;

  textTotalValue.textContent = formatNum(total);
  textAvgValue.textContent = formatNum(avg);
  textCountValue.textContent = numbers.length;
  textResult.hidden = false;
});

document.getElementById('textPrintBtn').addEventListener('click', () => window.print());
