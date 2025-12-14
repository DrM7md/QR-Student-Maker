/* ========= عناصر الصفحة ========= */
const fileInput = document.getElementById('excelFile');
const btnGenerate = document.getElementById('btnGenerate');
const btnClear = document.getElementById('btnClear');

const qrSizeEl = document.getElementById('qrSize');
const qrMarginEl = document.getElementById('qrMargin');
const qrECEl = document.getElementById('qrEC');

const grid = document.getElementById('grid');
const statusBox = document.getElementById('status');
const countInfo = document.getElementById('countInfo');

let lastGenerated = []; // {id, name, cls, folder, dataUrl}

/* ========= أدوات مساعدة ========= */
const showStatus = (msg, type = 'info') => {
  statusBox.classList.remove('hidden');
  statusBox.className = 'mb-4 text-sm rounded-lg p-3';

  if (type === 'error') statusBox.classList.add('bg-red-50', 'text-red-700', 'border', 'border-red-200');
  else if (type === 'success') statusBox.classList.add('bg-emerald-50', 'text-emerald-700', 'border', 'border-emerald-200');
  else statusBox.classList.add('bg-slate-50', 'text-slate-700', 'border', 'border-slate-200');

  statusBox.textContent = msg;
};

const clearUI = () => {
  grid.innerHTML = '';
  statusBox.classList.add('hidden');
  countInfo.textContent = '0 عنصر';
  lastGenerated = [];
};

/* ========= حماية نص للعرض ========= */
function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (m) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  }[m]));
}

/* ========= أسماء ملفات آمنة ========= */
function safeFileName(str) {
  return String(str)
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
}

/* 07/1 -> 07-1 */
function folderFromClass(cls) {
  return safeFileName(String(cls).trim().replace(/\//g, '-'));
}

/* ========= قراءة Excel =========
   الأعمدة:
   A = الرقم الشخصي
   B = اسم الطالب
   C = الصف/الشعبة مثال 07/1
   من A2
*/
async function readExcelRowsABC(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });

  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });

  const items = [];
  for (let i = 1; i < rows.length; i++) { // من الصف 2
    const nationalId = rows[i]?.[0]; // A
    const studentName = rows[i]?.[1]; // B
    const classSection = rows[i]?.[2]; // C

    if (nationalId == null || studentName == null || classSection == null) continue;

    const id = String(nationalId).trim();
    const name = String(studentName).trim();
    const cls = String(classSection).trim();

    if (!id || !name || !cls) continue;

    items.push({ id, name, cls });
  }

  // إزالة التكرار (نفس الطالب بنفس الشعبة)
  const seen = new Set();
  const unique = [];
  for (const it of items) {
    const key = `${it.id}__${it.cls}`;
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(it);
  }

  return unique;
}

/* ========= توليد QR (QRious) كـ DataURL ========= */
async function makeQrDataUrl(text, size, margin, ecLevel) {
  // QRious ما يدعم ECLevel بشكل مباشر - نخليه موجود للواجهة فقط
  const m = Number(margin) || 0;

  const out = document.createElement('canvas');
  out.width = size;
  out.height = size;
  const ctx = out.getContext('2d');

  // خلفية بيضاء
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, size, size);

  // QR داخلي
  const innerSize = Math.max(64, size - (m * 2 * 10)); // تقريب للهامش
  const qrCanvas = document.createElement('canvas');

  new QRious({
    element: qrCanvas,
    value: String(text),
    size: innerSize,
  });

  const x = (size - innerSize) / 2;
  const y = (size - innerSize) / 2;
  ctx.drawImage(qrCanvas, x, y, innerSize, innerSize);

  return out.toDataURL('image/png');
}

/* ========= بطاقة معاينة ========= */
function addCard(title, dataUrl) {
  const wrap = document.createElement('div');
  wrap.className = 'card';

  wrap.innerHTML = `
    <div class="flex items-start justify-between gap-2 mb-2">
      <div class="text-sm font-semibold break-all">${escapeHtml(title)}</div>
      <span class="badge">QR</span>
    </div>

    <div class="bg-slate-50 rounded-lg p-2 flex items-center justify-center">
      <img src="${dataUrl}" alt="QR" class="max-w-full h-auto" />
    </div>
  `;

  grid.appendChild(wrap);
}

/* ========= حفظ PNG داخل ZIP (مجلد لكل شعبة) ========= */
async function saveAllAsPngZip(items) {
  const zip = new JSZip();

  for (const it of items) {
    const folderName = it.folder || 'بدون_شعبة';
    const folder = zip.folder(folderName);

    // اسم الملف = اسم الطالب - الرقم الشخصي (تفادي التكرار)
    const fileName = `${safeFileName(it.name)} - ${safeFileName(it.id)}.png`;

    const base64 = it.dataUrl.split(',')[1];
    folder.file(fileName, base64, { base64: true });
  }

  const blob = await zip.generateAsync({ type: 'blob' });
  saveAs(blob, 'QR_By_Class.zip');
}

/* ========= حفظ PDF: صفحة واحدة لكل شعبة (Auto Columns + Auto Fit) ========= */
async function savePdfOnePagePerClass(items) {
  const { jsPDF } = window.jspdf;

  // ✅ حماية: إذا الخط غير موجود لا نكمل
  if (!window.ARABIC_FONT_BASE64 || typeof window.ARABIC_FONT_BASE64 !== 'string') {
    throw new Error('الخط العربي غير موجود. تأكد من fonts.js وربطه قبل app.js');
  }

  // ✅ نضيف الخط مرة واحدة فقط
  if (!jsPDF.__arabicFontAdded) {
    jsPDF.API.events.push([
      'addFonts',
      function () {
        this.addFileToVFS('Amiri-Regular.ttf', window.ARABIC_FONT_BASE64);
        this.addFont('Amiri-Regular.ttf', 'Amiri', 'normal');
      }
    ]);
    jsPDF.__arabicFontAdded = true;
  }

  const doc = new jsPDF({ unit: 'mm', format: 'a4' });

  const pageW = 210;
  const pageH = 297;

  // إعدادات عامة
  const margin = 8;
  const gapX = 3;
  const gapY = 4;

  const headerH = 12;
  const startX = margin;
  const startY = margin + headerH;

  const availW = pageW - margin * 2;
  const availH = pageH - startY - margin;

  // الهدف
  const targetMax = 33;

  // حدود الأعمدة (Auto)
  const minCols = 3;
  const maxCols = 6;

  // حدود حجم QR في PDF
  const minQr = 22;
  const maxQr = 70;

  // نسبة تقريبية لمساحة النص تحت QR (اسم + رقم)
  const ratio = 1.28; // cellH ≈ qrBox * 1.28

  // ✅ نحاول نختار أفضل عدد أعمدة + أكبر QR ممكن
 // ✅ نحاول نختار أفضل عدد أعمدة + أكبر QR ممكن (بحسبة حقيقية: QR + نص)
let best = null;

for (let colsTry = maxCols; colsTry >= minCols; colsTry--) {
  const neededRows = Math.ceil(targetMax / colsTry);

  // أقصى QR من العرض
  let qrBox = Math.floor((availW - gapX * (colsTry - 1)) / colsTry);
  qrBox = Math.max(minQr, Math.min(maxQr, qrBox));

  // ننقص تدريجيًا لين يركب 100% (الطول الحقيقي)
  while (qrBox >= minQr) {
    const nameFont = qrBox <= 28 ? 7 : (qrBox <= 38 ? 8 : 9);
    const idFont   = qrBox <= 28 ? 6 : (qrBox <= 38 ? 7 : 8);

    const textH = (nameFont <= 7 ? 7 : 8) + (idFont <= 7 ? 5 : 6);
    const cellH = qrBox + textH;

    const totalH = neededRows * cellH + gapY * (neededRows - 1);
    const totalW = colsTry * qrBox + gapX * (colsTry - 1);

    const fits = totalH <= availH + 0.001 && totalW <= availW + 0.001;
    if (fits) break;

    qrBox -= 1; // ننقص 1mm لين يركب
  }

  if (qrBox < minQr) continue;

  if (!best || qrBox > best.qrBox || (qrBox === best.qrBox && colsTry > best.cols)) {
    best = { cols: colsTry, qrBox };
  }
}

if (!best) best = { cols: 3, qrBox: 32 };

const cols = best.cols;
const qrBox = best.qrBox;

  // أحجام النص ديناميكية
  const nameFont = qrBox <= 28 ? 7 : (qrBox <= 38 ? 8 : 9);
  const idFont   = qrBox <= 28 ? 6 : (qrBox <= 38 ? 7 : 8);

  const textH = (nameFont <= 7 ? 7 : 8) + (idFont <= 7 ? 5 : 6);
  const cellH = qrBox + textH;

  // ترتيب الطلاب حسب الشعبة
  const groups = groupBy(items, (x) => x.folder);
  const classNames = Object.keys(groups).sort((a, b) => a.localeCompare(b, 'ar'));

  classNames.forEach((classFolder, classIndex) => {
    if (classIndex > 0) doc.addPage();

    doc.setFont('Amiri', 'normal');
    doc.setFontSize(16);
    doc.text(`الشعبة: ${classFolder}`, margin, margin);

    doc.setDrawColor(200);
    doc.line(margin, margin + 2, pageW - margin, margin + 2);

    const list = groups[classFolder];

    // سعة الصفحة الفعلية (بعد اختيار cols/qrBox)
    const rowsCapacity = Math.floor((availH + gapY) / (cellH + gapY));
    const pageCapacity = rowsCapacity * cols;

    let idx = 0;
    let pageLocal = 0;

    while (idx < list.length) {
      if (pageLocal > 0) {
        doc.addPage();
        doc.setFont('Amiri', 'normal');
        doc.setFontSize(16);
        doc.text(`الشعبة: ${classFolder} (تكملة)`, margin, margin);
        doc.setDrawColor(200);
        doc.line(margin, margin + 2, pageW - margin, margin + 2);
      }

      const slice = list.slice(idx, idx + pageCapacity);

      slice.forEach((it, i) => {
        const col = i % cols;
        const row = Math.floor(i / cols);

        const x = startX + col * (qrBox + gapX);
        const y = startY + row * (cellH + gapY);

        // إطار القص
        doc.setDrawColor(210);
        doc.rect(x, y, qrBox, cellH);

        // QR
        doc.addImage(it.dataUrl, 'PNG', x + 2, y + 2, qrBox - 4, qrBox - 4);

        // الاسم
        doc.setFont('Amiri', 'normal');
        doc.setFontSize(nameFont);
        doc.text(String(it.name), x + qrBox - 2, y + qrBox + (nameFont <= 7 ? 4.5 : 5.2), {
          maxWidth: qrBox - 4,
          align: 'right'
        });

        // الرقم
        doc.setFontSize(idFont);
        doc.text(String(it.id), x + qrBox - 2, y + qrBox + (nameFont <= 7 ? 8.2 : 9.0), {
          maxWidth: qrBox - 4,
          align: 'right'
        });
      });

      idx += slice.length;
      pageLocal++;
    }
  });

  doc.save('QR_By_Class.pdf');
}

/* ========= مساعد: groupBy ========= */
function groupBy(arr, keyFn) {
  return arr.reduce((acc, item) => {
    const k = keyFn(item);
    (acc[k] ||= []).push(item);
    return acc;
  }, {});
}

/* ========= تشغيل التوليد ========= */
btnGenerate.addEventListener('click', async () => {
  try {
    clearUI();

    const file = fileInput.files?.[0];
    if (!file) {
      showStatus('اختر ملف Excel أول.', 'error');
      return;
    }

    showStatus('جاري قراءة الملف...', 'info');

    const rows = await readExcelRowsABC(file);
    if (!rows.length) {
      showStatus('ما لقيت بيانات في الأعمدة A,B,C من الصف A2.', 'error');
      return;
    }

    const size = Number(qrSizeEl.value) || 256;
    const margin = Number(qrMarginEl.value) || 2;
    const ec = qrECEl.value || 'M';

    showStatus(`جاري توليد ${rows.length} QR...`, 'info');

    // توليد
    for (const r of rows) {
      const dataUrl = await makeQrDataUrl(r.id, size, margin, ec);

      lastGenerated.push({
        id: r.id,
        name: r.name,
        cls: r.cls,
        folder: folderFromClass(r.cls),
        dataUrl
      });

      addCard(`${r.name} — ${r.cls}`, dataUrl);
    }

    countInfo.textContent = `${lastGenerated.length} عنصر`;
    showStatus('تم التوليد ✅ جاري الحفظ حسب النوع المختار...', 'success');

    const saveType = document.querySelector('input[name="saveType"]:checked')?.value || 'png';

    if (saveType === 'png') {
      await saveAllAsPngZip(lastGenerated);
      showStatus('تم حفظ ZIP ✅ (مجلد لكل شعبة وبداخله PNG باسم الطالب)', 'success');
    } else {
      await savePdfOnePagePerClass(lastGenerated);
      showStatus('تم حفظ PDF ✅ (Auto Fit + Auto Columns)', 'success');
    }

  } catch (err) {
    console.error(err);
    showStatus(`صار خطأ: ${err?.message || err}`, 'error');
  }
});

btnClear.addEventListener('click', () => {
  fileInput.value = '';
  clearUI();
});
