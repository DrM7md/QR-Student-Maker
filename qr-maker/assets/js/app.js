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

const codeTypeEls = document.querySelectorAll('input[name="codeType"]');
const qrSettingsBox = document.getElementById('qrSettings');
const barcodeSettingsBox = document.getElementById('barcodeSettings');

const barWidthEl = document.getElementById('barWidth');
const barHeightEl = document.getElementById('barHeight');
const showTextEl = document.getElementById('showText');

const pageTitle = document.getElementById('pageTitle');
const pageSubtitle = document.getElementById('pageSubtitle');
const modeBadge = document.getElementById('modeBadge');

let lastGenerated = []; // {id, name, cls, folder, dataUrl, codeType}

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

/* ========= اختيار النوع + ثيم ========= */
function getCodeType() {
  return document.querySelector('input[name="codeType"]:checked')?.value || 'qr';
}

function applyUiByType() {
  const t = getCodeType();
  const isBarcode = t === 'barcode';

  // إظهار/إخفاء الإعدادات
  qrSettingsBox.classList.toggle('hidden', isBarcode);
  barcodeSettingsBox.classList.toggle('hidden', !isBarcode);

  // ثيم
  document.body.classList.toggle('theme-barcode', isBarcode);

  // عنوان/بادج
  modeBadge.textContent = isBarcode ? 'BARCODE' : 'QR';
  pageTitle.textContent = isBarcode ? 'مولّد Barcode (Code 128) من Excel' : 'مولّد QR من Excel';
  pageSubtitle.textContent = isBarcode
    ? 'يقرأ الأعمدة A,B,C من الصف A2 ويولّد Code128 مناسب لقارئ الليزر'
    : 'يقرأ الأعمدة A,B,C من الصف A2 ويولّد QR ويحفظ PDF أو PNG';

  btnGenerate.textContent = isBarcode ? 'توليد باركود' : 'توليد QR';
}

applyUiByType();
codeTypeEls.forEach(el => el.addEventListener('change', applyUiByType));

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

/* ========= توليد Barcode (Code128) كـ DataURL ========= */
async function makeBarcodeDataUrl(text, barWidth, barHeight, showText) {
  const canvas = document.createElement('canvas');

  JsBarcode(canvas, String(text), {
    format: 'CODE128',
    width: Number(barWidth) || 2,
    height: Number(barHeight) || 60,
    displayValue: !!showText,
    margin: 10,              // مهم للّيزر (Quiet Zone)
    background: '#ffffff',    // أبيض صريح
    lineColor: '#000000',
    fontSize: 14,
  });

  return canvas.toDataURL('image/png');
}

/* ========= بطاقة معاينة ========= */
function addCard(title, dataUrl, codeType) {
  const isBarcode = codeType === 'barcode';

  const wrap = document.createElement('div');
  wrap.className = 'card';

  wrap.innerHTML = `
    <div class="flex items-start justify-between gap-2 mb-2">
      <div class="text-sm font-semibold break-all">${escapeHtml(title)}</div>
      <span class="badge">${isBarcode ? 'BARCODE' : 'QR'}</span>
    </div>

    <div class="bg-slate-50 rounded-lg p-2 flex items-center justify-center">
      <img src="${dataUrl}" alt="${isBarcode ? 'Barcode' : 'QR'}" class="max-w-full h-auto" />
    </div>
  `;

  grid.appendChild(wrap);
}

/* ========= حفظ PNG داخل ZIP (مجلد لكل شعبة) ========= */
async function saveAllAsPngZip(items, outName) {
  const zip = new JSZip();

  for (const it of items) {
    const folderName = it.folder || 'بدون_شعبة';
    const folder = zip.folder(folderName);

    // اسم الملف = اسم الطالب - الرقم الشخصي
    const fileName = `${safeFileName(it.name)} - ${safeFileName(it.id)}.png`;

    const base64 = it.dataUrl.split(',')[1];
    folder.file(fileName, base64, { base64: true });
  }

  const blob = await zip.generateAsync({ type: 'blob' });
  saveAs(blob, outName);
}

/* ========= حفظ PDF: QR (صفحة واحدة لكل شعبة Auto Fit + Auto Columns) ========= */
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

/* ========= حفظ PDF: Barcode (بسيط صفحة لكل طالب) ========= */
async function saveBarcodePdfOnePagePerClass(items) {
  const { jsPDF } = window.jspdf;

  // خط عربي (اختياري لكنه يحسن العربي)
  if (!window.ARABIC_FONT_BASE64 || typeof window.ARABIC_FONT_BASE64 !== 'string') {
    throw new Error('الخط العربي غير موجود. تأكد من fonts.js وربطه قبل app.js');
  }

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

  const pageW = 210, pageH = 297;

  // ✅ تقليص قوي عشان نقدر 33 في الصفحة
  const margin = 7;
  const gapX = 4;
  const gapY = 5;

  const headerH = 10;
  const startX = margin;
  const startY = margin + headerH;

  const availW = pageW - margin * 2;
  const availH = pageH - startY - margin;

  // هدفنا مثل QR
  const targetMax = 33;

  // نجرب أعمدة أكثر عشان نوصل 33
  const minCols = 3;
  const maxCols = 4;

  // نصوص صغيرة
  const nameFont = 10;
  const clsFont = 9;

  // حساب أفضل Layout يركّب 33 ويعطي أكبر باركود ممكن
  function chooseBestLayout(sampleDataUrl) {
    let best = null;

    for (let colsTry = maxCols; colsTry >= minCols; colsTry--) {
      const neededRows = Math.ceil(targetMax / colsTry);

      const cellW = (availW - gapX * (colsTry - 1)) / colsTry;

      // نسبة الصورة (باركود) من نفس الـ dataUrl
      let ratio = 0.30;
      try {
        const p = doc.getImageProperties(sampleDataUrl);
        ratio = p.height / p.width;
      } catch (_) {}

      // عرض الباركود داخل الخلية
      const pad = 2;
      let barcodeW = cellW - pad * 2;
      // حدود عشان ما يصير صغير زيادة
      barcodeW = Math.max(38, Math.min(62, barcodeW));

      const barcodeH = barcodeW * ratio;

      // ارتفاع النص فوق
      const textH = 10; // تقريب: اسم + شعبة
      const cellH = pad + textH + 2 + barcodeH + pad;

      const totalH = neededRows * cellH + gapY * (neededRows - 1);

      const fits = totalH <= availH + 0.001;
      if (!fits) continue;

      // نختار اللي يعطي أكبر باركود (أفضل للقراءة بالليزر)
      if (!best || barcodeW > best.barcodeW || (barcodeW === best.barcodeW && colsTry > best.cols)) {
        best = { cols: colsTry, cellW, cellH, barcodeW, ratio, pad };
      }
    }

    // fallback
    if (!best) {
      const colsTry = 3;
      const cellW = (availW - gapX * (colsTry - 1)) / colsTry;
      let ratio = 0.30;
      try {
        const p = doc.getImageProperties(sampleDataUrl);
        ratio = p.height / p.width;
      } catch (_) {}
      const pad = 2;
      const barcodeW = Math.max(38, Math.min(62, cellW - pad * 2));
      const barcodeH = barcodeW * ratio;
      const textH = 10;
      const cellH = pad + textH + 2 + barcodeH + pad;
      best = { cols: colsTry, cellW, cellH, barcodeW, ratio, pad };
    }

    return best;
  }

  // groupBy
  const groups = groupBy(items, (x) => x.folder);
  const classNames = Object.keys(groups).sort((a, b) => a.localeCompare(b, 'ar'));

  const drawHeader = (title) => {
    doc.setFont('Amiri', 'normal');
    doc.setFontSize(14);
    doc.text(title, margin, margin);

    doc.setDrawColor(210);
    doc.line(margin, margin + 2, pageW - margin, margin + 2);
  };

  classNames.forEach((classFolder, classIndex) => {
    const list = groups[classFolder];

    if (classIndex > 0) doc.addPage();

    const layout = chooseBestLayout(list[0].dataUrl);

    // سعة الصفحة الحقيقية حسب cellH
    const rowsCapacity = Math.max(1, Math.floor((availH + gapY) / (layout.cellH + gapY)));
    const pageCapacity = rowsCapacity * layout.cols;

    let idx = 0;
    let pageLocal = 0;

    while (idx < list.length) {
      if (pageLocal === 0) {
        drawHeader(`الشعبة: ${classFolder}`);
      } else {
        doc.addPage();
        drawHeader(`الشعبة: ${classFolder} (تكملة)`);
      }

      const slice = list.slice(idx, idx + pageCapacity);

      slice.forEach((it, i) => {
        const col = i % layout.cols;
        const row = Math.floor(i / layout.cols);

        const x = startX + col * (layout.cellW + gapX);
        const y = startY + row * (layout.cellH + gapY);

        // إطار خفيف
        doc.setDrawColor(220);
        doc.rect(x, y, layout.cellW, layout.cellH);

        const pad = layout.pad;
        const rightX = x + layout.cellW - pad;

        // ✅ فوق الباركود: الاسم + الشعبة فقط
        doc.setFont('Amiri', 'normal');

        doc.setFontSize(nameFont);
        const nameLine = doc.splitTextToSize(String(it.name), layout.cellW - pad * 2)[0] || '';
        doc.text(nameLine, rightX, y + pad + 4, { align: 'right' });

        doc.setFontSize(clsFont);
        doc.text(`الشعبة: ${String(it.cls).replace(/\//g, '-')}`, rightX, y + pad + 9, { align: 'right' });

        // ✅ الباركود (وفيه الرقم تحت الباركود من JsBarcode)
        let ratio = layout.ratio;
        try {
          const p = doc.getImageProperties(it.dataUrl);
          ratio = p.height / p.width;
        } catch (_) {}

        const barcodeW = layout.barcodeW;
        const barcodeH = barcodeW * ratio;

        const barcodeX = x + (layout.cellW - barcodeW) / 2;
        const barcodeY = y + pad + 12;

        doc.addImage(it.dataUrl, 'PNG', barcodeX, barcodeY, barcodeW, barcodeH);

        // ❌ لا نطبع ID تحت مرة ثانية
        // (خلاص يكفي اللي تحت الباركود داخل الصورة)
      });

      idx += slice.length;
      pageLocal++;
    }
  });

  doc.save('BARCODE_By_Class.pdf');
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

    const codeType = getCodeType();

    showStatus('جاري قراءة الملف...', 'info');

    const rows = await readExcelRowsABC(file);
    if (!rows.length) {
      showStatus('ما لقيت بيانات في الأعمدة A,B,C من الصف A2.', 'error');
      return;
    }

    showStatus(`جاري توليد ${rows.length} عنصر...`, 'info');

    for (const r of rows) {
      let dataUrl;

      if (codeType === 'barcode') {
        const bw = Number(barWidthEl.value) || 2;
        const bh = Number(barHeightEl.value) || 60;
        const showTxt = !!showTextEl.checked;

        dataUrl = await makeBarcodeDataUrl(r.id, bw, bh, showTxt);
      } else {
        const size = Number(qrSizeEl.value) || 256;
        const margin = Number(qrMarginEl.value) || 2;
        const ec = qrECEl.value || 'M';

        dataUrl = await makeQrDataUrl(r.id, size, margin, ec);
      }

      lastGenerated.push({
        id: r.id,
        name: r.name,
        cls: r.cls,
        folder: folderFromClass(r.cls),
        dataUrl,
        codeType
      });

      addCard(`${r.name} — ${r.cls}`, dataUrl, codeType);
    }

    countInfo.textContent = `${lastGenerated.length} عنصر`;
    showStatus('تم التوليد ✅ جاري الحفظ حسب النوع المختار...', 'success');

    const saveType = document.querySelector('input[name="saveType"]:checked')?.value || 'png';

    if (saveType === 'png') {
      const zipName = (codeType === 'barcode') ? 'BARCODE_By_Class.zip' : 'QR_By_Class.zip';
      await saveAllAsPngZip(lastGenerated, zipName);
      showStatus('تم حفظ ZIP ✅ (مجلد لكل شعبة وبداخله PNG باسم الطالب)', 'success');
    } else {
if (codeType === 'barcode') {
  await saveBarcodePdfOnePagePerClass(lastGenerated);
  showStatus('تم حفظ PDF ✅ (Barcode: 33 في الصفحة لكل شعبة)', 'success');
} else {
  await savePdfOnePagePerClass(lastGenerated);
   showStatus('تم حفظ PDF ✅ (QR)', 'success');
}

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
