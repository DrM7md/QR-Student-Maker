<div align="center">

# QR Student Maker · مولّد رموز الطلاب

**توليد رموز QR أو باركود لكل طالب من ملف Excel، وتصديرها PDF (صفحة لكل شعبة) أو صور PNG في ملف مضغوط.**

![HTML](https://img.shields.io/badge/HTML-E34F26?logo=html5&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?logo=javascript&logoColor=black)
![Tailwind](https://img.shields.io/badge/Tailwind-06B6D4?logo=tailwindcss&logoColor=white)

</div>

---

## المزايا
- **قراءة Excel مباشرة** (`.xlsx` / `.xls`) داخل المتصفح، دون رفع أي بيانات لخادم.
- **نوعان من الرموز:** QR أو Barcode (Code 128).
- **تصدير PDF** صفحة لكل شعبة، بخط عربي سليم.
- **تصدير PNG** في ملف ZIP، مقسَّم إلى مجلدات حسب الشعبة.
- **إزالة التكرار** تلقائيًا للطالب المكرر في الشعبة نفسها.

## صيغة ملف Excel
الورقة الأولى، والبيانات تبدأ من الصف 2 (الصف الأول للعناوين):

| A | B | C |
|---|---|---|
| الرقم الشخصي | اسم الطالب | الشعبة |

## التشغيل
لا يحتاج تثبيتًا: افتح `qr-maker/index.html` في المتصفح، ثم اختر ملف Excel ونوع الكود واضغط توليد.

## البنية
| المسار | الدور |
|---|---|
| `qr-maker/index.html` | الواجهة |
| `qr-maker/assets/js/app.js` | قراءة Excel، التوليد، التصدير |
| `qr-maker/assets/js/fonts.js` | الخط العربي لملفات PDF |
| `qr-maker/assets/css/styles.css` | التنسيق |

المكتبات: SheetJS، QRious، JsBarcode، jsPDF، JSZip، FileSaver.

## ☕ ادعم المشروع

إذا أفادك هذا المشروع، يمكنك دعم تطويره عبر PayPal:

[![PayPal](https://img.shields.io/badge/PayPal-Support-00457C?logo=paypal&logoColor=white)](https://paypal.me/SamtSoftware)

👉 https://paypal.me/SamtSoftware
