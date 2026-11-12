# 📊 دليل استيراد البيانات اليومية واللحظية إلى برنامج AmiBroker

يحتوي هذا المستودع على بيانات محدثة يومياً للبورصة المصرية (بيانات لحظية 1 دقيقة + بيانات تاريخية يومية ومؤشرات).

يوضح هذا الدليل كيفية تحميل البيانات واستيرادها تلقائياً إلى برنامج AmiBroker بضغطة زر واحدة باستخدام كود AFL بسيط.

---

## 📁 1. تجهيز مجلد البيانات

1. قم بإنشاء المجلد الذي تفضله على جهازك لحفظ البيانات، على سبيل المثال:
   C:\Users\YourUsername\Downloads\MarketData\
   (عدل المسار ليناسب المكان الذي ستحتفظ فيه بالداتا بعد التحميل).

2. قم بتحميل ملفات البيانات المحدثة يومياً من هذا المستودع وضعها داخل المجلد أعلاه:
   - mubasher_data.csv (البيانات اللحظية بفريم الدقيقة).
   - mubasher_history.csv (البيانات اليومية التاريخية والمؤشرات).

---

## 💻 2. كود الاستيراد (AFL Script)

افتح برنامج AmiBroker ثم اذهب إلى:
Analysis ⬅️ Formula Editor

انسخ الكود التالي والصقه بالكامل، ثم احفظ الملف باسم Auto_Import.afl (احرص على حفظه بترميز ANSI):

// Save file as ANSI
// عدل المسار بالأسفل ليطابق مسار المجلد الذي تستخدمه للداتا (مع مراعاة وضع \\ المزدوجة)
tempFolder = "C:\\Users\\YourUsername\\Downloads\\MarketData\\";

// ============================================================================
// 1. استيراد البيانات اللحظية (Intraday)
// ============================================================================
if (ParamTrigger("Import Intra", "Import Intra"))
{
    csvPath = tempFolder + "mubasher_data.csv";
    fmtPath = tempFolder + "mubasher_import.format";

    // إنشاء ملف تعريف صيغة البيانات اللحظية تلقائياً
    fFmt = fopen(fmtPath, "w");
    if (fFmt)
    {
        fputs("$FORMAT Ticker, Date_YMD, Time, Open, High, Low, Close, Volume\n", fFmt);
        fputs("$SEPARATOR ,\n", fFmt);
        fputs("$AUTOADD 1\n", fFmt);
        fputs("$OVERWRITE 1\n", fFmt);
        fputs("$NOQUOTES 1\n", fFmt);
        fclose(fFmt);
    }

    // استيراد الملف وتحديث الشارتات
    AB = CreateObject("Broker.Application");
    AB.Import(0, csvPath, fmtPath);
    AB.RefreshAll();
}

// ============================================================================
// 2. استيراد البيانات اليومية والمؤشرات (Daily History)
// ============================================================================
if (ParamTrigger("Import History", "Import History"))
{
    csvHistFile = tempFolder + "mubasher_history.csv";
    fmtHistFile = tempFolder + "mubasher_history.format";

    // إنشاء ملف تعريف صيغة البيانات اليومية تلقائياً
    fFmt = fopen(fmtHistFile, "w");
    if (fFmt)
    {
        fputs("$FORMAT Ticker, Date_YMD, Open, High, Low, Close, Volume\n", fFmt);
        fputs("$DEFINTERVAL 1440\n", fFmt); // تحديد الفاصل كيومي لمنع تداخله مع اللحظي
        fputs("$SEPARATOR ,\n", fFmt);
        fputs("$AUTOADD 1\n", fFmt);
        fputs("$OVERWRITE 1\n", fFmt);
        fputs("$NOQUOTES 1\n", fFmt);
        fclose(fFmt);
    }

    // استيراد الملف وتحديث الشارتات
    AB = CreateObject("Broker.Application");
    AB.Import(0, csvHistFile, fmtHistFile);
    AB.RefreshAll();
}

Filter = BarIndex() == LastValue(BarIndex());
AddTextColumn("اضغط على الزر المطلوب من نافذة Parameters", "الحالة");

---

## 🚀 3. خطوات الاستيراد اليومية (بضغطة زر)

1. من القائمة العلوية في إيمي بروكر، افتح نافذة التحليل:
   Analysis ⬅️ New Analysis
2. اضغط على أيقونة المجلد (Pick) واختر ملف Auto_Import.afl.
3. اضغط على زر Parameters في شريط الأدوات، ستظهر لك نافذة تحتوي على زرين:
   - 🟢 Import Intra: اضغط عليه لاستيراد البيانات اللحظية (1 دقيقة).
   - 🔵 Import History: اضغط عليه لاستيراد البيانات اليومية والمؤشرات.
4. بمجرد الضغط على أي زر، ستتم قراءة الملف المحدث وتنزيل الشموع الجديدة في قاعدة البيانات وتحديث جميع الشارتات فوراً.

---

## ⚠️ ملاحظات هامة لضبط إيمي بروكر

- إعداد قاعدة البيانات: يجب أن تكون قاعدة البيانات الخاصة بك في AmiBroker مهيأة لاستقبال البيانات اللحظية:
  - اذهب إلى: File ⬅️ Database Settings.
  - تأكد أن Base time interval مضبوط على 1 Minute.
- مسار المجلد: إذا قمت بتغيير مجلد حفظ الملفات، تأكد من تعديل السطر الخاص بـ tempFolder في كود الـ AFL ليتطابق مع المسار الجديد لديك مع استخدام شرطة مائلة مزدوجة \\.




## 🔴 ثانياً: طريقة الاستيراد إلى برنامج MetaStock

يتم تحويل البيانات إلى صيغة MetaStock عبر أداة **The DownLoader** المرفقة مع البرنامج باتباع الخطوات التالية:

### خطوات التحويل عبر The DownLoader:
1. افتح برنامج **The DownLoader**.
2. من القائمة العلوية اختر **Tools** ⬅️ ثم اضغط **Convert**.
3. في نافذة التحويل (Conversion Options):
   - **Source (المصدر):**
     - **File Type:** اختر `ASCII Text`.
     - **Folder:** حدد المسار الذي وضعت به الملفات (مثال: `C:\MarketData\`).
     - **File Name:** اختر ملف `mubasher_history.csv` أو `mubasher_data.csv`.
   - **Destination (الهدف):**
     - **File Type:** اختر `MetaStock`.
     - **Folder:** حدد مجلد حفظ داتا الميتستوك على جهازك (مثال: `C:\MetaStock Data\`).
4. اضغط على زر **Options**:
   - تحت تبويب **Source**:
     - **Field Delimiter:** اختر `Comma` (,).
     - **Date Order:** اختر `Year/Month/Day` (Y-M-D).
     - **Fields:** تأكد من ترتيب الحقول كالتالي:
       - لليومي: `Ticker, Date, Open, High, Low, Close, Volume`.
       - للحظي: `Ticker, Date, Time, Open, High, Low, Close, Volume`.
5. اضغط **OK** ثم اضغط **OK** مرة أخرى لبدء التحويل.
6. افتح برنامج MetaStock وستجد جميع الأسهم والمؤشرات تم تحديثها وجاهزة للفتح.
