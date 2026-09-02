// قائمة الـ 58 امتحان الموجودة في مجلد Questions
const examFiles = [
    "1-اسبقيات استخدام وسائل امن المواصلات.xlsx",
    "2-التهديدات الامنية على وسائل المواصلات الاشارية.xlsx",
    "3-وثائق الرمز المستخدمة فى القوات المسلحة.xlsx",
    "إجراءات الوقاية من الأسلحة النووية.xlsx",
    "اجراءات تنظيم معركة.xlsx",
    "إستخدام اللواء المدرع فر م فرميكا الفرقة المدرعة.xlsx",
    "استخدام اللواء المشاة.xlsx",
    "اسلوب إعداد البحوث العسكرية.xlsx",
    "إسلوب تنظيم وإدارة معركة نطاق التغطية.xlsx",
    "إعادة التجميع وغيار القوات.xlsx",
    "أعمال قتال الإبرار الجوى - البحرى التكتيكى.xlsx",
    "أعمال قتال ل 15 مقل.xlsx",
    "أعمال قتال ل 16 مش - فر 16 مش.xlsx",
    "أقسام ومهام الشئون الإدارية_.xlsx",
    "الإتجاهات وكيفية تعديلها - سيكولوجية التعلم والتدريب.xlsx",
    "الإشراف على التدريب.xlsx",
    "الإعداد والتحضير لمشروع الملاحة البرية.xlsx",
    "الأقمار الصناعية وإستخداماتها-1.xlsx",
    "الأمن السيبراني.xlsx",
    "البيان العملى.xlsx",
    "التأمين الإدارى للواء فى الدفاع.xlsx",
    "التأمين الإدارى للواء فى الهجوم من الإتصال على دفاعات العدو المجهزة.xlsx",
    "التأمين الإلكترونى.xlsx",
    "التجهيز الهندسى لموقع ك.xlsx",
    "الخواص الفنية والتك لمعدات كيما -ك مش (معدات طهر- أجهزة سطع كيما - شع).xlsx",
    "الدروس المستفادة من حرب أكتوبر.xlsx",
    "الذكاء الإجتماعي.xlsx",
    "الذكاء الإصطناعي.xlsx",
    "العمليات الدفاعية الإسرائيلية.xlsx",
    "العمليات الهجومية الإسرائيلية.xlsx",
    "الغازات الحربية.xlsx",
    "اللواء المدرع المستقل الفرقة المدرعة.xlsx",
    "المسير والمعركة التصادمية.xlsx",
    "المعركة الدفاعية.xlsx",
    "المعركة الهجومية.xlsx",
    "الهجوم فى حالات خاصة.xlsx",
    "تحديد الإحداثيات.xlsx",
    "تخطيط و تنظيم التدريب.xlsx",
    "تنظيم وإستخدام عناصر الإستطلاع.xlsx",
    "تنظيم واستخدام عناصر المهندسيين العسكريين.xlsx",
    "تنظيم واستخدام قوات الدفاع الجوى.xlsx",
    "تنظيم واستخدام قوات المدفعية.xlsx",
    "تنظيم وإمكانيات وحدات الإستطلاع العام مستوى ل  فر.xlsx",
    "تنظيمات العدو.xlsx",
    "حرب المعلومات.xlsx",
    "خطة رفع حالات الاستعداد القتالى ( ك ).xlsx",
    "دور الأفرع الرئيسية والأسلحة والقوات المختلفة لصالح أعمال الحرب الإلكترونية.xlsx",
    "دور القوات المسلحة في ثورة 25 يناير.xlsx",
    "طرق إنشاء حقول الألغام-1.xlsx",
    "طرق فتح الثغرات فى حقول الألغام-1.xlsx",
    "فائق الصعوبة-1.xlsx",
    "فائق الصعوبة-2.xlsx",
    "فائق الصعوبة-3.xlsx",
    "مجمع-1.xlsx",
    "مجمع-2.xlsx",
    "مجمع-3.xlsx",
    "مخاطر أستخدام الأنترنت.xlsx",
    "معدلات الأداء لتنفيذ التجهيز الهندسى-1.xlsx"
];

// دالة توليد وعرض قائمة الامتحانات في الصفحة
function renderExamList() {
    const listContainer = document.getElementById("exams-list"); 
    if (!listContainer) return;
    
    listContainer.innerHTML = "";

    examFiles.forEach(fileName => {
        const button = document.createElement("button");
        // إزالة صيغة .xlsx من اسم الزر ليكون شكل العرض أنيقاً
        button.textContent = fileName.replace(".xlsx", "");
        button.className = "exam-btn"; // يمكنك إضافة تصميم لهذا الكلاس في ملف الـ CSS
        
        // عند الضغط على الامتحان، يتم توجيهه للمجلد الصحيح Questions/
        button.onclick = () => loadAndReadExcel(`Questions/${fileName}`);
        listContainer.appendChild(button);
    });
}

// دالة جلب وقراءة ملف الإكسيل باستخدام مكتبة SheetJS (XLSX)
async function loadAndReadExcel(filePath) {
    try {
        // إظهار رسالة تحميل اختيارية للمستخدم
        console.جاري_ التحميل("جاري تحميل الامتحان...");

        const response = await fetch(filePath);
        if (!response.ok) throw new Error("فشل في الوصول إلى ملف الامتحان.");
        
        const arrayBuffer = await response.arrayBuffer();
        
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const excelData = XLSX.utils.sheet_to_json(worksheet);

        // تشغيل واجهة الاختبار بالأسئلة المستخرجة
        startQuiz(excelData);
    } catch (error) {
        console.error("خطأ أثناء تحميل الملف:", error);
        alert("حدث خطأ أثناء تحميل ملف الامتحان، تأكد من اتصال الإنترنت أو صحة المسار.");
    }
}

// دالة بدء الاختبار (يمكنك ربطها بمنطق عرض الأسئلة الخاص بك)
function startQuiz(questionsArray) {
    console.log("تم تحميل الأسئلة بنجاح:", questionsArray);
    // هنا تكتب الكود الخاص بك لعرض السؤال الأول، الخيارات، وتتبع النتيجة
}

// تشغيل دالة عرض القائمة فور تحميل الصفحة
window.onload = renderExamList;
