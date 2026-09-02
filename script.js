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

// المتغيرات العامة لتخزين الأسئلة وحالة الاختبار
let allQuestions = [];
let currentQuizQuestions = [];
let currentQuestionIndex = 0;
let score = 0;
let userAnswersMap = {};

// التنقل بين الصفحات
function showPage(pageId) {
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });
    const target = document.getElementById(pageId);
    if (target) {
        target.classList.add('active');
    }
}

// دالة تحميل وعرض الملفات في صفحة "الملفات والتصنيفات"
function loadFiles() {
    const filesListDiv = document.getElementById('filesList');
    if (!filesListDiv) return;

    filesListDiv.innerHTML = '';
    
    examFiles.forEach(file => {
        const div = document.createElement('div');
        div.style.cssText = "padding: 10px; margin: 5px 0; background: #f4f4f4; border-radius: 5px; display: flex; justify-content: space-between; align-items: center;";
        
        const span = document.createElement('span');
        span.textContent = file.replace('.xlsx', '');
        
        const btn = document.createElement('button');
        btn.textContent = 'تحميل وفحص';
        btn.style.cssText = "padding: 5px 10px; cursor: pointer;";
        btn.onclick = () => loadExcelFile(`Questions/${file}`);
        
        div.appendChild(span);
        div.appendChild(btn);
        filesListDiv.appendChild(div);
    });
}

// تحميل ملف إكسيل فردي وقراءته
async function loadExcelFile(filePath) {
    try {
        const response = await fetch(filePath);
        if (!response.ok) throw new Error("تعذر الوصول لملف الامتحان");
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
        alert(`تم تحميل الملف بنجاح! عدد الأسئلة: ${rows.length}`);
        // تخزين الأسئلة مؤقتاً لاختبار سري
        currentQuizQuestions = rows;
    } catch (e) {
        console.error(e);
        alert("حدث خطأ أثناء تحميل الملف.");
    }
}

// تعبئة قوائم الفلاتر للاختبار
function loadQuizFilters() {
    const quizFileSelect = document.getElementById('quizFile');
    if (!quizFileSelect) return;

    quizFileSelect.innerHTML = '<option value="">كل الملفات (جميع الأسئلة)</option>';
    examFiles.forEach((file, index) => {
        const opt = document.createElement('option');
        opt.value = file;
        opt.textContent = file.replace('.xlsx', '');
        quizFileSelect.appendChild(opt);
    });
}

// بدء الاختبار من إعدادات الاختبار
async function startQuiz() {
    const selectedFileElement = document.getElementById('quizFile');
    const countInput = document.getElementById('quizCount');
    const limit = countInput ? parseInt(countInput.value) || 20 : 20;

    let targetQuestions = [];

    if (selectedFileElement && selectedFileElement.value) {
        // تحميل الملف المحدد
        try {
            const filePath = `Questions/${selectedFileElement.value}`;
            const response = await fetch(filePath);
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
            targetQuestions = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        } catch (e) {
            alert("خطأ في تحميل ملف الامتحان المحدد.");
            return;
        }
    } else {
        // إذا لم يتم تحديد ملف، نقوم بدمج عينة أو تنبيه المستخدم
        alert("يرجى اختيار ملف امتحان للبدء.");
        return;
    }

    if (targetQuestions.length === 0) {
        alert("الملف الخارجي فارغ أو لا يحتوي على أسئلة.");
        return;
    }

    // خلط الأسئلة واختيار العدد المطلوبة
    currentQuizQuestions = targetQuestions.sort(() => 0.5 - Math.random()).slice(0, limit);
    currentQuestionIndex = 0;
    score = 0;
    userAnswersMap = {};

    showPage('quizPage');
    renderQuestion();
    renderQuestionNumbersGrid();
}

// عرض السؤال الحالي في صفحة الاختبار
function renderQuestion() {
    if (currentQuizQuestions.length === 0) return;

    const q = currentQuizQuestions[currentQuestionIndex];
    document.getElementById('questionNumber').textContent = currentQuestionIndex + 1;
    document.getElementById('totalQuestions').textContent = currentQuizQuestions.length;
    document.getElementById('score').textContent = score;

    // افتراض أن أعمدة الإكسيل تحمل أسماء مثل: Question, Answer1, Answer2, Answer3, Answer4 أو ما شابه
    document.getElementById('questionText').textContent = q.Question || q.سؤال || q.Title || "النص غير متوفر";

    const answersDiv = document.getElementById('answers');
    answersDiv.innerHTML = '';

    // جلب الخيارات المتوفرة في صف الإكسيل
    const options = [
        q.Option1 || q.اختيار1 || q.A,
        q.Option2 || q.اختيار2 || q.B,
        q.Option3 || q.اختيار3 || q.C,
        q.Option4 || q.اختيار4 || q.D
    ].filter(Boolean);

    options.forEach((optText, idx) => {
        const btn = document.createElement('button');
        btn.textContent = optText;
        btn.style.cssText = "display: block; width: 100%; margin: 8px 0; padding: 10px; text-align: right; cursor: pointer;";
        
        btn.onclick = () => {
            // التحقق البسيط من الإجابة (يمكن ربطها بعمود الإجابة الصحيحة CorrectAnswer في الإكسيل)
            const correctAnswer = q.Correct || q.الإجابة || q.Answer;
            if (correctAnswer && String(correctAnswer).trim() === String(optText).trim()) {
                btn.style.backgroundColor = '#d4edda';
                score += 1;
            } else {
                btn.style.backgroundColor = '#f8d7da';
            }
            document.getElementById('score').textContent = score;
        };
        answersDiv.appendChild(btn);
    });
}

function nextQuestion() {
    if (currentQuestionIndex < currentQuizQuestions.length - 1) {
        currentQuestionIndex++;
        renderQuestion();
    } else {
        showPage('resultPage');
        document.getElementById('finalScore').textContent = `${score} من ${currentQuizQuestions.length}`;
    }
}

function renderQuestionNumbersGrid() {
    const grid = document.getElementById('questionNumbersList');
    if (!grid) return;
    grid.innerHTML = '';
    currentQuizQuestions.forEach((_, idx) => {
        const span = document.createElement('button');
        span.textContent = idx + 1;
        span.style.cssText = "margin: 2px; padding: 5px 10px; cursor: pointer;";
        span.onclick = () => {
            currentQuestionIndex = idx;
            renderQuestion();
        };
        grid.appendChild(span);
    });
}
