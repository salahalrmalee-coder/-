const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const db = require('./database');

// إنشاء تطبيق Express
const app = express();
const PORT = process.env.PORT || 3000;

// إنشاء مجلد uploads إذا لم يكن موجودًا
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

// تكوين multer لرفع الملفات
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir);
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({ 
    storage: storage,
    fileFilter: function (req, file, cb) {
        const filetypes = /xlsx|xls/;
        const mimetype = filetypes.test(file.mimetype);
        const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
        
        if (mimetype && extname) {
            return cb(null, true);
        }
        cb(new Error('الملف يجب أن يكون بصيغة Excel (xlsx أو xls)'));
    }
});

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

// مسار لاستيراد ملف Excel
app.post('/api/import/excel', upload.single('file'), async (req, res) => {
    console.log('--- بدء عملية استيراد ملف ---');
    console.log('الملف المستلم:', req.file ? req.file.originalname : 'لا يوجد ملف');
    
    if (!req.file) {
        console.error('خطأ: لم يتم تحميل أي ملف');
        return res.status(400).json({ 
            success: false,
            message: 'لم يتم تحميل أي ملف',
            imported: 0,
            failed: 0,
            errors: ['لم يتم العثور على ملف مرفق']
        });
    }

    try {
        // قراءة ملف Excel
        console.log('جاري قراءة ملف Excel...');
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        
        console.log(`تم العثور على ${data.length} سجل في الملف`);
        
        if (data.length === 0) {
            console.error('خطأ: الملف فارغ');
            return res.status(400).json({
                success: false,
                message: 'الملف لا يحتوي على بيانات',
                imported: 0,
                failed: 0,
                errors: ['الملف فارغ أو لا يحتوي على بيانات']
            });
        }

        // معالجة البيانات
        console.log('بدء معالجة البيانات...');
        let importedCount = 0;
        let failedCount = 0;
        const errors = [];

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowNumber = i + 2; // +2 لأن الصف الأول هو العناوين والترقيم يبدأ من 0
            
            try {
                console.log(`\n--- معالجة السطر ${rowNumber} ---`);
                console.log('بيانات السطر:', JSON.stringify(row, null, 2));
                
                // تخطي الصفوف الفارغة
                if (!row['الاسم الكامل'] && !row['full_name']) {
                    console.log(`تم تخطي السطر ${rowNumber}: لا يوجد اسم`);
                    continue;
                }

                // تحضير البيانات
                const controllerData = {
                    full_name: row['الاسم الكامل'] || row['full_name'] || '',
                    birth_date: row['تاريخ الميلاد'] || row['birth_date'] || '',
                    license_number: row['رقم الرخصة'] || row['license_number'] || '',
                    qualification: row['الأهلية'] || row['qualification'] || '',
                    workplace: row['مكان العمل'] || row['workplace'] || '',
                    // ... باقي الحقول
                };

                console.log('بيانات جاهزة للإدراج:', JSON.stringify(controllerData, null, 2));

                // إدراج البيانات في قاعدة البيانات
                const stmt = db.prepare(`
                    INSERT OR REPLACE INTO controllers (
                        full_name, birth_date, license_number, 
                        qualification, workplace
                        -- ... باقي الحقول
                    ) VALUES (
                        @full_name, @birth_date, @license_number,
                        @qualification, @workplace
                        -- ... باقي القيم
                    )
                `);

                const result = stmt.run(controllerData);
                console.log(`تمت إضافة/تحديث السجل برقم: ${result.lastInsertRowid}`);
                importedCount++;
                
            } catch (error) {
                console.error(`خطأ في السطر ${rowNumber}:`, error.message);
                errors.push(`سطر ${rowNumber}: ${error.message}`);
                failedCount++;
            }
        }

        // حذف الملف المؤقت
        try {
            if (fs.existsSync(req.file.path)) {
                fs.unlinkSync(req.file.path);
                console.log('تم حذف الملف المؤقت بنجاح');
            }
        } catch (error) {
            console.error('خطأ في حذف الملف المؤقت:', error.message);
            // لا نعيد خطأ هنا لأن العملية اكتملت بنجاح
        }

        console.log('--- انتهت عملية الاستيراد ---');
        console.log(`العدد الإجمالي للسجلات: ${data.length}`);
        console.log(`تم استيراد: ${importedCount} سجل`);
        console.log(`فشل استيراد: ${failedCount} سجل`);
        
        // إرسال الرد
        const response = {
            success: true,
            message: `تم استيراد ${importedCount} سجل بنجاح`,
            imported: importedCount,
            failed: failedCount,
            total: data.length
        };

        if (errors.length > 0) {
            response.errors = errors;
        }

        res.json(response);
        
    } catch (error) {
        console.error('خطأ في معالجة الملف:', error);
        
        // محاولة حذف الملف المؤقت في حالة الخطأ
        try {
            if (req.file && fs.existsSync(req.file.path)) {
                fs.unlinkSync(req.file.path);
            }
        } catch (deleteError) {
            console.error('خطأ في حذف الملف المؤقت بعد الخطأ:', deleteError);
        }
        
        res.status(500).json({
            success: false,
            message: 'حدث خطأ أثناء معالجة الملف',
            error: error.message,
            stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
});

// تعريف المسارات (Routes)
app.get('/api/controllers', (req, res) => {
    try {
        const controllers = db.getAllControllers();
        res.json(controllers);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/controllers', (req, res) => {
    try {
        const id = db.addController(req.body);
        res.status(201).json({ id, ...req.body });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.put('/api/controllers/:id', (req, res) => {
    try {
        const updated = db.updateController(req.params.id, req.body);
        if (updated) {
            res.json({ id: req.params.id, ...req.body });
        } else {
            res.status(404).json({ error: 'لم يتم العثور على المراقب الجوي' });
        }
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.delete('/api/controllers/:id', (req, res) => {
    try {
        const deleted = db.deleteController(req.params.id);
        if (deleted) {
            res.json({ message: 'تم حذف المراقب الجوي بنجاح' });
        } else {
            res.status(404).json({ error: 'لم يتم العثور على المراقب الجوي' });
        }
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// تقديم الملفات الثابتة
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// بدء تشغيل الخادم
app.listen(PORT, () => {
    console.log(`الخادم يعمل على المنفذ ${PORT}`);
    console.log(`رابط التطبيق: http://localhost:${PORT}`);
    console.log(`مسار حفظ الملفات المؤقتة: ${uploadsDir}`);
});
