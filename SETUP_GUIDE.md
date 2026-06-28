# Policy Tracker Management System - Setup Guide

## دليل إعداد نظام متابعة السياسات

### الملفات المرفقة

| الملف | الوصف |
|-------|-------|
| `Policy_Tracker_Management_System.accdb` | قاعدة بيانات Access تحتوي على جميع الجداول والبيانات التجريبية |
| `PolicyTracker_VBA_Module.bas` | الوحدة الرئيسية - دوال حساب أيام العمل، OLA، إدارة السياسات |
| `PolicyTracker_OLA_Module.bas` | وحدة حساب OLA التفصيلي لكل إدارة ومرحلة |
| `PolicyTracker_FormEvents.bas` | أكواد أحداث النماذج (للنسخ في كل نموذج) |
| `PolicyTracker_FormCreator.bas` | وحدة إنشاء النماذج تلقائياً |

---

## خطوات الإعداد

### الخطوة 1: فتح قاعدة البيانات
1. افتح ملف `Policy_Tracker_Management_System.accdb` في Microsoft Access
2. اضغط "Enable Content" إذا ظهر تحذير الأمان

### الخطوة 2: استيراد وحدات VBA
1. اضغط `Alt + F11` لفتح محرر VBA
2. من القائمة: `File > Import File`
3. استورد الملفات التالية بالترتيب:
   - `PolicyTracker_VBA_Module.bas`
   - `PolicyTracker_OLA_Module.bas`
   - `PolicyTracker_FormEvents.bas`
   - `PolicyTracker_FormCreator.bas`

### الخطوة 3: إنشاء الاستعلامات
1. من محرر VBA، افتح نافذة Immediate (اضغط `Ctrl + G`)
2. اكتب: `CreateAllQueries` ثم اضغط Enter
3. ستظهر رسالة تأكيد

### الخطوة 4: إنشاء النماذج تلقائياً
1. في نافذة Immediate، اكتب: `CreateAllForms` ثم اضغط Enter
2. سيتم إنشاء جميع النماذج تلقائياً

### الخطوة 5: إعداد النظام
1. في نافذة Immediate، اكتب: `SetupSystem` ثم اضغط Enter

### الخطوة 6: تعيين النموذج الرئيسي
1. اذهب إلى `File > Options > Current Database`
2. في حقل "Display Form" اختر `frmMainMenu`
3. أعد فتح قاعدة البيانات

---

## هيكل النماذج

### frmMainMenu (الشاشة الرئيسية)
- عنوان: "نظام متابعة السياسات - Policy Tracker"
- 7 أزرار للتنقل بين النماذج
- يتم تحديث SLA تلقائياً عند الفتح

### frmAddPolicy (إضافة سياسة)
- حقول: رقم السياسة، الاسم، القطاع، الإدارة المالكة، المسؤول، الملاحظات
- زر حفظ يقوم بالتسجيل في tblPolicies و tblWorkflowLog

### frmUpdatePolicy (تحديث سياسة)
- قائمة منسدلة لاختيار السياسة
- عرض المعلومات الحالية تلقائياً
- تغيير المرحلة مع تسجيل في WorkflowLog
- تسجيل "تم الاستلام من" مع قائمة ديناميكية حسب المرحلة

### frmDashboard (لوحة المتابعة)
- إجمالي السياسات وتوزيعها على المراحل
- حالات OLA
- متوسط المدد
- أكثر جهة تسبب تأخير

---

## البيانات التجريبية

| رقم السياسة | الاسم | المرحلة الحالية |
|-------------|-------|-----------------|
| POL-001 | سياسة أمن المعلومات | مرحلة المراجعة الفنية |
| POL-002 | سياسة الخصوصية وحماية البيانات | مرحلة التسجيل |
| POL-003 | سياسة إدارة المخاطر | معتمدة |
| POL-004 | سياسة العمل عن بعد | ملغي / مرفوض |
| POL-005 | سياسة الاستدامة البيئية | مرحلة المراجعة العمومية |

---

## الربط مع Power BI

الجداول جاهزة للربط مع Power BI عبر:
1. `Get Data > Access Database`
2. اختر الملف `.accdb`
3. الجداول المهمة:
   - `tblWorkflowLog` - المصدر الأساسي للتحليلات الزمنية
   - `tblPolicies` - الحالة الحالية
   - `tblStages` - مراجع المراحل
   - `tblDepartments` - مراجع الإدارات
   - `tblSLASetup` - إعدادات OLA/SLA
