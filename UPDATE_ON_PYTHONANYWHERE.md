# تحديث المشروع على PythonAnywhere من غير حذف

## الطريقة 1: إذا المشروع على PythonAnywhere مربوط بـ Git (مستنسخ من GitHub)

1. ادخل على **PythonAnywhere** → **Dashboard** → تبويب **Consoles**.
2. افتح **Bash console** (أو استخدم الـ console اللي كنت بتشغّل منه المشروع).
3. ادخل لمجلد المشروع، مثلاً:
   ```bash
   cd ~/Aramco-Dashboard
   # أو المسار اللي مشروعك فيه، غالباً شيء مثل:
   # cd ~/aramco_project
   ```
4. جب آخر التحديثات من GitHub:
   ```bash
   git fetch origin
   git checkout main
   git pull origin main
   ```
5. إذا طلب منك username/password لـ GitHub، استخدم **Personal Access Token** مكان كلمة المرور.
6. بعد الـ pull، أعد تحميل الـ Web App:
   - روح **Web** tab في الـ Dashboard.
   - اضغط الزر **Reload** بجانب عنوان الموقع (الزر الأخضر).

بهذا يكون الكود الجديد على السيرفر والموقع محدّث من غير ما تحذف أي حاجة.

---

## الطريقة 2: إذا المشروع على PythonAnywhere مش مربوط بـ Git (رفعته يدوي أو zip)

### أ) ربط المشروع بـ Git على PythonAnywhere (مرة واحدة)

1. من الـ **Bash console** على PythonAnywhere:
   ```bash
   cd ~
   # لو المشروع في مجلد اسمه مثلاً mysite أو Aramco-Dashboard
   cd اسم_مجلد_المشروع
   git init
   git remote add origin https://github.com/enasmohmed/Aramco-Dashboard.git
   git fetch origin
   git checkout -b main origin/main
   ```
2. من هنا فصاعداً استخدم **الطريقة 1** (git pull) عشان أي تحديث.

### ب) أو رفع الملفات الجديدة يدوياً (بدون Git)

1. من جهازك: ارفع الملفات اللي اتغيّرت عبر **Files** في PythonAnywhere (نسخ/لصق أو upload).
2. أو: من جهازك شغّل المشروع من مجلد المشروع واعمل **zip** للمشروع (بدون `env` وبدون `.git` لو حابب)، ثم من **Files** على PythonAnywhere ارفع الـ zip وافك الضغط فوق الملفات القديمة (Replace).
3. بعد ما تتأكد إن الملفات الجديدة موجودة، روح **Web** → **Reload** للموقع.

---

## إذا ظهر: "Your local changes would be overwritten by merge"

معناها فيه ملفات معدّلة على السيرفر (مثل `db.sqlite3` أو ملفات في `media/`) والـ pull هيمسحها. الحل:

**1. خبّئ التعديلات المحلية (stash):**
```bash
git stash push -m "pythonanywhere local" db.sqlite3 media/excel_uploads/latest.xlsx
```

**2. اعمل pull عادي:**
```bash
git pull origin main
```

**3. بعد الـ pull:**
- **لو عايز كود الجديد فقط وتقبل نسخة الملفات من الجيت:** اتجاهل الـ stash (ما تعملش حاجة)، وامشي. قاعدة البيانات والملفات المرفوعة على السيرفر هتبقى اللي في الريبو.
- **لو عايز ترجّع بيانات السيرفر (مثلاً قاعدة البيانات أو الإكسل اللي كان مرفوع):**
  ```bash
  git stash pop
  ```
  لو طلع تعارض (conflict)، ممكن تتجاهله وتخلي الملفات زي ما هي على السيرفر.

**4. من تبويب Web اضغط Reload للموقع.**

---

## ملخص سريع (المشروع مربوط بـ Git)

```text
Bash على PythonAnywhere:
  cd ~/مجلد_المشروع
  git pull origin main

ثم من صفحة Web اضغط Reload.
```

لو طلع خطأ "local changes would be overwritten" استخدم الـ stash ثم pull كما فوق.

بكده التحديث يترفع من غير ما تحذف المشروع أو تعيده من الأول.
