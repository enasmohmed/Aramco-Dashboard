# رفع المشروع على GitHub بعد توحيد main و master

## ما تم تنفيذه محلياً

1. **حل تعارضات الدمج** في الملفات (custom_tags.py, views.py, excel-sheet-table.html) والإبقاء على نسخة HEAD (جدة والدمام والرياض، all_sheet، get_failed_shipments_count).
2. **إكمال الـ merge** بعمل commit للدمج على الفرع `main`.
3. **توحيد الفرعين:** تم جعل الفرع `master` يساوي `main` محلياً.

الآن الفرعان `main` و `master` يشيران لنفس آخر commit على جهازك.

---

## ما تفعله أنت (بعد تسجيل الدخول لـ GitHub)

افتح الترمينال داخل مجلد المشروع ونفّذ:

```bash
cd "/media/enas/01DCA27B13BCF520/Data/projects/Data Anlysis/aramco_project"

# رفع الفرع main إلى GitHub
git push origin main

# جعل الفرع master على GitHub مطابقاً لـ main (توحيدهما)
git push origin master --force-with-lease
```

إذا كان المستودع على GitHub يستخدم الفرع الافتراضي `main`، بعد هذا يكون كل شيء محدثاً ومتطابقاً.

---

## إن كان يطلب منك اسم مستخدم وكلمة مرور

- **اسم المستخدم:** حسابك على GitHub (مثل: enasmohmed)
- **كلمة المرور:** لا تستخدم كلمة مرور الحساب؛ استخدم **Personal Access Token (PAT)** من GitHub:
  1. GitHub → Settings → Developer settings → Personal access tokens
  2. إنشاء token جديد مع صلاحية `repo`
  3. استخدم الـ token مكان كلمة المرور عند تنفيذ `git push`

أو استخدم SSH بدلاً من HTTPS:

```bash
git remote set-url origin git@github.com:enasmohmed/Aramco-Dashboard.git
git push origin main
git push origin master --force-with-lease
```

(يتطلب أن يكون مفتاح SSH مضافاً لحسابك على GitHub.)

---

## الرابط

المستودع: https://github.com/enasmohmed/Aramco-Dashboard
