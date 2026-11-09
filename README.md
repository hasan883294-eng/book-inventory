# انبار کتاب

پنل مدیریت موجودی کتاب. فرانت‌اند روی GitHub Pages و بک‌اند روی Google Apps Script (با دیتابیس Google Sheets) است.

## ساختار
- `index.html` — صفحه سایت (GitHub Pages)
- `Code.gs` — API در Apps Script (فقط POST با JSON)
- `appsscript.json` — تنظیمات Apps Script (Web App با دسترسی `ANYONE_ANONYMOUS`)

## راه‌اندازی
1. در Apps Script یک پروژه بسازید، `Code.gs` و `appsscript.json` را جایگزین کنید.
2. یک بار تابع `setup` را از ادیتور اجرا کنید و دسترسی‌ها را بپذیرید.
3. در **Project Settings → Script Properties** کلید `SITE_PASSWORD` را به رمز دلخواه تغییر دهید.
4. **Deploy → New deployment → Web app** با `Execute as: Me` و `Who has access: Anyone` و لینک `/exec` را کپی کنید.
5. در `index.html` مقدار `API_URL` را با لینک `/exec` جایگزین کنید.
6. این ریپو را در GitHub منتشر کنید و از **Settings → Pages** شاخه `main` را فعال کنید.

## امنیت
رمز و کلید نشست هرگز داخل کد یا ریپو نوشته نشوند؛ فقط در Script Properties باشند.
