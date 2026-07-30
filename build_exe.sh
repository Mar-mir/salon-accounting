#!/bin/bash
# ─── build_exe.sh ───
# اسکریپت ساخت فایل اجرایی ویندوز (.exe)
# روی ویندوز یا لینوکس اجرا کنید (نه ترمکس اندروید)
#
# پیش‌نیاز:
#   pip install pyinstaller openpyxl
#
# اجرا:
#   bash build_exe.sh
#
# خروجی:
#   پوشه dist/ حاوی فایل اجرایی

set -e

echo "🔨 در حال ساخت فایل اجرایی..."
echo ""

# بررسی نصب بودن pyinstaller
if ! command -v pyinstaller &> /dev/null; then
    echo "❌ PyInstaller نصب نیست. در حال نصب..."
    pip install pyinstaller
fi

# ساخت فایل اجرایی
pyinstaller \
    --onefile \
    --windowed \
    --name "حسابداری_سالن" \
    --add-data "data:data" \
    --clean \
    salon_app.py

echo ""
echo "✅ ساخت تمام شد!"
echo "📁 فایل اجرایی در: dist/حسابداری_سالن.exe"
echo ""
echo "نکته: فایل data/ کنار exe کپی شود یا --add-data استفاده شود"
