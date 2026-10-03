import re
import pandas as pd

# فایل خام
with open("1.txt1", "r", encoding="utf-8") as f:
    text = f.read()


# استخراج تاریخ‌ها
dates_block = re.search(
    r'categories:\s*\[(.*?)\]',
    text,
    re.DOTALL
).group(1)

dates = re.findall(
    r'"([^"]+)"',
    dates_block
)


# استخراج اعداد
values_block = re.search(
    r'data:\s*\[(.*?)\]',
    text,
    re.DOTALL
).group(1)

values = re.findall(
    r'"(\d+)"',
    values_block
)


# ساخت جدول
df = pd.DataFrame({
    "date": dates,
    "value": values
})


# ذخیره خروجی
df.to_excel(
    "marketcap_clean.xlsx",
    index=False
)

print(df.head())
print("تعداد رکورد:", len(df))