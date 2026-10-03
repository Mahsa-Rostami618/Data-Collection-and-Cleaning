from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import tempfile
import re

input_file = Path(r"Z:\Mahsa Rostami\14050601.xlsx")
output_file = Path(r"Z:\Mahsa Rostami\14050601_Unlocked.xlsx")

with tempfile.TemporaryDirectory() as temp:

    temp = Path(temp)

    # باز کردن فایل XLSX
    with ZipFile(input_file, "r") as z:
        z.extractall(temp)

    # بررسی تمام Sheet ها
    worksheets = temp / "xl" / "worksheets"

    removed = 0

    for xml_file in worksheets.glob("*.xml"):

        text = xml_file.read_text(
            encoding="utf-8"
        )

        # حذف sheetProtection
        new_text, count = re.subn(
            r"<sheetProtection\b[^>]*/>",
            "",
            text,
            flags=re.IGNORECASE
        )

        if count:

            xml_file.write_text(
                new_text,
                encoding="utf-8"
            )

            removed += count

            print(
                f"🔓 قفل {xml_file.name} حذف شد"
            )

    # ساخت فایل Excel جدید
    with ZipFile(
        output_file,
        "w",
        ZIP_DEFLATED
    ) as z:

        for file in temp.rglob("*"):

            if file.is_file():

                z.write(
                    file,
                    file.relative_to(temp)
                )

print()
print("======================================")
print("✅ عملیات تمام شد")
print(f"🔓 تعداد قفل حذف شده: {removed}")
print()
print("📁 فایل بدون قفل:")
print(output_file)
print("======================================")