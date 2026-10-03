import time
import json
from pathlib import Path
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from winotify import Notification


# ============================================================
# تنظیمات
# ============================================================

URL = "https://www.ifb.ir/Finstars/AllCrowdFundingProject.aspx"

DRIVER_PATH = r"C:\chrome-test\chrome-win64\chromedriver.exe"

CHECK_INTERVAL = 600  # 10 دقیقه

SNAPSHOT_FILE = Path(__file__).resolve().parent / "ifb_row2_data.json"


# ============================================================
# XPath های مورد نظر
# ============================================================

XPATHS = {
    "مقدار 1 - ستون 7":
        "/html/body/div/div[3]/div[1]/div[1]/div[1]/form/div[4]/div[2]/div/table/tbody/tr[2]/td[7]",

    "مقدار 2 - ستون 1":
        "/html/body/div/div[3]/div[1]/div[1]/div[1]/form/div[4]/div[2]/div/table/tbody/tr[2]/td[1]",

    "مقدار 3 - ستون 5":
        "/html/body/div/div[3]/div[1]/div[1]/div[1]/form/div[4]/div[2]/div/table/tbody/tr[2]/td[5]",

    "مقدار 4 - ستون 2":
        "/html/body/div/div[3]/div[1]/div[1]/div[1]/form/div[4]/div[2]/div/table/tbody/tr[2]/td[2]",

    "مقدار 5 - ستون 3":
        "/html/body/div/div[3]/div[1]/div[1]/div[1]/form/div[4]/div[2]/div/table/tbody/tr[2]/td[3]",
}


# ============================================================
# اعلان
# ============================================================

def send_notification(changed_items):

    message = (
        f"{len(changed_items)} مورد تغییر کرده است.\n\n"
        + "\n".join(changed_items)
    )

    toast = Notification(
        app_id="IFB Monitor",
        title="🔔 تغییر در اطلاعات IFB",
        msg=message
    )

    toast.show()


# ============================================================
# ساخت Chrome
# ============================================================

def create_driver():

    options = Options()

    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    options.add_argument("--disable-cache")
    options.add_argument("--disk-cache-size=0")
    options.add_argument("--media-cache-size=0")

    options.add_argument("--window-size=1920,1080")

    service = Service(DRIVER_PATH)

    driver = webdriver.Chrome(
        service=service,
        options=options
    )

    driver.set_page_load_timeout(30)

    return driver


# ============================================================
# خواندن 5 مقدار
# ============================================================

def get_values(driver):

    print("در حال باز کردن سایت...", flush=True)

    try:
        driver.get(URL)
    except Exception:
        print("⚠️ زمان بارگذاری تمام شد؛ ادامه بررسی...", flush=True)

    time.sleep(5)

    print("صفحه آماده شد.", flush=True)

    values = {}

    for name, xpath in XPATHS.items():

        element = driver.find_element(
            By.XPATH,
            xpath
        )

        value = element.text.strip()

        values[name] = value

    return values


# ============================================================
# ذخیره اطلاعات
# ============================================================

def save_data(data):

    with open(
        SNAPSHOT_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            data,
            file,
            ensure_ascii=False,
            indent=2
        )


# ============================================================
# خواندن اطلاعات قبلی
# ============================================================

def load_data():

    if not SNAPSHOT_FILE.exists():
        return None

    try:

        with open(
            SNAPSHOT_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            return json.load(file)

    except Exception:

        return None


# ============================================================
# بررسی تغییرات
# ============================================================

def check_changes(current_data):

    old_data = load_data()

    # اولین اجرا
    if old_data is None:

        save_data(current_data)

        print()
        print("اولین اجرا است.")
        print("مقادیر اولیه ذخیره شدند.")
        print("فعلاً نوتیفیکیشن ارسال نمی‌شود.")

        return

    changed_items = []

    # بررسی هر 5 مقدار
    for name in XPATHS:

        old_value = old_data.get(name, "")
        new_value = current_data.get(name, "")

        if old_value != new_value:

            changed_items.append(
                f"{name}: {old_value} → {new_value}"
            )

    print()
    print(
        f"تعداد موارد تغییرکرده: {len(changed_items)}"
    )

    # ========================================================
    # حداقل 2 مورد باید تغییر کرده باشد
    # ========================================================

    if len(changed_items) >= 2:

        print()
        print("⚠️ حداقل 2 مقدار تغییر کرده است!")

        send_notification(changed_items)

        print("🔔 نوتیفیکیشن ارسال شد.")

    elif len(changed_items) == 1:

        print(
            "فقط یک مقدار تغییر کرده؛ "
            "نوتیفیکیشن ارسال نمی‌شود."
        )

    else:

        print(
            "هیچ تغییری وجود ندارد."
        )

    # وضعیت جدید را ذخیره می‌کنیم
    save_data(current_data)


# ============================================================
# یک بار بررسی
# ============================================================

def run_check():

    driver = None

    try:

        driver = create_driver()

        current_data = get_values(driver)

        print()
        print("مقادیر فعلی:")

        for name, value in current_data.items():

            print(
                f"{name}: {value}"
            )

        check_changes(current_data)

    except Exception as e:

        print()
        print("❌ خطا:")
        print(e)

    finally:

        if driver:

            driver.quit()


# ============================================================
# برنامه اصلی
# ============================================================

print("=" * 70)

print(
    "سیستم پایش ردیف دوم جدول تامین مالی جمعی IFB"
)

print("=" * 70)

print()
print("تعداد موارد تحت پایش: 5")
print("شرط اعلان: حداقل 2 تغییر")
print("فاصله بررسی: 10 دقیقه")
print("برای توقف: Ctrl+C")
print()


try:

    while True:

        print("-" * 70)

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        print(
            f"[{now}] شروع بررسی..."
        )

        run_check()

        next_time = (
            datetime.now().timestamp()
            + CHECK_INTERVAL
        )

        next_datetime = datetime.fromtimestamp(
            next_time
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        print()
        print(
            f"بررسی بعدی: {next_datetime}"
        )

        print(
            "انتظار 10 دقیقه..."
        )

        time.sleep(CHECK_INTERVAL)


except KeyboardInterrupt:

    print()
    print("=" * 70)
    print("برنامه متوقف شد.")
    print("=" * 70)