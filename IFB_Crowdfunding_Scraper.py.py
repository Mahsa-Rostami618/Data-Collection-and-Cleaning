from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import time
import os


# =========================================================
# تنظیمات
# =========================================================

URL = "https://www.ifb.ir/Finstars/AllCrowdFundingProject.aspx"

OUTPUT_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "IFB_CrowdFunding.xlsx"
)

# مسیر ChromeDriver
CHROMEDRIVER = (
    r"C:\Users\SP\Downloads\chromedriver-win64 (2)"
    r"\chromedriver-win64\chromedriver.exe"
)


# =========================================================
# تنظیمات Chrome
# =========================================================

options = Options()

options.add_argument("--start-maximized")
options.add_argument("--disable-gpu")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--no-sandbox")


service = Service(CHROMEDRIVER)

driver = webdriver.Chrome(
    service=service,
    options=options
)

wait = WebDriverWait(driver, 30)


# =========================================================
# توابع کمکی
# =========================================================

def clean_text(text):
    """
    تمیز کردن متن فارسی
    """
    if text is None:
        return ""

    text = text.replace("\xa0", " ")
    text = text.replace("\n", " ")
    text = " ".join(text.split())

    return text.strip()


def get_table():
    """
    پیدا کردن جدول طرح‌های تأمین مالی جمعی
    """

    main_xpath = (
        "/html/body/div/div[3]/div[1]/div[1]/div[1]/form/div[4]"
    )

    main_section = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, main_xpath)
        )
    )

    tables = main_section.find_elements(
        By.TAG_NAME,
        "table"
    )

    if not tables:
        raise Exception(
            "هیچ جدولی داخل بخش موردنظر پیدا نشد."
        )

    # جدول موردنظر را بر اساس وجود عنوان «نام طرح» پیدا می‌کنیم
    for table in tables:

        table_text = clean_text(table.text)

        if "نام طرح" in table_text:
            return table

    # اگر عنوان پیدا نشد، اولین جدول را برمی‌گردانیم
    return tables[0]


def set_page_size_50():
    """
    تلاش برای تغییر تعداد رکوردها به 50
    """

    print("در حال بررسی گزینه تعداد رکوردها...")

    selects = driver.find_elements(
        By.TAG_NAME,
        "select"
    )

    for select in selects:

        try:

            options_elements = select.find_elements(
                By.TAG_NAME,
                "option"
            )

            for option in options_elements:

                if clean_text(option.text) == "50":

                    driver.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent("
                        "new Event('change', {bubbles:true})"
                        ");",
                        select,
                        option.get_attribute("value")
                    )

                    time.sleep(3)

                    print("تعداد نمایش روی 50 تنظیم شد.")

                    return True

        except Exception:
            pass

    print(
        "گزینه 50 پیدا نشد؛ با تعداد فعلی ادامه می‌دهیم."
    )

    return False


def get_headers(table):
    """
    استخراج نام ستون‌های جدول
    """

    headers = []

    theads = table.find_elements(
        By.TAG_NAME,
        "thead"
    )

    if theads:

        ths = theads[0].find_elements(
            By.TAG_NAME,
            "th"
        )

        for th in ths:

            text = clean_text(th.text)

            if text:
                headers.append(text)

    if headers:
        return headers

    # اگر thead وجود نداشت
    first_row = table.find_elements(
        By.TAG_NAME,
        "tr"
    )

    if first_row:

        cells = first_row[0].find_elements(
            By.XPATH,
            "./th|./td"
        )

        for cell in cells:

            text = clean_text(cell.text)

            if text:
                headers.append(text)

    return headers


def get_data_rows(table):
    """
    استخراج ردیف‌های واقعی جدول
    """

    rows = table.find_elements(
        By.XPATH,
        ".//tbody/tr"
    )

    valid_rows = []

    for row in rows:

        try:

            if not row.is_displayed():
                continue

            cells = row.find_elements(
                By.XPATH,
                "./td"
            )

            if not cells:
                continue

            row_text = clean_text(row.text)

            if not row_text:
                continue

            valid_rows.append(row)

        except Exception:
            continue

    return valid_rows


def get_detail_from_row(row):
    """
    کلیک روی «کلیک نمایید» همان ردیف
    و استخراج جزئیات بازشده
    """

    try:

        links = row.find_elements(
            By.XPATH,
            ".//*[contains(normalize-space(.), 'کلیک نمایید')]"
        )

        if not links:
            return ""

        click_element = links[0]

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            click_element
        )

        time.sleep(0.3)

        # متن قبل از کلیک برای تشخیص تغییر
        old_body_text = clean_text(
            driver.find_element(
                By.TAG_NAME,
                "body"
            ).text
        )

        driver.execute_script(
            "arguments[0].click();",
            click_element
        )

        time.sleep(1)

        # -------------------------------------------------
        # روش اول:
        # بررسی ردیف بعدی
        # -------------------------------------------------

        try:

            next_row = row.find_element(
                By.XPATH,
                "./following-sibling::tr[1]"
            )

            next_text = clean_text(
                next_row.text
            )

            if next_text:

                # اگر متن خیلی کوتاه و شبیه جدول بود،
                # روش‌های بعدی را امتحان می‌کنیم
                if len(next_text) > 30:

                    return next_text

        except Exception:
            pass

        # -------------------------------------------------
        # روش دوم:
        # پیدا کردن متن جزئیات در ردیف‌های بعدی
        # -------------------------------------------------

        try:

            following_rows = row.find_elements(
                By.XPATH,
                "./following-sibling::tr[position() <= 2]"
            )

            for fr in following_rows:

                text = clean_text(fr.text)

                if len(text) > 30:

                    return text

        except Exception:
            pass

        # -------------------------------------------------
        # روش سوم:
        # بررسی المنت‌هایی که پس از کلیک ظاهر شده‌اند
        # -------------------------------------------------

        detail_candidates = driver.find_elements(
            By.XPATH,
            "//*[contains(@class,'details') "
            "or contains(@class,'detail') "
            "or contains(@class,'description') "
            "or contains(@class,'child')]"
        )

        for element in detail_candidates:

            try:

                if element.is_displayed():

                    text = clean_text(
                        element.text
                    )

                    if len(text) > 30:

                        return text

            except Exception:
                pass

        # -------------------------------------------------
        # اگر جزئیات در ساختار خاصی بود، از تغییر متن
        # صفحه هم کمک می‌گیریم
        # -------------------------------------------------

        new_body_text = clean_text(
            driver.find_element(
                By.TAG_NAME,
                "body"
            ).text
        )

        if new_body_text != old_body_text:

            # متن کلی صفحه را برنمی‌گردانیم
            # مگر اینکه اطلاعات جدید پیدا شده باشد
            if len(new_body_text) > len(old_body_text):

                return new_body_text[
                    len(old_body_text):
                ].strip()

        return ""

    except Exception as e:

        print(
            "خطا در استخراج جزئیات:",
            str(e)
        )

        return ""


def close_detail_if_needed():
    """
    اگر جزئیات باز شده باشد، تلاش برای بستن آن
    """

    try:

        buttons = driver.find_elements(
            By.XPATH,
            "//button[contains(@class,'close')]"
            "|//a[contains(@class,'close')]"
            "|//*[contains(text(),'بستن')]"
        )

        for button in buttons:

            try:

                if button.is_displayed():

                    driver.execute_script(
                        "arguments[0].click();",
                        button
                    )

                    time.sleep(0.3)

                    return

            except Exception:
                pass

    except Exception:
        pass


def get_page_number():
    """
    تشخیص شماره صفحه فعلی
    """

    try:

        active = driver.find_elements(
            By.XPATH,
            "//*[contains(@class,'active') "
            "and normalize-space(text())]"
        )

        for element in active:

            text = clean_text(element.text)

            if text.isdigit():

                return text

    except Exception:
        pass

    return "?"


def go_to_next_page():
    """
    رفتن به صفحه بعد
    """

    # روش اول: لینک یا دکمه دارای next
    next_candidates = driver.find_elements(
        By.XPATH,
        "//a[contains(@class,'next')]"
        "|//button[contains(@class,'next')]"
        "|//*[contains(@aria-label,'Next')]"
        "|//*[contains(@title,'Next')]"
        "|//*[normalize-space(text())='Next']"
        "|//*[normalize-space(text())='بعدی']"
    )

    for element in next_candidates:

        try:

            if element.is_displayed() and element.is_enabled():

                classes = element.get_attribute("class") or ""

                if "disabled" in classes.lower():
                    continue

                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});",
                    element
                )

                time.sleep(0.3)

                driver.execute_script(
                    "arguments[0].click();",
                    element
                )

                time.sleep(2)

                return True

        except Exception:
            continue

    # -----------------------------------------------------
    # روش دوم:
    # پیدا کردن شماره صفحه فعال و کلیک صفحه بعد
    # -----------------------------------------------------

    try:

        page_links = driver.find_elements(
            By.XPATH,
            "//*[self::a or self::button]"
            "[normalize-space(text())]"
        )

        current_page = get_page_number()

        if current_page != "?":

            current_number = int(current_page)

            next_number = str(current_number + 1)

            for element in page_links:

                try:

                    if clean_text(element.text) == next_number:

                        if element.is_displayed():

                            driver.execute_script(
                                "arguments[0].scrollIntoView({block:'center'});",
                                element
                            )

                            time.sleep(0.3)

                            driver.execute_script(
                                "arguments[0].click();",
                                element
                            )

                            time.sleep(2)

                            return True

                except Exception:
                    continue

    except Exception:
        pass

    return False


# =========================================================
# اجرای برنامه
# =========================================================

all_data = []

try:

    print("=" * 60)
    print("شروع دریافت اطلاعات تأمین مالی جمعی IFB")
    print("=" * 60)

    print("\nدر حال باز کردن سایت...")

    driver.get(URL)

    wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "/html/body/div/div[3]/div[1]/div[1]/div[1]/form/div[4]"
            )
        )
    )

    time.sleep(3)

    print("سایت با موفقیت باز شد.")

    # -----------------------------------------------------
    # تنظیم 50 رکورد
    # -----------------------------------------------------

    set_page_size_50()

    # -----------------------------------------------------
    # پیمایش صفحات
    # -----------------------------------------------------

    page_counter = 1

    while True:

        print("\n")
        print("=" * 60)
        print(f"در حال پردازش صفحه {page_counter}")
        print("=" * 60)

        table = get_table()

        headers = get_headers(table)

        print("\nستون‌های پیدا شده:")

        for i, header in enumerate(headers):
            print(i, "=>", header)

        rows = get_data_rows(table)

        print(
            f"\nتعداد ردیف‌های این صفحه: {len(rows)}"
        )

        if not rows:

            print(
                "در این صفحه ردیفی پیدا نشد."
            )

            break

        # -------------------------------------------------
        # پردازش هر ردیف
        # -------------------------------------------------

        for row_index in range(len(rows)):

            try:

                # جدول را دوباره می‌گیریم چون بعد از کلیک
                # DOM ممکن است تغییر کند
                table = get_table()

                rows = get_data_rows(table)

                if row_index >= len(rows):
                    break

                row = rows[row_index]

                cells = row.find_elements(
                    By.XPATH,
                    "./td"
                )

                values = []

                for cell in cells:

                    values.append(
                        clean_text(cell.text)
                    )

                print(
                    f"\nردیف {row_index + 1} از {len(rows)}"
                )

                if values:

                    print(
                        " | ".join(values[:4])
                    )

                # -------------------------------------------------
                # استخراج جزئیات
                # -------------------------------------------------

                detail = get_detail_from_row(row)

                print(
                    "جزئیات:",
                    "دریافت شد" if detail else "پیدا نشد"
                )

                # -------------------------------------------------
                # ساخت رکورد
                # -------------------------------------------------

                record = {}

                # ستون‌های جدول
                for i, value in enumerate(values):

                    if i < len(headers):

                        column_name = headers[i]

                    else:

                        column_name = f"ستون {i + 1}"

                    record[column_name] = value

                # جزئیات
                record["جزئیات کامل"] = detail

                all_data.append(record)

                # -------------------------------------------------
                # اگر جزئیات باز مانده، تلاش برای بستن
                # -------------------------------------------------

                close_detail_if_needed()

                time.sleep(0.2)

            except Exception as e:

                print(
                    f"خطا در ردیف {row_index + 1}:",
                    str(e)
                )

                continue

        # -----------------------------------------------------
        # رفتن به صفحه بعد
        # -----------------------------------------------------

        print(
            f"\nصفحه {page_counter} تمام شد."
        )

        next_page = go_to_next_page()

        if not next_page:

            print(
                "صفحه بعدی پیدا نشد."
            )

            break

        page_counter += 1

        # جلوگیری از حلقه بی‌نهایت
        if page_counter > 500:

            print(
                "حداکثر 500 صفحه پردازش شد."
            )

            break

    # =========================================================
    # ساخت Excel
    # =========================================================

    print("\n")
    print("=" * 60)
    print("در حال ساخت فایل Excel...")
    print("=" * 60)

    if not all_data:

        print(
            "هیچ داده‌ای برای ذخیره پیدا نشد."
        )

    else:

        df = pd.DataFrame(all_data)

        # حذف ستون‌های کاملاً خالی
        df = df.dropna(
            axis=1,
            how="all"
        )

        # ذخیره Excel
        df.to_excel(
            OUTPUT_FILE,
            index=False,
            engine="openpyxl"
        )

        print(
            f"\nتعداد کل رکوردهای دریافت‌شده: {len(df)}"
        )

        print(
            "\nفایل Excel ساخته شد:"
        )

        print(
            OUTPUT_FILE
        )

        print(
            "\nپایان عملیات."
        )


    input(
        "\nبرای بستن Chrome کلید Enter را بزنید..."
    )


except Exception as e:

    print("\n")
    print("=" * 60)
    print("خطای کلی برنامه")
    print("=" * 60)

    print(
        str(e)
    )

    input(
        "\nبرای بستن Chrome کلید Enter را بزنید..."
    )


finally:

    driver.quit()
