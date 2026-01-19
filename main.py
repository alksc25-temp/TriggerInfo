import os
import re
import requests
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# ---------------- FETCH IPO DATA ----------------

def get_ipos():
    print("-- Starting IPO data extraction")
    today = datetime.today().date()

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    print("-- Launching Chrome WebDriver in headless mode")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    print("-- Navigating to IPO GMP report page")
    driver.get("https://www.investorgain.com/report/live-ipo-gmp/331/all/")

    print("-- Waiting for IPO table to load")
    table = wait.until(EC.presence_of_element_located((By.ID, "report_table")))
    rows = table.find_elements(By.TAG_NAME, "tr")

    ipo_data = []
    print("-- Extracting IPO rows")

    for row in rows[1:]:  # skip header
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) > 8:
            name = cols[0].text.strip()
            gmp_text = cols[1].text.strip()
            sub = cols[3].text.strip()
            start = cols[7].text.strip()
            end = cols[8].text.strip()

            print(f"-- Processing IPO: {name}")

            # Extract GMP percentage
            match = re.search(r"\(([\d\.]+)%\)", gmp_text)
            gmp_value = float(match.group(1)) if match else 0

            try:
                def extract_date(text, today):
                    match = re.search(r"\d{1,2}-[A-Za-z]{3}", text)
                    if match:
                        return datetime.strptime(match.group(), "%d-%b").date().replace(year=today.year)
                    return None

                start_date = extract_date(start, today)
                end_date = extract_date(end, today)
                print(f"-- Extracted dates: Start={start_date}, End={end_date}")

            except Exception as e:
                print("-- Date extraction failed:", e)
                continue

            ipo_data.append((name, gmp_value, start_date, end_date, sub))

    driver.quit()
    print(f"-- IPO data extraction complete. Total IPOs found: {len(ipo_data)}")
    return ipo_data

# ---------------- TELEGRAM ----------------

def send_telegram_message(message):
    TELEGRAM_TOKEN = os.getenv("TG_BOT_TOKEN")
    TELEGRAM_CHAT_ID = os.getenv("TG_CHAT_ID")

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": message}

    try:
        requests.post(url, data=payload, timeout=10)
        print("-- Telegram message sent successfully")
    except Exception as e:
        print("-- Telegram send failed:", e)

# ---------------- MAIN VALIDATION LOGIC ----------------

def process_ipos(ipos):
    print("-- Starting IPO validation logic")

    today = datetime.today().date()
    print(f"-- Today date: {today}")

    for name, gmp, start, end, sub in ipos:

        if not end:
            print(f"-- Skipping IPO {name} due to missing end date")
            continue

        # 🔹 Check ONLY if today is closing day OR one day before closing
        if today == end or today == (end - timedelta(days=1)):
            print(f"-- IPO in alert window: {name}, End={end}, GMP={gmp}")

            try:
                if float(gmp) >= 5:   # ✅ Your rule
                    print(f"-- GMP PASSED for {name}, GMP={gmp}")

                    day_text = "Closing Today" if today == end else "Closing Tomorrow"

                    message = (
                        f"🚀 IPO PROCEED ALERT ({day_text})\n\n"
                        f"Name: {name}\n"
                        f"GMP: {gmp}%\n"
                        f"Subscription: {sub}\n"
                        f"Start Date: {start}\n"
                        f"End Date: {end}\n\n"
                        f"Status: PROCEED"
                    )

                    send_telegram_message(message)

                else:
                    print(f"-- GMP below threshold for {name}, GMP={gmp}")

            except Exception as e:
                print(f"-- Error processing GMP for {name}: {e}")

        else:
            print(f"-- IPO {name} not in closing window (Today={today}, End={end}), skipping")


# ---------------- RUN DAILY ----------------

print("-- Script execution started")
ipos = get_ipos()
process_ipos(ipos)
print("-- Script execution finished")
