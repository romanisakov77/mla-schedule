
import datetime
import os
import pickle
import subprocess
import shutil
import pytz

from ics import Calendar, Event
from pytz import timezone
from datetime import datetime, date, timedelta, timezone
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select

# Constants
MYFBO_USERNAME = "roman@isakovthepilot.com"
MYFBO_PASSWORD = "Delta2025!"
MYFBO_URL = "https://prod.myfbo.com/link.asp?fbo=malh"
TIMEZONE = "Pacific/Honolulu"
CALENDAR_DEPTH = 30

#SCHEDULE_PATH = "/Users/romanisakov/Scripts/mla-schedule/schedule.ics"
REPO_PATH = "/Users/romanisakov/Scripts/mla-schedule"

RESOURCE_LIST = ["Roman Isakov", "Kira Hayamoto", "Logan Nagel", "Ben Fouts", "Tyler Walton", "Christoph Maurer"]


def login_to_myfbo(driver):
    driver.get(MYFBO_URL)

    try:
        continue_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.NAME, "anyway"))
        )
        print("⚠️ 'Continue Anyway' page detected. Clicking through...")
        continue_btn.click()
    except:
        print("✅ No interstitial — proceeding to login form.")

    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "myfbo2"))
    )

    email_input = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.NAME, "email"))
    )
    email_input.send_keys(MYFBO_USERNAME)

    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "password"))
    )
    password_input.send_keys(MYFBO_PASSWORD)

    submit_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//input[@type="submit" or @type="button" or @type="image"]'))
    )
    submit_btn.click()
    print("✅ Login button clicked.")

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "myfbo2"))
    )
    print("✅ Re-entered main frame after login.")

def navigate_to_schedule(driver):
    print("🔄 Navigating to Reservations...")

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "myfbo2")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "tf")))
    print("✅ Entered 'tf' frame (top menu)")

    schedules_td = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "SchedRpt"))
    )
    ActionChains(driver).move_to_element(schedules_td).perform()
    print("🖱️ Hovered over 'Schedules' tab")

    WebDriverWait(driver, 5).until(lambda d: True)

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "myfbo2")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "tf")))

    reservation_link = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'menu_events.asp')]"))
    )
    reservation_link.click()
    print("✅ Clicked 'Reservation Lists'")

 

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "myfbo2")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "wa")))

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.NAME, "dfr"))
    )
    print("📆 Reservation list page loaded.")



    today = date.today().strftime("%-m/%-d/%y")  # e.g., 5/15/25
    future = (date.today() + timedelta(days=CALENDAR_DEPTH)).strftime("%-m/%-d/%y")

    # Set 'From' date
    from_input = driver.find_element(By.NAME, "dfr")
    driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", from_input, today)

    # Set 'To' date
    to_input = driver.find_element(By.NAME, "dto")
    driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", to_input, future)


    sres_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "sres"))
    )
    Select(sres_dropdown).select_by_visible_text(RESOURCE_NAME)


    # Click "Date / Time" under Active Reservations
    driver.find_element(
        By.XPATH,
        "//input[@type='button' and @onclick=\"actGo('all_schedule.asp?order=1&tor=F');\"]"
    ).click()

def parse_reservations(driver):
    def extract_equipment():
        try:
            equip_td = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//a[contains(text(), 'Equipment')]/parent::td/following-sibling::td[1]"
                ))
            )
            return equip_td.text.strip()
        except Exception as e:
            print(f"⚠️ Equipment extraction failed: {e}")
            return ""


    def go_back_to_schedule():
        try:
            back_link = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Calling Page"))
            )
            back_link.click()
        except:
            driver.back()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "TABLE_1")))

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "myfbo2")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "wa")))

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "TABLE_1")))

    events = []

    i = 0
    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "#TABLE_1 tbody tr")
        if i >= len(rows):
            break

        row = rows[i]
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 6:
            i += 1
            continue

        from_cell = cells[1].text.strip()
        to_cell = cells[2].text.strip()
        rmk_cell = cells[8]
        base_name = cells[7].text.strip()

        raw_customer = cells[5].text.strip().replace('\u200b', '').replace('\ufeff', '').replace('\xa0', ' ')
        if raw_customer.strip().lower() == "i'm busy":
            print(f"⏭️ Skipped busy block from {from_cell} to {to_cell}")
            i += 1
            continue

        name_parts = raw_customer.split()
        customer = f"{name_parts[0]} {name_parts[-1]}" if len(name_parts) >= 2 else raw_customer

        try:
            start_hst = datetime.strptime(from_cell, "%m/%d/%y %H:%M")
            end_hst = datetime.strptime(to_cell, "%m/%d/%y %H:%M")
        except ValueError as e:
            print(f"⚠️ Error parsing times: {e}")
            continue

        # Convert HST to UTC
        start_utc = (start_hst + timedelta(hours=10)).replace(tzinfo=timezone.utc)
        end_utc = (end_hst + timedelta(hours=10)).replace(tzinfo=timezone.utc)

        try:
            img = rmk_cell.find_element(By.TAG_NAME, "img")
            remark = img.get_attribute("title").strip()
            #print(f"✅ Extracted remark: {remark}")
        except:
            remark = ""

        # Click "Detail" button to get equipment info
        try:
            detail_button = cells[10].find_element(By.XPATH, ".//input[@value='Detail']")
            detail_button.click()

            equipment = extract_equipment()
            go_back_to_schedule()
        except Exception as e:
            print(f"❌ Failed to get equipment: {e}")
            equipment = ""

        if equipment:
            tail_number = equipment.split()[-1]
            summary = f"Flight - {customer} ({tail_number})"
        else:
            summary = f"Ground - {customer}"


        event = {
            "summary": summary,
            "description": remark if remark else "",
            "location": base_name if base_name else "",
            "start": start_utc,
            "end": end_utc,
        }

        print(f"✅ Parsed event: {event['summary']} | {base_name} | {remark}")
        events.append(event)

        i += 1

    print(f"✅ Total events parsed: {len(events)}")
    return events


def generate_ics(events, filename):
    # Delete the file if it exists
    if os.path.exists(filename):
        os.remove(filename)
        print(f"🗑️ Removed existing file: {filename}")

    cal = Calendar()
    for ev in events:
        e = Event()
        e.name = ev["summary"]
        e.begin = ev["start"].strftime("%Y-%m-%dT%H:%M:%SZ")  # UTC time with 'Z'
        e.end = ev["end"].strftime("%Y-%m-%dT%H:%M:%SZ")
        e.location = ev["location"]
        e.description = ev["description"]
        e.alarms = []
        cal.events.add(e)
    with open(filename, "w") as f:
        f.writelines(cal)
    print(f"✅ Exported {len(events)} events to {filename}")

# def publish_to_github(repo_path=REPO_PATH):

#     # Generate timestamp for commit message
#     now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#     commit_message = f"Update schedule ({now})"

#     # Commit and push
#     subprocess.run(["git", "-C", repo_path, "add", "schedule.ics"])
#     subprocess.run(["git", "-C", repo_path, "commit", "-m", commit_message])
#     subprocess.run(["git", "-C", repo_path, "push"])

def main():
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)


    #repo_path = "/Users/romanisakov/Scripts/mla-schedule"

    try:
        login_to_myfbo(driver)

        for name in RESOURCE_LIST:
            print(f"\n📅 Generating schedule for: {name}")
            global RESOURCE_NAME
            RESOURCE_NAME = name  # this updates the resource selected in dropdown

            navigate_to_schedule(driver)
            events = parse_reservations(driver)

            safe_filename = name.replace(" ", "_") + ".ics"
            full_path = os.path.join(REPO_PATH, safe_filename)

            generate_ics(events, filename=full_path)

            subprocess.run(["git", "-C", REPO_PATH, "add", safe_filename])
            print(f"✅ Added {safe_filename} to Git")

        # Only commit & push once, after all files are added
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_message = f"Update multiple schedules ({now})"
        subprocess.run(["git", "-C", REPO_PATH, "commit", "-m", commit_message])
        subprocess.run(["git", "-C", REPO_PATH, "push"])
        print("🚀 Git push completed.")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
