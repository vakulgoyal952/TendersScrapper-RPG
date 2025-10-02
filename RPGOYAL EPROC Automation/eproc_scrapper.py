import os
import sys
import datetime
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# ====== HELPER FUNCTION TO GET BASE PATH ======
def get_base_path():
    if getattr(sys, 'frozen', False):  # If bundled as .exe
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

base_path = get_base_path()

# ====== SETUP LOGGER WITH TIMESTAMPED FILE ======
log_dir = os.path.join(base_path, "logs")
os.makedirs(log_dir, exist_ok=True)

log_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_file = os.path.join(log_dir, f"tender_scraper_{log_time}.log")

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console.setFormatter(formatter)
logging.getLogger().addHandler(console)

logging.info("Logger initialized. Starting tender scraping process.")

# ====== DRIVER INITIALIZATION ======
def init_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    options.page_load_strategy = 'eager'
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

# ====== NAVIGATE TO TENDER LISTING ======
def navigate_to_main_page(driver):
    driver.get("https://eproc.rajasthan.gov.in/")
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "alertbutclose"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PageLink_10"))).click()
    logging.info("Navigated to active tenders page.")

# ====== FETCH ORGANISATIONS AND LINKS ======
def get_tenders(driver):
    return driver.find_elements(By.XPATH, "//a[contains(@id,'DirectLink')]")

# ====== FETCH INDIVIDUAL FIELD FROM TENDER PAGE ======
def get_field_text_from_popup(driver, field_name):
    try:
        xpath = f"//td/b[normalize-space()='{field_name}']/ancestor::td/following-sibling::td[1]"
        element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return element.text.strip()
    except Exception as e:
        logging.warning(f"Could not extract '{field_name}': {e}")
        return "Not Found"

# ====== EXTRACT FIELDS FOR ONE TENDER ======
def get_org_tenders(org_name, org_id):
    driver = init_driver()
    navigate_to_main_page(driver)

    try:
        org_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, org_id)))
        org_element.click()
        logging.info(f"Selected organisation: {org_name}")
    except Exception as e:
        logging.error(f"Failed to click organisation {org_name}: {e}")
        driver.quit()
        return []

    tender_links = [t.get_attribute("href") for t in get_tenders(driver)]
    total = len(tender_links)
    logging.info(f"Found {total} tenders for organisation: {org_name}")

    fields_to_extract = [
        "Title", "Work Description", "Tender Value in ₹",
        "Location", "Organisation Chain", "Bid Submission End Date"
    ]

    tenders = []
    for i, link in enumerate(tender_links, start=1):
        try:
            driver.get(link)
            tender_data = {
                'Organisation': org_name,
                'Tender Link': link
            }
            for field in fields_to_extract:
                tender_data[field] = get_field_text_from_popup(driver, field)

            tenders.append(tender_data)
            remaining = total - i
            percent = round((i / total) * 100, 2)
            logging.info(f"[{i}/{total}] - Recorded tender: {tender_data['Title']} | Remaining: {remaining} | {percent}% done")

        except Exception as e:
            logging.error(f"Error processing tender {i}/{total}: {e}")
            continue

    driver.quit()
    logging.info(f"Completed scraping for organisation: {org_name}")
    return tenders

def extract_orgs_and_anchor_ids_selenium(driver):
    org_dict = {}
    table = driver.find_element(By.ID, "table")
    rows = table.find_elements(By.TAG_NAME, "tr")
    
    for row in rows:
        classes = row.get_attribute("class")
        if classes and "list_header" in classes:
            continue  # skip header

        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) >= 3:
            org_name = cells[1].text.strip()
            try:
                anchor = cells[2].find_element(By.TAG_NAME, "a")
                anchor_id = anchor.get_attribute("id")
                if anchor_id and anchor_id.startswith("DirectLink"):
                    org_dict[org_name] = anchor_id
            except NoSuchElementException:
                pass

    return org_dict

# ====== MAIN EXECUTION ======
if __name__ == "__main__":
    driver = init_driver()
    navigate_to_main_page(driver)
    org_dict = extract_orgs_and_anchor_ids_selenium(driver)
    for org_name, org_id in org_dict.items():
        tenders = get_org_tenders(org_name, org_id)

        if not tenders:
            logging.warning("No tenders scraped. Exiting.")
            continue

        df = pd.DataFrame(tenders)

        df = df.rename(columns={
            'Title': 'Title',
            'Organisation Chain': 'Organisation Chain',
            'Bid Submission End Date': 'Closing Date',
            'Tender Value in ₹': 'Amount',
            'Work Description': 'Work Description',
            'Location': 'Location',
        })
        df = df[['Title', 'Tender Link', 'Organisation Chain', 'Closing Date', 'Amount', 'Work Description', 'Location']]

        def make_hyperlink(title, url):
            return f'=HYPERLINK("{url}", "{title}")'

        df['Title'] = [make_hyperlink(t, u) for t, u in zip(df['Title'], df['Tender Link'])]
        df = df[['Title', 'Organisation Chain', 'Closing Date', 'Amount', 'Work Description', 'Location']]

        date = datetime.datetime.now().strftime("%Y-%m-%d")
        folder_name = os.path.join(base_path, f'tenders_{date}')
        filename = f'{org_name}_tenders_{date}.xlsx'

        os.makedirs(folder_name, exist_ok=True)

        df.to_excel(os.path.join(folder_name, filename), index=True, index_label="Sno")

        logging.info(f"Tenders exported to Excel: {filename}")
