import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import requests, re, time
from Reordered_List import list as list_data
from api_token_key import key as api_key
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

name = phone = house = None
street = '---'
flat = '-'
region = district = city = None

street_keys = ['вул', 'в ул', 'ву л', 'в у л', 'бульвар', 'проспект', 'набережна', 'ул', 'вулиця', 'улица']
valid_codes = ['39','68','96','97','98','66','50','95','99','63','67','93','91','92','94','62','73','77']

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SAMPLE_SPREADSHEET_ID = "1v4-sFcPHpYlvGhlkq_hddcRBuIQtfbbLKR4WonaJ3XM"

def sheets_system(SCOPES, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME):
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=SAMPLE_RANGE_NAME
        ).execute()
        values = result.get("values", [])
    except HttpError as err:
        print(err)
    return values

def system_processing_address(address, api_key):
    url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={api_key}&language=uk"
    response = requests.get(url)
    data = response.json()

    #pprint(data)
    region = district = city = None

    if data["status"] == "OK":
        address_words = set(re.sub(r'[^\w\s]', '', address.lower()).split())

        best_match = None
        best_score = 0

        for result in data["results"]:
            components = result["address_components"]
            component_names = [c["long_name"].lower() for c in components]

            score = sum(1 for word in address_words if any(word in comp for comp in component_names))
            if score > best_score:
                best_score = score
                best_match = components

        if not best_match:
            best_match = data["results"][0]["address_components"]

        for c in best_match:
            if "administrative_area_level_1" in c["types"]:
                region = c["long_name"].replace(" область", "").strip()
            elif "administrative_area_level_2" in c["types"] or "administrative_area_level_3" in c["types"]:
                district = c["long_name"].replace(" район", "").strip()
            elif any(t in c["types"] for t in ["locality", "sublocality", "postal_town"]):
                city = c["long_name"].strip()

    if region in list_data:
        found = False
        for d, cities in list_data[region].items():
            for item in cities:
                item_clean = re.sub(r"^(м\.|с\.|смт)\s+", "", item.lower())
                city_clean = city.lower() if city else ""
                if city_clean == item_clean or item.lower() in address.lower():
                    district = d
                    city = item
                    found = True
                    break
            if found:
                break

        if not found:
            district = next(iter(list_data[region]))
            city = list_data[region][district][0]
    return region, district, city

def street_check(line):
    line = line.strip().lower()
    if not line:
        return False
    elif '@' in line or re.search('\.com|\.ua|\.net|\.ru|\.mail', line):
        return False
    elif re.search(r'\d{8,}', line):
        return False
    elif len(line.split()) >= 4 and not re.search('\d', line):
        return False
    else: return True

def system_processing_streets(address_street, street_keys):
    street = '---'
    house = '-'
    flat = False
    street_iteration = False

    match = re.search('\d', address_street)
    if match and '@' not in address_street:
        result_number = re.findall('\d+/\d+|\d+\s/\s\d+|\d+[а-яА-Я]|\d+-[а-дА-Д]|\d+', address_street)
        house = result_number[0]
        if len(result_number) >= 2:
            flat = result_number[1]

    for key in street_keys:
        words_in_address = re.findall('[^0-9.,\s]+', address_street.lower())
        if key in words_in_address:
            edited_street = re.findall('[^0-9.,\s]+', address_street.lower())
            for word in edited_street:
                if len(word) >= 3 and word != key and edited_street.index(key) < edited_street.index(word):
                    if not street_iteration:
                        result_street = word
                        street_iteration = True
                    else:
                        if word not in ['буд', 'кв', 'квартира', 'будинок']:
                            result_street += ' ' + word
            street = result_street
            break

    if street == '---' and street_check(address_street):
        edited_street = re.findall('[^0-9.,\s]+', address_street.lower())
        for word in edited_street:
            if len(word) >= 3:
                if not street_iteration:
                    result_street = word
                    street_iteration = True
                else:
                    if word not in ['буд', 'кв', 'квартира', 'будинок']:
                        result_street += ' ' + word
        street = result_street
    return street, house, flat

def system_processing_phone(phone):
    phone = (re.findall('\d+', phone))[0]
    first_three = phone[:3]
    if first_three == '380':
        phone = re.sub('380', '', phone, count=1)
    if phone[:1] == '0' and len(phone) == 10:
        phone = re.sub('0', '', phone, count=1)
    next_first_two = phone[:2]
    if next_first_two not in valid_codes or len(phone) != 9:
        phone = None
    return phone

def send_forms_to_website(name, phone, house, street, flat, region, district, city):
    wait.until(EC.presence_of_element_located((By.ID, "inp_client_name"))).send_keys(name)
    wait.until(EC.presence_of_element_located((By.ID, "inp_client_phone"))).send_keys(phone)
    wait.until(EC.presence_of_element_located((By.ID, "inp_client_street"))).send_keys(street)
    wait.until(EC.presence_of_element_located((By.ID, "inp_client_house"))).send_keys(house)
    if flat:
        wait.until(EC.presence_of_element_located((By.ID, "inp_client_flat"))).send_keys(flat)
    region_input = wait.until(EC.element_to_be_clickable((By.ID, "inp_client_region")))
    region_input.click()
    if city == "м. Київ":
        elements = driver.find_elements(By.CLASS_NAME, "item")
        for el in elements:
            if 'Київська' in el.text:
                el.click()
                break
        elements = driver.find_elements(By.CLASS_NAME, "item")
        for el in elements:
            if "м. Київ" in el.text:
                el.click()
                break
    else:
        elements = driver.find_elements(By.CLASS_NAME, "item")
        for el in elements:
            if region in el.text:
                el.click()
                break
        elements = driver.find_elements(By.CLASS_NAME, "item")
        for el in elements:
            if district in el.text:
                el.click()
                break
        elements = driver.find_elements(By.CLASS_NAME, "item")
        for el in elements:
            if city in el.text:
                el.click()
                break

with open("last_sheet.txt", "r+") as file:
    last_sheet = file.read().strip()
    if not last_sheet:
        start_sheet = input("Введіть значення start_sheet: ")
        file.truncate(0)
        file.write(start_sheet)
    else:
        while True:
            test = input(f"Продовжуємо з попереднього місця(так/ні)\nОстання заявка (№{last_sheet}): ")
            if test.lower() == 'так':
                start_sheet = str(int(last_sheet) + 1)
                break
            elif test.lower() == 'ні':
                start_sheet = input("Введіть значення start_sheet: ")
                file.truncate(0)
                file.write(start_sheet)
                break
            else:
                print("Не правильний ввід, повторіть спробу")
SAMPLE_RANGE_NAME = f"Лист1!A{start_sheet}:E{int(start_sheet)+25}"

driver = webdriver.Chrome()
driver.get("https://optica.ukrtelecom.ua/?utm_source=facebook_fttb&utm_medium=cpc")
wait = WebDriverWait(driver, 10)

values = sheets_system(SCOPES, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME)
processed_rows = 0
for row in values:
    if not any(row):
        break
    try:
        name = row[2]
        phone = system_processing_phone(row[3])
        address_region = row[0]
        address_street = row[1]

        street, house, flat = system_processing_streets(address_street, street_keys)
        region, district, city = system_processing_address(address_region, api_key)
        if phone != None:
            send_forms_to_website(name, phone, house, street, flat, region, district, city)
            time.sleep(10)
        processed_rows += 1

    except Exception as e:
        print(f"Помилка при обробці рядка: {row} — {e}")
        with open("last_sheet.txt", "w") as file:
            file.write(row)
        continue
