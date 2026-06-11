from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time
import re
import datetime

BASE_URL = "https://www.metalinfo.ru"
START_URL = BASE_URL + "/ru/directory"


options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")


def get_driver():

    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def get_company_links(driver):
    driver.get(START_URL)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    links = []
    for a in soup.select("div.col-sm-13.directory-last-added a[href^='/ru/directory/']"):
        href = a.get("href")
        if href and not href.endswith("/ru/directory/"):
            links.append(BASE_URL + href.strip())
    return links


def extract_email(text):

    emails = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', text)
    for e in emails:
        if not e.lower().endswith(("metalinfo.ru", "example.com")):
            return e
    return ""


def get_text_after_label(soup, label_text):

    tag = soup.find("dt", string=re.compile(label_text))
    if tag:
        dd = tag.find_next_sibling("dd")
        if dd:
            return dd.get_text(" ", strip=True)
    return ""


def parse_company_page(driver, url):
    driver.get(url)
    time.sleep(1.5)
    soup = BeautifulSoup(driver.page_source, "html.parser")

    name = soup.find("h1").get_text(strip=True) if soup.find("h1") else ""

    tel_tag = soup.find("dd", itemprop="telephone")
    phone = tel_tag.get_text(strip=True) if tel_tag else ""

    email = extract_email(driver.page_source)

    site = ""
    for a in soup.find_all("a", href=re.compile(r"^http")):
        href = a["href"]
        if "metalinfo.ru" not in href:
            site = href
            break

    addr_tag = soup.find("dd", itemprop="address")
    address = addr_tag.get_text(" ", strip=True) if addr_tag else ""

    boss = get_text_after_label(soup, "Руководитель компании")
    description = soup.find("dd", itemprop="description")
    description = description.get_text(" ", strip=True) if description else ""

    products = get_text_after_label(soup, "Продукция")
    update_date = get_text_after_label(soup, "Дата обновления")
    views = get_text_after_label(soup, "Количество просмотров")

    return {
        "Название": name,
        "Телефон": phone,
        "E-mail": email,
        "Сайт": site,
        "Адрес": address,
        "Руководитель": boss,
        "Сфера деятельности": description,
        "Продукция": products,
        "Дата обновления": update_date,
        "Количество просмотров": views,
        "URL страницы": url
    }


def save_to_excel(data_list, filename="metalinfo_contacts.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Компании"
    headers = [
        "Название", "Телефон", "E-mail", "Сайт", "Адрес", "Руководитель",
        "Сфера деятельности", "Продукция", "Дата обновления",
        "Количество просмотров", "URL страницы"
    ]
    ws.append(headers)
    for item in data_list:
        ws.append([item.get(h, "") for h in headers])
    wb.save(filename)
    print(f"✅ Сохранено в {filename}")


def scrape_once():
    driver = get_driver()
    try:
        links = get_company_links(driver)
        print(f"Найдено {len(links)} компаний")

        all_data = []
        for i, link in enumerate(links, 1):
            print(f"[{i}/{len(links)}] {link}")
            try:
                info = parse_company_page(driver, link)
                print(f"   → {info['E-mail'] or 'без e-mail'}")
                all_data.append(info)
            except Exception as e:
                print(f"⚠️ Ошибка {link}: {e}")

        save_to_excel(all_data)
    finally:
        driver.quit()


def main():
    while True:
        print(f"\n🕒 Запуск парсинга: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        scrape_once()
        print("💤 Сон на 24 часа...")
        time.sleep(86400)


if __name__ == "__main__":
    main()