from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pickle

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from io import BytesIO
import requests
from docx.shared import Inches
import time

# global variables
admins = [] # redacted


def login_with_cookie(date):
    url = '''redacted''' + date
    # setup
    options = Options()
    options.add_argument("--disable-notifications")
    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome(service=Service())
    driver.get(url)
    set_cookie(driver)
    doc = Document()

    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CLASS_NAME, "pageNum"))
        )

        page_number = int(driver.find_element(By.CLASS_NAME, "pageNum").text.split("/")[1])
        for i in range(page_number):
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CLASS_NAME, "message-group"))
            )
            authors = driver.find_elements(By.CLASS_NAME, "nickName")
            content_wrappers = driver.find_elements(By.CLASS_NAME, "ant-dropdown-trigger")
            times = driver.find_elements(By.CLASS_NAME, "time-style")

            for author, content_wrapper, curr_time in zip(authors, content_wrappers, times):
                try:
                    content = content_wrapper.find_element(By.TAG_NAME, "p").text
                    render_file(doc, author.text, curr_time.text, content, date)
                except:
                    img_link = content_wrapper.find_element(By.TAG_NAME, "img").get_attribute("src")
                    add_image(doc, author.text, curr_time.text, img_link, date)
            click = driver.find_element(By.CLASS_NAME, "el-icon-arrow-right")
            driver.execute_script('arguments[0].click();', click)

            time.sleep(1)
    except Exception as e:
        # no cookie found, get cookie
        print("error:")
        print(e)
        driver.quit()
    finally:
        print("all done")
        driver.quit()


def login_without_cookie():
    url = "" #redacted
    # setup
    options = Options()
    options.add_argument("--disable-notifications")
    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome(service=Service())
    driver.get(url)
    phone_number = ""  # redacted
    enter_phone = driver.find_element(By.ID, "one")
    enter_phone.send_keys(phone_number)

    get_verification_code = driver.find_element(By.ID, "getCode")
    get_verification_code.click()
    verification_code = input("Enter validation code: \n")

    enter_code = driver.find_element(By.ID, "two")
    enter_code.send_keys(verification_code)

    login = driver.find_element(By.CLASS_NAME, "login")
    login.click()

    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "select_list"))
        )
        get_cookie(driver)

    except Exception as e:
        print(e)
        driver.quit()


def set_cookie(driver):
    cookies = pickle.load(open("cookies.pkl", "rb"))  # load_cookie
    for cookie in cookies:
        driver.add_cookie(cookie)


def get_cookie(driver):
    pickle.dump(driver.get_cookies(), open("cookies.pkl", "wb"))  # store_cookie


def render_file(doc, name, time, content, doc_name):
    # credit: https://stackoverflow.com/questions/61801936/set-background-color-shading-on-normal-text-with-python-docx
    if name in admins:
        color = 'FF7276'
    else:
        color = 'd3d3d3'
    doc.add_paragraph(name + "            " + time)
    p = doc.add_paragraph()
    txt = content
    run = p.add_run(txt)
    tag = run._r
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    run.font.size = Pt(11)
    tag.rPr.append(shd)
    doc.save(doc_name+".docx")


def add_image(doc, name, time, img_link, doc_name):
    # https://stackoverflow.com/questions/24341589/python-docx-add-picture-from-the-web
    doc.add_paragraph(name + "            " + time)
    response = requests.get(img_link)
    binary_img = BytesIO(response.content)
    doc.add_picture(binary_img, width=Inches(2))
    doc.save(doc_name+'.docx')


def main():
    dates = input("Please enter the date, Example：20220401. You can enter multiple dates，separated by spaces，example：20220401 20220402 20220403\n").split()
    for date in dates:
        login_with_cookie(date)


if __name__ == '__main__':
    main()



