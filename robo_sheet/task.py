""" An example robot. """

from RPA.Cloud.Google import Google
from RPA.Browser.Selenium import Browser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import time
from selenium.webdriver.support.select import Select



library = Google()
service_credentials = 'service_account.json'
library.init_sheets_client(service_credentials)
SHEET_ID = "14QTHDUS3Og3yR8bIwTZTR5FZ1MkUCLnj6-WF-1vBC8w"
NAME_SHEET_RANGE = "SHEET1!A2:A16"
DATE_OF_BIRTHS_SHEET_RANGE = "SHEET1!B2:B16"

browser = Browser()
url = "https://en.wikipedia.org/wiki/Wiki"


def open_the_website(url: str):
    browser.open_available_browser(url)


def search_for(term: str):
    input_field = "css:#searchInput"
    browser.input_text(input_field, term)
    browser.press_keys(input_field, "ENTER")


def get_date_of_birth_and_url(name):
    search_for(name)
    for e in browser.driver.find_elements_by_class_name("bday"):
        return e.get_attribute("innerText"), browser.driver.current_url
        
    return None, None


def get_names_from_sheet():
    spreadsheet_content = library.get_values(
        sheet_id=SHEET_ID, sheet_range=NAME_SHEET_RANGE)
    values = spreadsheet_content['values']
    return [name[0] for name in values]


def write_date_of_birth_to_sheet(idols_info):
    library.insert_values(sheet_id=SHEET_ID, 
        sheet_range=DATE_OF_BIRTHS_SHEET_RANGE, 
        values=[info['birthday'] for info in idols_info], 
        major_dimension="ROWS")


def create_calendar_entry(idols_info):
    open_the_website("https://calendar.google.com")
    browser.driver.implicitly_wait(3)
    browser.input_text("css:#identifierId", "sudo.quang.12@gmail.com")
    browser.press_keys("css:#identifierId", "ENTER")
    browser.driver.implicitly_wait(3)
    password_textbox = browser.driver.find_element_by_name("password")
    password_textbox.send_keys("Zcaffee@1012")
    password_textbox.send_keys(Keys.ENTER)
    browser.driver.implicitly_wait(3)
    time.sleep(5)

    for idol_info in idols_info:
        browser.driver.get("https://calendar.google.com/calendar/u/0/r/eventedit?tab=mc")
        title_txtbox = browser.driver.find_element_by_css_selector("input[aria-label='Title']")
        title_txtbox.send_keys(idol_info['name'] + " Birthday!")

        recurrence_select = browser.driver.find_element_by_css_selector("div[aria-label='Recurrence']")
        recurrence_select.click()
        time.sleep(2)

        yearly_recurrence_select = recurrence_select.find_elements_by_css_selector("div[role='option']")[4]
        yearly_recurrence_select.click()
        time.sleep(2)

        description_txtbox = browser.driver.find_element_by_css_selector("div[aria-label='Description']")
        description_txtbox.send_keys("This is birthday of " + idol_info['name'])
        description_txtbox.send_keys(Keys.ENTER)
        description_txtbox.send_keys("Wikipage: " + idol_info['link'])
        time.sleep(5)

        start_date_txtbox = browser.driver.find_element_by_css_selector("input[aria-label='Start date']")
        time.sleep(1)
        start_date_txtbox.send_keys(Keys.BACKSPACE)
        time.sleep(1)
        start_date_txtbox.send_keys(idol_info['birthday'])
        start_date_txtbox.send_keys(Keys.ENTER)
        time.sleep(5)

        save = browser.driver.find_element_by_css_selector("div[aria-label='Save']")
        time.sleep(1)
        save.click()
        time.sleep(5)
        print("MOVE NEXT!!!!")


def main():
    
    try:
        open_the_website(url)
        names = get_names_from_sheet()
        idols_info = []
        for name in names:
            date_of_births, link = get_date_of_birth_and_url(name)
            idols_info.append({
                "name": name,
                "link": link,
                "birthday": date_of_births
            })

        write_date_of_birth_to_sheet(idols_info)
        print("idols_info: ")
        print(idols_info)
        create_calendar_entry(idols_info)
    
    except Exception as e:
        raise e  
    finally:
        print("DONE")
        import time
        print("xxxxx")
        time.sleep(5)
        browser.close_all_browsers()


if __name__ == "__main__":
    main()
