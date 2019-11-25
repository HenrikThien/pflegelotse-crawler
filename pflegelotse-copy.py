import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json 
import time
import sys

###
# Funktioniert nur mit der Ambulanten suche im Moment
# @author: Henrik Thien henrikthienth@gmail.com
# @site: github.com/HenrikThien/pflegelotse
###
def main():
    city = input("In welcher Stadt soll gesucht werden: ")

    citySearch = requests.get("https://geocoder.123map.de/pcplacestreet.json?thm=itsg-geoc1&limit=7&qb=" + city)

    citys = json.loads(citySearch.text)

    print("Es wird gesucht in : " + citys[0]['value'])

    city = citys[0]

    browser = webdriver.Chrome('../../chromedriver')
    browser.set_page_load_timeout(30)
    browser.get("https://pflegelotse.de")

    half_box = browser.find_element_by_class_name("box")
    einrichtung_suchen = half_box.find_elements_by_tag_name("button")[0]
    einrichtung_suchen.click()

    inputElement = browser.find_element_by_id('ctl00_ContentPlaceHolder1_suche_bezirk')
    inputElement.send_keys(city['value'])

    WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_suche_bezirk_listbox")))
    WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "tt-suggestion")))

    selection = browser.find_element_by_id('ctl00_ContentPlaceHolder1_suche_bezirk_listbox')
    first_entry = browser.find_element_by_class_name('tt-suggestion')

    first_entry.click()

    button = browser.find_element_by_id("ctl00_ContentPlaceHolder1_suche_btn_versorgung1")
    button.click()

    submit = browser.find_element_by_id('ctl00_ContentPlaceHolder1_suche_btn_suche')
    submit.click()

    stealing_process(browser)

    
def stealing_process(browser):
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "results_ambulant")))

    res_table = browser.find_element_by_id("results_ambulant")
    tbody = res_table.find_elements_by_tag_name("tbody")
    rows = tbody[0].find_elements_by_tag_name("tr")

    dienste = []

    loop_num = len(rows)

    for i in range(loop_num):
        row = rows[i]

        cols = rows[i].find_elements_by_tag_name("td")

        name = cols[0]
        ort = cols[2]
        strasse = cols[3]
        telefon = cols[4]

        print(name.text)
        #print(ort.text);
        #print(strasse.text);

        if telefon.text != "--":
            print(telefon.text)

        row.click()
        
        try:
            WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_aZurueck")))

            email = ""
            website = ""
            address = ""

            try:
                address = browser.find_element_by_id("ctl00_ContentPlaceHolder1_p_adresse_header").text
                print(address)
            except:
                address = ""

            try:
                email = browser.find_element_by_id("ctl00_ContentPlaceHolder1_a_mail_header").text
                print(email)
            except:
                email = ""

            try:
                website = browser.find_element_by_id("ctl00_ContentPlaceHolder1_a_webseite_header").text
                print(website)
            except:
                website = ""

            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_spanZurueck")))
            back_button = browser.find_element_by_id("ctl00_ContentPlaceHolder1_spanZurueck")
            back_button.click()
        
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "results_ambulant")))
            res_table = browser.find_element_by_id("results_ambulant")
            tbody = res_table.find_elements_by_tag_name("tbody")
            rows = tbody[0].find_elements_by_tag_name("tr")
            print()
        except:
            e = sys.exc_info()[0]
            print(e)

    try:
        next_page = browser.find_element_by_id("ctl00_ContentPlaceHolder1_lvDesktopAmbulantPager_ctl02_NextButton")
        next_page.click()
        stealing_process(browser) 
    except:
        browser.quit()


if __name__ == '__main__':
    main()
