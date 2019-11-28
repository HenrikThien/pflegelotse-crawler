import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import json 
import time
import sys
import xlsxwriter
import traceback
from datetime import date
import chromedriver_binary
import os,sys

###
# Funktioniert nur mit der Ambulanten suche im Moment
# @author: Henrik Thien henrikthienth@gmail.com
# @site: github.com/HenrikThien/pflegelotse
###
dienste_list = []

def main():
    city = input("In welcher Stadt soll gesucht werden: ")

    citySearch = requests.get("https://geocoder.123map.de/pcplacestreet.json?thm=itsg-geoc1&limit=7&qb=" + city, verify=False)

    citys = json.loads(citySearch.text)

    print("Es wird gesucht in : " + citys[0]['value'])

    city = citys[0]

    versorgungsform = input("Welche Versorgungsform wird gesucht? (ambulant = a, stationär = s): ")
    pa_py = versorgungsform

    pflegeart = "x"

    if pa_py == "s":
        pflegeart = input("Welche Pflegeart? (v)ollstationäre, (t)ages, (n)acht, (k)urzzeit: ")


    kilometer = input("Suche im Umkreis von (5, 10, 15, 25, 50) km: ")

    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
    DRIVER_BIN = os.path.join(PROJECT_ROOT, "chromedriver")

    chrome_options = Options()  
    chrome_options.add_argument("--headless")

    #options=chrome_options
    browser = webdriver.Chrome(executable_path=DRIVER_BIN,options=chrome_options)
    browser.set_page_load_timeout(30)
    browser.get("https://pflegelotse.de")

    half_box = browser.find_element_by_class_name("box")
    einrichtung_suchen = half_box.find_elements_by_tag_name("button")[0]
    einrichtung_suchen.click()

    searchFieldEntry(browser, 'ctl00_ContentPlaceHolder1_suche_bezirk', 'ctl00_ContentPlaceHolder1_suche_bezirk_listbox', city)

    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_suche_umkreis")))
    select = Select(browser.find_element_by_id('ctl00_ContentPlaceHolder1_suche_umkreis'))
    select.select_by_value(kilometer)

    browser.find_element_by_id("ctl00_ContentPlaceHolder1_suche_btn_versorgung1").click()
    
    if pa_py == "s":
        browser.find_element_by_id("ctl00_ContentPlaceHolder1_suche_btn_versorgung2").click()
        selectStationaerField(browser, pflegeart)

    submit = browser.find_element_by_id('ctl00_ContentPlaceHolder1_suche_btn_suche')
    submit.click()

    if pa_py == "a":
        stealing_process_ambulant(browser)
    elif pa_py == "s":
        stealing_process_stationaer(browser, pflegeart)
    else:
        print("Für diese Vorsorgeform gibt es noch keine Funktion..")

    # excel datei am ende erstellen
    print("excel datei wird erstellt, bitte warten...")

    type = "ambulant"
    if pa_py == "s":
        type = "stationaer"

    create_excel_file(type, city['value'])
    
    try:
        browser.quit()
    except:
        print()

    print("Programm beendet.")
    os.system("pause")

def selectStationaerField(browser, pflegeart):
    WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.ID, "stationaer")))
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "nophone")))
    time.sleep(0.5)

    if pflegeart == "t":
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_suche_btn_pflegeart2")))
        browser.find_element_by_id("ctl00_ContentPlaceHolder1_suche_btn_pflegeart2").click()
    elif pflegeart == "n":
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_suche_btn_pflegeart3")))
        browser.find_element_by_id("ctl00_ContentPlaceHolder1_suche_btn_pflegeart3").click()
    elif pflegeart == "k":
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_suche_btn_pflegeart4")))
        browser.find_element_by_id("ctl00_ContentPlaceHolder1_suche_btn_pflegeart4").click()
    else:
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_suche_btn_pflegeart1")))
        browser.find_element_by_id("ctl00_ContentPlaceHolder1_suche_btn_pflegeart1").click()

def searchFieldEntry(browser, inputName, listboxName, city):
    inputElement = browser.find_element_by_id(inputName)
    inputElement.send_keys(city['value'])

    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, listboxName)))
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "tt-suggestion")))

    selection = browser.find_element_by_id(listboxName)

    entries = browser.find_elements_by_class_name('tt-suggestion')

    for i in entries:
        if city['value'] == i.text:
            i.click()

def stealing_process_stationaer(browser, pflegeart):
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "results_vollstationaer")))

    res_table = browser.find_element_by_id("results_vollstationaer")
    tbody = res_table.find_elements_by_tag_name("tbody")
    rows = tbody[0].find_elements_by_tag_name("tr")
    
    dienste = []

    loop_num = len(rows)

    for i in range(loop_num):
        row = rows[i]

        try:
            name = ""
            telefon = ""

            try:
                cols = rows[i].find_elements_by_tag_name("td")

                name = cols[0].text
                atrs = name.splitlines()
                name = atrs[0]
                tel = atrs[3].split(":")[1].strip()

            except:
                print("td wurde nicht gefunden!?")

            if tel == "--":
                tel = ""

            try:
                row.click()           
                fetch_infos(browser, name, tel)
            except:
                link = browser.find_element_by_id("ctl00_ContentPlaceHolder1_lvDesktopVollstationaer_ctrl%d_DetailButton" % i)
                browser.execute_script("arguments[0].click();", link)
                fetch_infos(browser, name, tel)

            WebDriverWait(browser, 40).until(EC.presence_of_element_located((By.ID, "results_vollstationaer")))

            res_table = browser.find_element_by_id("results_vollstationaer")
            tbody = res_table.find_elements_by_tag_name("tbody")
            rows = tbody[0].find_elements_by_tag_name("tr")

        except Exception as e: 
            print(e)

    try:
        back_button = ""

        if pflegeart == "v":
            back_button = "ctl00_ContentPlaceHolder1_lvDesktopVollstationaerPager_ctl02_NextButton"
        else:
            back_button = "ctl00_ContentPlaceHolder1_lvDesktopTeilstationaerPager_ctl02_NextButton"

        next_page = browser.find_element_by_id(back_button)
        disabled = attribute_exists(next_page, "disabled")
        if disabled != True:
            next_page.click()
            stealing_process_stationaer(browser, pflegeart)
        else:
            browser.quit()
    except Exception as e:
        browser.quit()

def attribute_exists(element, attribute):
    result = False
    try:
        value = element.get_attribute(attribute)
        if value is not None:
            result = True
        else:
            result = False
    except:
        result = False
    return result 


def stealing_process_ambulant(browser):
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "results_ambulant")))

    res_table = browser.find_element_by_id("results_ambulant")
    tbody = res_table.find_elements_by_tag_name("tbody")
    rows = tbody[0].find_elements_by_tag_name("tr")
    cols = []

    loop_num = len(rows)

    for i in range(loop_num):
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "results_ambulant")))

        row = rows[i]

        try:
            row = rows[i]
            name = ""
            telefon = ""
            try:
                cols = rows[i].find_elements_by_tag_name("td")

                name = cols[0].text
                telefon = cols[4].text
            except:
                print("td wurde nicht gefunden!?")

            #print("found: %s" % name)

            if telefon == "--":
                telefon = ""

            #row.click()           
            #fetch_infos(browser, name, telefon)

            try:
                row.click()           
                fetch_infos(browser, name, telefon)
            except:
                link = browser.find_element_by_id("ctl00_ContentPlaceHolder1_lvDesktopAmbulant_ctrl%d_DetailButton" % i)
                browser.execute_script("arguments[0].click();", link)
                fetch_infos(browser, name, telefon)

            WebDriverWait(browser, 40).until(EC.presence_of_element_located((By.ID, "results_ambulant")))
            res_table = browser.find_element_by_id("results_ambulant")
            tbody = res_table.find_elements_by_tag_name("tbody")
            rows = tbody[0].find_elements_by_tag_name("tr")

        except Exception as e: 
            print(e)

    try:
        next_page = browser.find_element_by_id("ctl00_ContentPlaceHolder1_lvDesktopAmbulantPager_ctl02_NextButton")
        disabled = attribute_exists(next_page, "disabled")
        if disabled != True:
            next_page.click()
            stealing_process_ambulant(browser)
        else:
            browser.quit()
    except:
        browser.quit()


def fetch_infos(browser, name, telefon):
    WebDriverWait(browser, 40).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_spanZurueck")))

    email = ""
    website = ""
    address = ""

    strasse = ""
    plz = ""
    ort = ""

    try:
        address = browser.find_element_by_id("ctl00_ContentPlaceHolder1_p_adresse_header").text
        adrLines = address.splitlines()
        strasse = adrLines[0]
        plz = adrLines[1].split()[0]
        ort = adrLines[1].split()[1]
    except:
        address = ""

    try:
        email = browser.find_element_by_id("ctl00_ContentPlaceHolder1_a_mail_header").text
    except:
        email = ""

    try:
        website = browser.find_element_by_id("ctl00_ContentPlaceHolder1_a_webseite_header").text
    except:
        website = ""

    print("fetched %s" % name)
            
    dienst = {
        "name": name,
        "tel": telefon,
        "strasse": strasse,
        "plz": plz,
        "ort": ort,
        "email": email,
        "web": website
    }
    x1 = json.dumps(dienst)
    obj = json.loads(x1)
    dienste_list.append(obj)

    WebDriverWait(browser, 40).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_spanZurueck")))
    back_button = browser.find_element_by_id("ctl00_ContentPlaceHolder1_spanZurueck")
    back_button.click()

def create_excel_file(type, ort):
    today = date.today()

    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
    DRIVER_BIN = os.path.join(PROJECT_ROOT, "pf-%s-%s-%s.xlsx" % (today,type,ort))

    save_dir = DRIVER_BIN
    print("Speichere die Excel Datei hier: %s" % save_dir)
    workbook = xlsxwriter.Workbook(save_dir)
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Nachname')
    worksheet.write('B1', 'Vorname')
    worksheet.write('C1', 'Anrede')
    worksheet.write('D1', 'Firma')
    worksheet.write('E1', '')
    worksheet.write('F1', 'Tel')
    worksheet.write('G1', 'Straße')
    worksheet.write('H1', 'PLZ,Ort')
    worksheet.write('I1', 'email')
    worksheet.write('J1', 'Web')
    worksheet.write('K1', 'Angemeldet')
    worksheet.write('L1', 'Forumteilnahme')

    row = 1

    for dienst in dienste_list:
        ortPlz = "%s, %s" % (dienst["plz"], dienst["ort"])
        worksheet.write(row, 3, dienst["name"])
        worksheet.write(row, 5, dienst["tel"])
        worksheet.write(row, 6, dienst["strasse"])
        worksheet.write(row, 7, ortPlz)
        worksheet.write(row, 8, dienst["email"])
        worksheet.write(row, 9, dienst["web"])
        row = row + 1


    workbook.close()
    

if __name__ == '__main__':
    main()