# coding=utf-8
import json
import argparse
import urllib
from datetime import timedelta, datetime, date

import os
from zipfile import ZipFile

from exchangelib import DELEGATE, Account, Credentials, EWSDateTime, EWSTimeZone, CalendarItem
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

loginURL = 'https://wayf.cesi.fr/login?client_name=ClientIdpViaCesiFr&needs_client_redirection=true'
apiURL = 'https://ent.cesi.fr/api/seance/all?start='
calendarURL = 'https://ent.cesi.fr/mon-compte/mon-emploi-du-temps/'
path2phantom = "phantomjs.exe"

parser = argparse.ArgumentParser(description="Importe le calendrier de l'ENT sur Microsoft Exchange")
parser.add_argument("username", nargs=1, help="Email utilise pour l'ENT")
parser.add_argument("password", nargs=1, help="Mot de passe utilise pour l'ENT")
parser.add_argument("-s", dest="semaine", default=[1], type=int, nargs=1, help="Nombre de semaine a importer")
parser.add_argument("--blank", dest="blank", action="store_true", help="Lance le programme sans ajouter les evenements au calendrier")
parser.add_argument("--rollback", dest="rollback", action="store_true", help="Supprime tous les evenements ajoutes par ce programme")
parser.add_argument("--folder", dest="folder", nargs=1, default="calendar", help="Choisis un dossier différent pour le calendrier, ce dossier doit avoir ete cree au prealable")
args = parser.parse_args()
semaine = args.semaine[0]
username = args.username[0]
password = args.password[0]
blank = args.blank

if not os.path.isfile(path2phantom):
    print "Telechargement de phantomjs en cours ..."
    urllib.urlretrieve("https://bitbucket.org/ariya/phantomjs/downloads/phantomjs-2.1.1-windows.zip", "tmp.zip")
    zip = ZipFile('tmp.zip')
    zip.extract("phantomjs-2.1.1-windows/bin/phantomjs.exe")
    zip.close()
    os.rename("phantomjs-2.1.1-windows/bin/phantomjs.exe", "phantomjs.exe")
    os.remove("tmp.zip")
    os.rmdir("phantomjs-2.1.1-windows/bin")
    os.rmdir("phantomjs-2.1.1-windows")

def waitForID(browser, id):
    try:
        WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.ID, id)))
    except TimeoutException:
        print "Erreur de connexion à l'ENT"
        browser.quit()
        exit(0)


def waitForClass(browser, CLASS):
    try:
        WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, CLASS)))
    except TimeoutException:
        print "Erreur de connexion à l'ENT"
        browser.quit()
        exit(0)


def getUserID(browser):
    browser.get(calendarURL)
    waitForClass(browser, "fc-widget-header")
    if "data-code-personne-courante" in browser.page_source:
        return browser.page_source.split("data-code-personne-courante=")[1].split('"')[1]
    else:
        print "Impossible de recuperer l'ID utilisateur"
        browser.quit()
        exit(0)


def getStartDate(browser):
    browser.get(calendarURL)
    waitForClass(browser, "fc-widget-header")
    if "data-date" in browser.page_source:
        return browser.page_source.split("data-date=")[1].split('"')[1]
    else:
        print "Impossible de recuperer l'ID utilisateur"
        browser.quit()
        exit(0)


def getEndDate(browser):
    browser.get(calendarURL)
    waitForClass(browser, "fc-widget-header")
    if "data-date" in browser.page_source:
        return browser.page_source.split("data-date=")[6].split('"')[1]
    else:
        print "Impossible de recuperer l'ID utilisateur"
        browser.quit()
        exit(0)

def connectENT(browser):
    browser.get(loginURL)
    waitForID(browser, "userNameArea")

    browser.find_element_by_id("userNameInput").send_keys(username)
    browser.find_element_by_css_selector("input[type='password']").send_keys(password)
    browser.find_element_by_id("submitButton").click()
    waitForID(browser, "msg")


def apiCall(browser, start, end, userid):
    browser.get("about:blank")
    browser.get(apiURL + start + "&end=" + end + "&codePersonne=" + userid)
    if browser.page_source != "<html><head></head><body></body></html>":
        data = browser.page_source.split('<pre style="word-wrap: break-word; white-space: pre-wrap;">')[1].split("</pre>")[0]
        return json.loads(data)
    else:
        return ""


def toDate(str):
    tmp = map(int, str.split('T')[0].split('-'))
    return date(tmp[0], tmp[1], tmp[2])


def toDateTime(date):
    tmp = date.split('T')
    date = map(int, tmp[0].split('-'))
    hours = map(int, tmp[1].split("+")[0].split(":"))
    return datetime(date[0], date[1], date[2], hours[0], hours[1], hours[2])


def dateToString(date):
    return str(date.year) + "-" + str(date.month) + "-" + str(date.day)


if semaine < 1:
    semaine = 1

tz = EWSTimeZone.localzone()
credentials = Credentials(username=username, password=password)
account = Account(primary_smtp_address=username, credentials=credentials, autodiscover=True, access_type=DELEGATE)

folder = account.calendar
if not args.folder == "calendar":
    for f in account.root.walk():
        if f.name == args.folder[0]:
            folder = f
    if folder == account.calendar:
        print "Ce calendrier n'existe pas, verifiez que le nom entre correspond a un calendrier existant"
        exit(0)

if args.rollback:
    items = folder.filter(categories="Bot Calendrier")
    counter = 0
    for item in items:
        if not blank:
            item.delete()
        counter += 1
    if counter > 1:
        print str(counter) + " evenements supprimes."
    else:
        print str(counter) + " evenement supprime."
    exit(0)

browser = webdriver.PhantomJS(path2phantom)

connectENT(browser)

id = getUserID(browser)
start = toDate(getStartDate(browser))
end = toDate(getEndDate(browser)) + timedelta(days=1)

counter = 0

for y in range(0, semaine):
    if y > 0:
        start += timedelta(days=7)
        end += timedelta(days=7)
    json_data = apiCall(browser, dateToString(start), dateToString(end), id)
    if json_data != "":
        for x in json_data:
            startDate = toDateTime(x['start'])
            startEWS = tz.localize(EWSDateTime(startDate.year, startDate.month, startDate.day, startDate.hour, startDate.minute))

            endDate = toDateTime(x['end'])
            endEWS = tz.localize(EWSDateTime(endDate.year, endDate.month, endDate.day, endDate.hour, endDate.minute))

            inters = ""
            salles = ""
            if x['salles']:
                for salle in x['salles']:
                    if " " in salle['nomSalle']:
                        salle['nomSalle'] = salle['nomSalle'].split(" ")[0]
                    if "-" in salle['nomSalle']:
                        salle['nomSalle'] = salle['nomSalle'].split("-")[0]
                    salles += salle['nomSalle'] + " - "
                salles = salles[:-3]

            if x['intervenants']:
                for inter in x['intervenants']:
                    inters += inter['nom'] + " " + inter['prenom']
                    if inter['adresseMail'] != "":
                        inters += " - " + inter['adresseMail']
                    inters += "\n"
            if not account.calendar.filter(start__gte=startEWS, end__lte=endEWS, subject=x['title']):
                counter += 1
                if not blank:
                    item = CalendarItem(folder=folder, categories=["Bot Calendrier"], subject=x['title'], start=startEWS, end=endEWS, body=inters, location=salles)
                    item.save()

if counter > 1:
    print str(counter) + " evenements ajoutes."
else:
    print str(counter) + " evenement ajoute."

browser.quit()
