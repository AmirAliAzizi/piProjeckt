from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from pick import pick
from argparse import ArgumentParser
from pynput.mouse import Controller
import time
import requests
import json

Maincounter = 0  # Maincounter
max_length_Resultlist = 0  # max length Resultlist
xtime = 0  # Changetime (display duration of one Modulwebsite)
zusatzseiten_counter = 0  # extra Website request (after x Modulwebsite one extra Website)
webseiten_list = []  # Modullist (Mainwebsite)
webseiten_list_offnerFehler = []  # Modullist (Moduls with errors)
webseiten_list_ohne_offnerFehler = []  # Modullist (Moduls without errors)
modulname_liste = []  # Modulnamelist (--goto selection)
result = []  # temporary storage (from --goto selection)
pfade_list_pi = []  # extra Websitelinks (from Pfade_list_pi1.json)


def parse_args():
    """
    Parse arguments.

    Handle the argument parser and save
    the information in a global variable.

    :return:
    """
    input_arguments = ArgumentParser()

    # Path argument: changetime
    input_arguments.add_argument('--time',
                                 required=False,
                                 default=25,
                                 type=int,
                                 help=' Mit dem Befehl und einer ganzen Zahl dahinter kann man den Wechselintervall der Websiten bestimmen.')

    # Project name argument: selection of moduls
    input_arguments.add_argument('--goto',
                                 action="store_true",
                                 required=False,
                                 help=' Mit dem Befehl kann man einzelne Module, die man ansehen möchte auswählen. ')

    # Project name argument: after x modulwebsite one extra website
    input_arguments.add_argument('--comm',
                                 required=False,
                                 default=5,
                                 type=int,
                                 help=' Mit dem Befehl kann man festlegen, nach wie vielen Modulen die Hauptübersicht oder die Zusatzseiten angezeigt wird / werden.')
    # Project name argument: moduls with errors
    input_arguments.add_argument('--wbug',
                                 required=False,
                                 action="store_true",
                                 help=' Mit dem Befehl werden nur die Module angezeigt, die noch offene Fehler haben')
    return vars(input_arguments.parse_args())


def get_sorted_module_list_ohne_fehler(liste):
    """
        Converts Modulname into a list for calling. (Moduls without Errors)

        :param liste: Modullist without errors
        :return: none
    """

    global webseiten_list
    for element in liste:
        webseiten_list.append("http://vms075/teststatistik/#/" + element + "/fehlerstatistik")

    return webseiten_list


def get_sorted_module_list(liste):
    """
        Converts Modulname into a list for calling

        :param liste: Modullist
        :return: none
    """

    global webseiten_list
    for element in liste:
        webseiten_list.append("http://vms075/teststatistik/#/" + element + "/teststatus")
        webseiten_list.append("http://vms075/teststatistik/#/" + element + "/fehlerstatistik")

    return webseiten_list


def run():
    """
    Websitelist calling

    :return: none
    """
    global Maincounter

    print(pfade_list_pi)
    chromewebdriver = "C:/Arbeitsplatz/python/raspiprojekt/chromedriver"
    browser = webdriver.Chrome(chromewebdriver)
    browser.maximize_window()
    browser.fullscreen_window()
    browser.get("http://powerscout.bender.de")  # Website login ("powerscout.bender")
    einlogen = browser.find_element_by_id("username")
    einlogen.click()
    einlogen.send_keys("SwQS_Bender")
    einlogen1 = browser.find_element_by_id("password")
    einlogen1.click()
    einlogen1.send_keys("SwQS2019")
    einlogen3 = browser.find_element_by_id("loginSubmit")
    einlogen3.click()
    time.sleep(10)

    # Only one time
    starttime = time.time()
    mouse = Controller()
    prev_mouse_pos = mouse.position

    while True:  # Infinitie loop
        counter_Seiten = 0
        counter_ex_Seiten = 0
        while counter_Seiten < len(webseiten_list):  # Loop for all websitecall
            # Currently mouse position
            curr_mouse_pos = mouse.position

            # Compare current mouse position with previous mouse position
            if curr_mouse_pos == prev_mouse_pos:
                # Mousemovement --> next website
                browser.get(webseiten_list[counter_Seiten])
                counter_Seiten += 1
                time.sleep(xtime)
                if counter_Seiten % zusatzseiten_counter == 0:  # --comm
                    if counter_ex_Seiten == len(pfade_list_pi):
                        counter_ex_Seiten = 0;
                    browser.get(pfade_list_pi[counter_ex_Seiten])
                    time.sleep(xtime)
                    counter_ex_Seiten += 1
                    if counter_ex_Seiten >= len(pfade_list_pi):  # Not again login powerscout.bender
                        try:
                            left_menu_dashboards = browser.find_element_by_id("left_menu_dashboards")
                            if left_menu_dashboards:
                                left_menu_dashboards.click()
                            time.sleep(xtime)
                            counter_ex_Seiten = 0
                        except NoSuchElementException as ex:
                            print(ex)
                    continue
                Maincounter += 1
            # Break
            time.sleep(xtime - ((time.time() - starttime) % xtime))
            # Update previos mouseposition
            prev_mouse_pos = curr_mouse_pos


def start():
    """

    :return:
    """
    global Maincounter
    global max_length_Resultlist
    global xtime
    global zusatzseiten_counter

    while True:
        # Build first jupiter websitelist
        teststatistik_JSON_datei = requests.get("http://vms075/teststatistik/jsonInput/graphData.json").json()
        arguments = parse_args()

        if arguments["wbug"]:
            for modulname in teststatistik_JSON_datei:
                summe_der_fehler_dynamic = sum(
                    teststatistik_JSON_datei[modulname]["dynamic"][0]["openError"][0].values())
                summe_der_fehler_static = sum(teststatistik_JSON_datei[modulname]["static"][0]["openError"][0].values())

                if summe_der_fehler_dynamic > 0 or summe_der_fehler_static > 0:
                    modulname = modulname.replace("/", "^")  # Change "/" to "^"
                    modulname_liste.append(modulname)
        else:
            for modulname in teststatistik_JSON_datei:
                dic = teststatistik_JSON_datei[modulname]["testStatus"][0]
                if dic["new"] == 0 and dic["running"] == 0 and dic["feedback"] == 0 and dic["return"] == 0 and dic[
                    "entwurf"] == 0:
                    modulname = modulname.replace("/", "^")  # Change "/" to "^"
                    webseiten_list_ohne_offnerFehler.append(modulname)
                else:
                    modulname = modulname.replace("/", "^")  # Change "/" to "^"
                    modulname_liste.append(modulname)
        with open("/home/pi/Desktop/Pfade_list_pi1.json", "r") as json_file:  # Read jsonfile ("Pfade_list_pi1.json")
            neue_liste = json.load(json_file)
            for path in neue_liste["Pfade"]:
                pfade_list_pi.append(path)

        if arguments["time"]:
            xtime = arguments["time"]
        if arguments["comm"]:
            if Maincounter <= len(webseiten_list_ohne_offnerFehler):
                zusatzseiten_counter = arguments["comm"]
            else:
                zusatzseiten_counter = 2 * arguments["comm"]
        if arguments["goto"]:
            title = 'Bitte wählen Sie die Module Ihrer Wahl (drücken Sie die LEERTASTE, um sie zu markieren und die EINGABETASTE, um fortzufahren): '
            options = modulname_liste
            selected = pick(options, title, multi_select=True, min_selection_count=1)
            print(selected)
            for t in selected:
                for x in t:
                    result.append(x)
                while max_length_Resultlist < len(result):
                    webseiten_list_offnerFehler.append(result[max_length_Resultlist])
                    max_length_Resultlist += 2
            get_sorted_module_list(webseiten_list_offnerFehler)
        else:
            get_sorted_module_list_ohne_fehler(webseiten_list_ohne_offnerFehler)
            get_sorted_module_list(modulname_liste)
        run()


if __name__ == 'Raspi_RDTS':
    start()
