import os, comtypes.client
import tkinter as tk
import pyperclip
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

# ==========Globale Variablen==========#

inp = ''
outp = ''


# ==========Tkinter GUI Steuerungsfunktionen==========#

def get_in_path():
    '''
    Hole Eingangspfad (Ordner der Word Dokumente)
    :return: Eingangspfad als String
    '''

    save_path = filedialog.askdirectory(
        title="Wähle Pfad zu Eingabe-Dokumenten"
    )

    if save_path:
        return save_path
    else:
        return "Kein Ordner angegeben"


def get_out_path():
    '''
    Hole Speicherpfad (Ordner wo abgespeichert wird)
    :return: Speicherpfad als String
    '''

    save_path = filedialog.askdirectory(
        title="Wähle Pfad zu Speicher-Dokumenten"
    )

    if save_path:
        return save_path
    else:
        return "Kein Ordner angegeben"


def btn_Indestination_folder_pressed():
    '''
    Ruft get_in_path() auf, sobald Knopf in Fenster gedrückt
    '''

    in_path = get_in_path()
    tk_in_text.set(in_path)
    global inp
    inp = in_path


def btn_Outdestination_folder_pressed():
    '''
    Ruft get_out_path() auf, sobald Knopf in Fenster gedrückt
    '''

    out_path = get_out_path()
    tk_out_text.set(out_path)
    global outp
    outp = out_path


def get_directory_length(output_path):
    '''
    Zählt Dateien im input_path
    :param: Speicherpfad
    :return: Anzahl der Dateien im Eingabepfad
    '''

    dir_content_counter = 0
    for file in os.listdir(output_path):
        dir_content_counter += 1
    return dir_content_counter


def open_output_path():
    '''
    Öffnet den Speicherpfad für mögliche Kontrolle vor dem Upload
    '''

    os.startfile(outp)


def bar():
    '''
    Angepasster Fremdcode (siehe Quelle 4)
    Funktion lässt einen Ladebalken laufen
    '''

    progress_bar['value'] = 20
    window.update_idletasks()
    sleep(1)

    progress_bar['value'] = 40
    window.update_idletasks()
    sleep(1)

    progress_bar['value'] = 60
    window.update_idletasks()
    sleep(1)

    progress_bar['value'] = 80
    window.update_idletasks()
    sleep(1)

    progress_bar['value'] = 100


# ==========Word-PDF Konvertierungsfunktionen (Angepasster Fremdcode (siehe Quelle 1))==========#

def word_pdf_konv(input_path, output_path):
    '''
    Führt Konvertierung durch
    :param input_path: Eingangspfad
    :param output_path: Speicherpfad
    '''

    pdf_format_key = 17
    file_in = os.path.abspath(input_path)
    file_out = os.path.abspath(output_path)
    word_app = comtypes.client.CreateObject('Word.Application')
    document = word_app.Documents.Open(file_in)
    file_out = file_out.replace(".docx", "")
    document.SaveAs(file_out, FileFormat=pdf_format_key)
    document.Close()
    word_app.Quit()


def start():
    '''
    Starten Konvertierung sobald Start-Knopf gedrückt wurde
    Konvertiert alle Dateien im Eingangspfad und löscht unerwünschte Dateien
    welche bei der Konvertierung entstehen können
    '''

    bar()
    destinationIn = inp
    destinationOut = outp

    for file in os.listdir(destinationIn):
        word_pdf_konv(destinationIn + "\\" + file, destinationOut + "\\" + file + ".pdf")

    for file in os.listdir(destinationOut):  # Schleife zum Löschen unerwünschter Dateien (erwähnt in Doku 3.1)
        if file.startswith("~$", 0, 2):
            os.remove(destinationOut + "\\" + file + ".pdf")

    tk_progress.set("Konvertierung abgeschlossen")


# ==========Funktionen für das Hochladen in kostenlose Cloud Services==========#

def upload_to_dropbox():
    '''
    Holt sich Pfad zum Chromedriver um Chrome zu starten
    Führt Web-scraping Schritte durch um Dateien hochzuladen
    '''

    chrome_path = get_chromedriver()
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(chrome_path)
    driver.get("https://www.dropbox.com/login")
    driver.maximize_window()

    wait = WebDriverWait(driver, 45)

    email = tk_email.get()
    password = tk_password.get()

    email_field = driver.find_element_by_name("login_email")
    password_field = driver.find_element_by_name("login_password")

    email_field.click()
    email_field.send_keys(email)

    password_field.click()
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    sleep(35)  # Dieser Schritt ist leider notwendig, da ich durch den automatischen Chrome-Aufruf aufgefordert werde,
    # mich als Mensch zu indentifizieren und sonst mein Programm fehler wirft, da es die folgenden HTML-Elemente noch nicht kennt

    out_now_Xbtn = driver.find_element_by_class_name("dig-IconButton-content")
    out_now_Xbtn.click()
    sleep(2)
    upload_folder_btn = driver.find_element_by_xpath("/html/body/div[1]/div[6]/div/div/div[2]/div/div/div/nav/ul/li[2]/button")
    upload_folder_btn.click()

    copy_value = outp.replace("/", "\\")
    pyperclip.copy(copy_value)  # Kopieren des Pfads in Zwischenablage

    # =========Warte bis Ordner ausgewählt (liegt in Zwischenablage)==========#

    try:
        sleep(10)
        upload_btn = driver.find_element_by_class_name("dig-Button dig-Button--primary dig-Button--standard")
        upload_btn.click()
    except:
        print("Schon wegge - x - t")


def upload_to_pcloud():
    '''
    Holt sich Pfad zum Chromedriver um Chrome zu starten
    Führt Web-scraping Schritte durch um Dateien hochzuladen
    '''

    chrome_path = get_chromedriver()
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(chrome_path)
    driver.get("https://www.pcloud.com/de/eu")
    driver.maximize_window()

    email = tk_email.get()
    password = tk_password.get()

    login_field = driver.find_element_by_xpath("/html/body/header/div[3]/div/div[2]/div[2]")
    sleep(2)
    login_field.click()
    sleep(2)

    email_field = driver.find_element_by_xpath("/html/body/div[17]/div/div/div/div/div/div[1]/input")
    password_field = driver.find_element_by_xpath("/html/body/div[17]/div/div/div/div/div/div[1]/div[3]/input")

    email_field.send_keys(email)
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    sleep(15)  # Warten bis Seite aufgebaut wurde

    upload_btn = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div[2]/div[1]/div[2]/div/div[1]/div/div[3]/div[3]/div[1]")
    upload_btn.click()
    sleep(1)
    dir_upload_btn = driver.find_element_by_xpath("/html/body/div[7]/div[1]/div[3]/div[1]/span[2]")
    dir_upload_btn.click()
    sleep(1)
    choose_dir_btn = driver.find_element_by_xpath("/html/body/div[7]/div[1]/div[3]/div[3]/div")
    choose_dir_btn.click()

    copy_value = outp.replace("/", "\\")
    pyperclip.copy(copy_value)  # Kopieren des Pfads in Zwischenablage


def upload_to_onedrive():
    '''
    Holt sich Pfad zum Chromedriver um Chrome zu starten
    Führt Web-scraping Schritte durch um Dateien hochzuladen
    '''

    chrome_path = get_chromedriver()
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(chrome_path)
    driver.get("https://onedrive.live.com/about/de-de/signin/")
    driver.maximize_window()

    wait = WebDriverWait(driver, 30)

    email = tk_email.get()
    password = tk_password.get()

    wait.until(EC.frame_to_be_available_and_switch_to_it((By.CLASS_NAME, "SignIn")))

    email_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.outer > div > main > div:nth-child(2) > div.row.margin-bottom-16 > div > input")))
    email_field.send_keys(email)
    move_on_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.outer > div > main > div:nth-child(2) > div.row.inline-block.no-margin-top-bottom.button-container > input")))
    move_on_btn.click()
    sleep(2)

    password_field = driver.find_element_by_id("i0118")
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)
    sleep(3)

    upload_btn = driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div[2]/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/button/span")
    upload_btn.click()

    dir_upload_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/ul/li[2]/button/div/span")))
    dir_upload_btn.click()

    copy_value = outp.replace("/", "\\")
    pyperclip.copy(copy_value)  # Kopieren des Pfads in Zwischenablage


def start_cloud_functions():
    '''
    Holt sich ausgewählten Cloud Dienst aus Dropdown Liste und startet die dazu passende Upload-Funktion
    '''

    cloud_option = tk_option_list.get()
    if cloud_option == option_list[0]:
        upload_to_dropbox()
    elif cloud_option == option_list[1]:
        upload_to_pcloud()
    elif cloud_option == option_list[2]:
        upload_to_onedrive()


def get_chromedriver():
    '''
    Öffnet chromedriver_path.txt und gibt gespeicherten Inhalt an Tkinter Variable chromedriver_text
    :return: Pfad des Chromedrivers
    '''

    chromedriver_path = open('chromdriver_path.txt', 'r').read()
    tk_chromedriver_text.set(chromedriver_path)
    if chromedriver_path:
        chromedriver_path = chromedriver_path + "/chromedriver.exe"
        return chromedriver_path
    else:
        return 'Chromedriver nicht gefunden'


# ==========Cloud option List (siehe Quelle 2)==========#

option_list = [
    "DropBox",
    "pCloud",
    "OneDrive"
]

if __name__ == "__main__":  # Bei Programmstart tue folgendes

    # ==========Tkinter Fenster Einstellungen==========#

    window = Tk()
    window.title("Word-PDF Konverter")
    window.geometry('426x700')

    # ==========Tkinter Variablen==========#

    tk_in_text = StringVar()
    tk_out_text = StringVar()
    tk_convert_counter = StringVar()
    tk_progress = StringVar()
    tk_chromedriver_text = StringVar()
    tk_email = StringVar()
    tk_password = StringVar()
    tk_option_list = StringVar()
    tk_cloud_service = StringVar()

    tk_option_list.set(option_list[0])  # Default des Drop-Down Menüs

    def on_optionChange_change_label(event):
        '''
        Angepasster Fremdcode (siehe Quelle 7)
        Dynamische Änderung des Fenster-Textes "Anmelden in ... DropBox/pCloud/OneDrive"
        '''

        login_label = Label(window, text=f"Anmelden in {tk_option_list.get()}", font="calibri 13 bold", )
        login_label.grid(row=13, column=0, pady=10)


    # ==========Erstellen der verschiedenen Tkinter Elemente==========#

    # Konvertierung
    askForIn_label = tk.Label(window, text="Wähle Ordner aus")
    askForOut_label = tk.Label(window, text="Wähle Zielordner aus")
    search_inp_btn = Button(window, text="Suche Ordner", command=btn_Indestination_folder_pressed)
    search_sav_btn = Button(window, text="Zielordner auswählen", command=btn_Outdestination_folder_pressed)
    inpath_entry = Entry(window, width=70, textvariable=tk_in_text, state="readonly")
    savepath_entry = Entry(window, width=70, textvariable=tk_out_text, state="readonly")
    openOutFolder_btn = Button(window, text="Zielordner öffnen", command=open_output_path)
    start_btn = tk.Button(window, text="Starte Konvertierung", command=start, bg="lightgreen", width=30, height=2, bd=1, relief="ridge")
    progress_bar = Progressbar(window)
    progress_label = Label(window, textvariable=tk_progress, font="bold")

    # Upload Cloud
    login_to_label = Label(window, text="In welche Cloud hochladen?", font="Calibri 13")
    options_dropdown = tk.OptionMenu(window, tk_option_list, *option_list, command=on_optionChange_change_label)
    login_label = Label(window, text=f"Anmelden in {option_list[0]}", font="calibri 13 bold")
    username_label = Label(window, text="Email-Adresse:")
    username_entry = Entry(window, textvariable=tk_email, width=35)
    password_label = Label(window, text="Passwort:")
    password_entry = Entry(window, textvariable=tk_password, show="*", width=35)
    upload_btn = tk.Button(window, text="Starte Upload", command=start_cloud_functions, bg="lightgreen", width=30, height=2, bd=1, relief="ridge")

    # =========Grid Einstellungen (Positionierung der einzelnen Elemente im Fenster)==========#

    askForIn_label.grid(row=1, column=0, pady=10)
    inpath_entry.grid(row=2, column=0, pady=5)
    search_inp_btn.grid(row=3, column=0, pady=5)

    askForOut_label.grid(row=4, column=0, pady=10)
    savepath_entry.grid(row=5, column=0, pady=5)
    search_sav_btn.grid(row=6, column=0, pady=5)

    openOutFolder_btn.grid(row=7, column=0, pady=5)

    start_btn.grid(row=8, column=0, pady=10)
    progress_bar.grid(row=9, column=0, pady=5)
    progress_label.grid(row=10, column=0, pady=3)

    login_to_label.grid(row=11, column=0, pady=10)
    options_dropdown.grid(row=12, column=0, pady=3)
    login_label.grid(row=13, column=0, pady=10)
    username_label.grid(row=14, pady=3)
    username_entry.grid(row=15, pady=3)
    password_label.grid(row=16, pady=3)
    password_entry.grid(row=17, pady=3)
    upload_btn.grid(row=18, pady=10)

    window.mainloop()
