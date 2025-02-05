from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
import time
import os

# Erstellen eines Profils für den Firefox-Browser, um die Sitzung beizubehalten
profile_path = "/path/to/your/firefox/profile"  # Bitte ersetzen durch den Pfad zu deinem Firefox-Profil

# Funktion, die den gesamten Ablauf ausführt
def run_process(driver):
    try:
        # Die Kalenderseite in zwei Tabs öffnen
        url = "https://outlook.office365.com/calendar/published/26cedecc84d5479da18cd374edb8c267@nuuone.de/78c202f1f89e46f3a620b84f2a71854111886123736600783633/calendar.html"

        # Falls keine Tabs offen sind, neue öffnen
        if len(driver.window_handles) == 0:
            driver.get(url)
            driver.execute_script("window.open('');")
            driver.switch_to.window(driver.window_handles[1])
            driver.get(url)
        else:
            print("Bestehende Tabs gefunden. Fortfahren mit dem Ablauf.")

        # Funktion, die im aktuellen Tab zur Wochenansicht wechselt
        def switch_to_week_view(driver):
            try:
                wait = WebDriverWait(driver, 10)
                ansicht_wechsler_dropdown = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div/div/div[2]/div[2]/div/div/div/div[2]/div[1]/button"))
                )
                ansicht_wechsler_dropdown.click()
                wochenansicht_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div/ul/li[2]/button"))
                )
                wochenansicht_button.click()
                print("Zur Wochenansicht gewechselt.")
            except Exception as e:
                print("Fehler beim Wechseln zur Wochenansicht:", e)

        # Funktion, um den Kalender-Button an der Seite aufzuklappen
        def open_calendar_sidebar(driver):
            try:
                wait = WebDriverWait(driver, 10)
                kalender_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div/div/div[1]/div/div/button"))
                )
                kalender_button.click()
                print("Kalender-Button aufgeklappt.")
            except Exception as e:
                print("Fehler beim Aufklappen des Kalender-Buttons:", e)

        # Funktion, um zur nächsten Woche zu wechseln
        def switch_to_next_week(driver):
            try:
                wait = WebDriverWait(driver, 10)
                next_week_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div/div/div[1]/div/div[2]/div/div/div[2]/div[2]/table/tbody/tr[7]/td[4]"))
                )
                next_week_button.click()
                print("Zur nächsten Woche gewechselt.")
            except Exception as e:
                print("Fehler beim Wechseln zur nächsten Woche:", e)

        # Wechsel zur Wochenansicht im ersten Tab
        driver.switch_to.window(driver.window_handles[0])
        print("Wechsel zur ersten Woche")
        switch_to_week_view(driver)

        # Wechsel zur Wochenansicht im zweiten Tab und Kalender-Button aufklappen
        driver.switch_to.window(driver.window_handles[1])
        print("Wechsel zur zweiten Woche")
        switch_to_week_view(driver)
        open_calendar_sidebar(driver)

        # Warte kurz, um sicherzustellen, dass der Kalender-Button vollständig aufgeklappt ist
        time.sleep(5)

        # Zur nächsten Woche wechseln
        switch_to_next_week(driver)

    except Exception as e:
        print("Ein Fehler ist aufgetreten:", e)

# Hauptprogramm: Browser mit Profil starten und Schleife ausführen
if __name__ == "__main__":
    options = Options()
    options.set_preference("profile", profile_path)
    service = Service("/path/to/geckodriver")  # Bitte ersetzen durch den Pfad zu deinem Geckodriver

    driver = webdriver.Firefox(service=service, options=options)
    try:
        while True:
            print("Starte neuen Zyklus...")
            run_process(driver)
            print("Warte 60 Minuten, bevor der nächste Zyklus beginnt.")
            time.sleep(3600)  # 3600 Sekunden = 60 Minuten
    finally:
        driver.quit()