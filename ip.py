from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import time
import os

# Nastavení WebDriveru
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')

# Zde si můžete zvolit jiný prohlížeč (např. Firefox)
driver = webdriver.Chrome(options=options)

# URL stránky
url = 'https://mysignins.microsoft.com/'

# Otevření stránky
driver.get(url)

# Čekání na přihlášení uživatele
input("Přihlaste se ke svému účtu Microsoft a po úspěšném přihlášení stiskněte Enter...")

# Nastavení limitu záznamů ke zpracování
LIMIT = 1000  # Můžete si upravit podle potřeby

# Soubor, do kterého budeme ukládat data
output_file = 'microsoft_signin_data.txt'

# Inicializace počítadla záznamů
record_count = 0

# Otevření souboru pro zápis s UTF-8 kódováním
with open(output_file, 'w', encoding='utf-8') as file:
    try:
        # Čekání na načtení seznamu aktivit
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'ms-Collapsible')))

        while record_count < LIMIT:
            # Refresování seznamu kontejnerů s aktivitami
            containers = driver.find_elements(By.CLASS_NAME, 'ms-Collapsible')

            for container in containers:
                try:
                    header = container.find_element(By.CLASS_NAME, 'ms-CollapsibleHeader')

                    # Pokud kontejner není rozbalený, rozbal ho
                    if header.get_attribute('aria-expanded') == 'false':
                        header.click()
                        time.sleep(1)  # Krátké čekání pro rozbalení

                    # Extrakce dat z rozbaleného kontejneru
                    time_data = header.get_attribute('aria-label')
                    
                    # Získání prvních 5 slov z time_data
                    time_data = ' '.join(time_data.split()[:5])
                    
                    # Vyhledání statusu a IP adresy
                    status_element = container.find_element(By.CLASS_NAME, 'statusLabelTestSelector')
                    status = status_element.text.strip()
                    ip_element = container.find_element(By.CLASS_NAME, 'ipContainerSelector')
                    ip_address = ip_element.text.strip()

                    # Kontrolní výpisy pro debugování
                    print(f"Časový údaj: {time_data}")
                    print(f"IP adresa: {ip_address}")
                    print(f"Výsledek přihlášení: {status}")

                    # Zapsání do souboru
                    file.write(f"{time_data};{ip_address};{status}\n")

                    record_count += 1
                    if record_count >= LIMIT:
                        break

                except StaleElementReferenceException:
                    # Ignorování chyby, pokusíme se to na další iteraci
                    continue
                except TimeoutException:
                    # Ignorování chyby při hledání prvků, pokusíme se to na další iteraci
                    continue

            # Skrolování dolů pro načtení dalších záznamů
            actions = ActionChains(driver)
            actions.send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(2)  # Čekání na načtení nových záznamů

    finally:
        # Uzavření prohlížeče
        driver.quit()

# Výpis výsledku
print(f"Extrahováno a uloženo {record_count} záznamů do souboru '{output_file}'.")
