from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import csv
import os
from openpyxl import Workbook, load_workbook

# Setup driver Chrome con webdriver-manager
options = Options()
# options.add_argument('--headless')  # scommenta per far girare senza GUI
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

start_url = "https://www.gelbeseiten.de/suche/cocktailbars/bundesweit"
driver.get(start_url)

csv_filename = 'cocktailbars.csv'
excel_filename = 'cocktailbars.xlsx'


def load_all_results():
    """Clicca ripetutamente su 'Mehr Anzeigen' fino a che il bottone non è più presente."""
    print("Caricamento di tutti i risultati...")
    while True:
        try:
            load_more_btn = wait.until(EC.element_to_be_clickable((By.ID, 'mod-LoadMore--button')))
            print("Clicco su 'Mehr Anzeigen' per caricare altri risultati...")
            driver.execute_script("arguments[0].scrollIntoView(true);", load_more_btn)  # Scrolla al bottone
            load_more_btn.click()
            # Attendi che nuovi risultati vengano caricati
            # Qui aspettiamo che il bottone "Mehr Anzeigen" sparisca o che vengano aumentati gli articoli
            time.sleep(2)  # piccola pausa per stabilità
            wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, 'article.mod-Treffer')) > 0)
            # Proviamo a vedere se il bottone appare di nuovo in 5s, altrimenti esci
            try:
                wait.until(EC.element_to_be_clickable((By.ID, 'mod-LoadMore--button')), timeout=5)
            except:
                print("Pulsante 'Mehr Anzeigen' non più visibile.")
                break
        except Exception:
            print("Nessun altro pulsante 'Mehr Anzeigen' trovato o errore, termino caricamento.")
            break


def save_to_csv(row, fieldnames):
    file_exists = os.path.isfile(csv_filename)
    with open(csv_filename, mode='a', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)


def save_to_excel(row):
    if os.path.exists(excel_filename):
        workbook = load_workbook(excel_filename)
        sheet = workbook.active
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['Nome', 'Indirizzo', 'Telefono', 'Sito Web', 'Email', 'Link Dettaglio'])

    sheet.append([row['Nome'], row['Indirizzo'], row['Telefono'], row['Sito Web'], row['Email'], row['Link Dettaglio']])
    workbook.save(excel_filename)


print("Inizio caricamento completo dei risultati...")
load_all_results()

print("Parsing pagina...")
soup = BeautifulSoup(driver.page_source, 'html.parser')
bars = soup.find_all('article', class_='mod-Treffer')

fieldnames = ['Nome', 'Indirizzo', 'Telefono', 'Sito Web', 'Email', 'Link Dettaglio']

print(f"Trovati {len(bars)} bar. Inizio estrazione dati...")

for index, bar in enumerate(bars, start=1):
    try:
        nome_tag = bar.find('h2', class_='mod-Treffer__name')
        nome = nome_tag.get_text(strip=True) if nome_tag else ''

        indirizzo_tag = bar.find('div', class_='mod-AdresseKompakt__adress-text')
        indirizzo = indirizzo_tag.get_text(separator=' ', strip=True) if indirizzo_tag else ''

        telefono_tag = bar.find('a', class_='mod-TelefonnummerKompakt__phoneNumber')
        telefono = telefono_tag.get_text(strip=True) if telefono_tag else ''

        dettaglio_link_tag = bar.find('a', href=True)
        dettaglio_link = dettaglio_link_tag['href'] if dettaglio_link_tag else ''
        if dettaglio_link.startswith('/'):
            dettaglio_link = "https://www.gelbeseiten.de" + dettaglio_link

        sito_web = ''
        email = ''

        if dettaglio_link:
            # Vai alla pagina dettaglio
            driver.get(dettaglio_link)
            try:
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'aktionsleiste-button')))
            except Exception:
                pass  # se non c'è non bloccare

            detail_soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Estraggo il sito web
            sito_tag = detail_soup.select_one('div.aktionsleiste-button a[href]')
            if sito_tag:
                sito_web = sito_tag['href']

            # Estraggo l'email se presente
            email_button = detail_soup.find('div', id='email_versenden')
            if email_button:
                data_link = email_button.get('data-link', '')
                if data_link.startswith('mailto:'):
                    email = data_link.split(':')[1].split('?')[0]

            driver.back()
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'mod-Treffer')))
            time.sleep(1)

        row = {
            'Nome': nome,
            'Indirizzo': indirizzo,
            'Telefono': telefono,
            'Sito Web': sito_web,
            'Email': email,
            'Link Dettaglio': dettaglio_link
        }

        print(f"[{index}/{len(bars)}] Salvo: {row}")
        save_to_csv(row, fieldnames)
        save_to_excel(row)

    except Exception as e:
        print(f"Errore elaborazione bar '{nome}': {e}")

driver.quit()
print("Scraping completato! File CSV e Excel salvati.")
