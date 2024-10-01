import csv

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time

driver = webdriver.Chrome()

# Navigate to the page
url = "https://opi.dfo.kz/p/ru/dfo-search/accountants-search"
driver.get(url)

# Wait for the button to be clickable and then click it
wait = WebDriverWait(driver, 10)
search_button = wait.until(EC.element_to_be_clickable((By.ID, "search-AccountantsSearchResultTable")))
search_button.click()

# Wait for the search results to load
time.sleep(5)

# Find all elements with the class name "btn-goto-object"
contacts_buttons = driver.find_elements(By.CLASS_NAME, "btn-goto-object")

contacts = []

# Loop over each contact button and perform actions
for i, button in enumerate(contacts_buttons):
    # Open the contact page in a new tab
    button.send_keys(Keys.CONTROL + Keys.RETURN)

    # Switch to the new tab
    driver.switch_to.window(driver.window_handles[-1])

    try:
        # Wait for the "Контакты" link to be visible and click it
        contacts_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Контакты")))
        contacts_link.click()

        # Wait for the page to load after clicking
        time.sleep(3)

        # Verify the "Контакты" link has been clicked and print the page content
        print(f"Content of page {i + 1}:")
        soup = BeautifulSoup(driver.page_source, "lxml")
        company_name = soup.find("label", attrs={"for": "flnameru"})
        ceo = soup.find("label", attrs={"for": "flfirstperson"})
        accountant = soup.find("label", attrs={"for": "flaccname"})
        country = soup.find("label", attrs={"for": "fladrcountry"})
        post_index = soup.find("label", attrs={"for": "fladrindex"})
        region = soup.find("label", attrs={"for": "fladrreg"})
        address = soup.find("label", attrs={"for": "fladradr"})
        phones = soup.find("label", attrs={"for": "fladrphone"})
        emails = soup.find("label", attrs={"for": "fladrmail"})
        website = soup.find("label", attrs={"for": "fladrweb"})

        # Add the parsed contact to the list
        contacts.append({
            'Company': company_name.text.strip() if company_name else None,
            'CEO': ceo.text.strip() if ceo else None,
            'Accountant': accountant.text.strip() if accountant else None,
            'Country': country.text.strip() if country else None,
            'Post Index': post_index.text.strip() if post_index else None,
            'Region': region.text.strip() if region else None,
            'Address': address.text.strip() if address else None,
            'Phones': phones.text.strip() if phones else None,
            'Emails': emails.text.strip() if emails else None,
            'Website': website.text.strip() if website else None
        })

    except Exception as e:
        print(f"An error occurred on page {i + 1}: {e}")

    # Close the current tab and switch back to the original tab
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    # Write the contacts to a CSV file
    csv_file = 'contacts.csv'

    # Define the column names (keys in the dictionaries)
    csv_columns = ['Company', 'CEO', 'Accountant', 'Country', 'Post Index', 'Region', 'Address', 'Phones', 'Emails', 'Website']

    try:
        with open(csv_file, 'w', newline='', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=csv_columns)
            writer.writeheader()
            writer.writerows(contacts)
            print(f"Contacts saved to {csv_file}")
    except IOError:
        print("I/O error when writing to CSV file")

# Close the browser when done
driver.quit()
