from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
import time
import csv
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from datetime import datetime
from selenium.webdriver.chrome.service import Service
service = Service()
options = webdriver.ChromeOptions()
# your custom options...
options.page_load_strategy = 'eager'
driver = webdriver.Chrome(
    service=service,
    options=options
)

driver.maximize_window()
driver.get('https://www.close.com/integrations')

with open('close.csv', 'w', newline='',encoding='utf-8') as csvfile:
    fieldnames = ['Listing URL', 'ListingName', 'Built By', 'Icon URL',
                  'Description', 'Connect with xxx URL', 'ListingRank', 'ListingScrapeDate']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()

    

try:
    data = [] 
    main_links = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="w-dyn-list"]//div[@role="list"]//div[@role="listitem"]//a[@class="integrations_menu-link"]')))
    for _ in range(len(main_links))[6:]:
        try:
            time.sleep(2)
            current_category = main_links[_].text
            print(f"Current Category: {current_category}")
            link = main_links[_]
            driver.execute_script("window.open(arguments[0], '_blank');", link.get_attribute('href'))

            # Switch to the new tab
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(2)

            a_links = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="w-dyn-list"]//div[@class="w-dyn-item"]//a[@class="integrations_category-grid-item w-inline-block"]')))
            for idx, link in enumerate(a_links, start=1):
                try:
                    driver.execute_script("window.open(arguments[0], '_blank');", link.get_attribute('href'))

                    # Switch to the new tab
                    driver.switch_to.window(driver.window_handles[-1])
                    element = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.TAG_NAME, "div"))
                    )

                    list_page_url = driver.current_url
                    scraping_date = datetime.now().strftime("%Y-%m-%d")
                    rank = f"{current_category}-{idx}"
                    print(rank)
                    time.sleep(2)
                    try:
                    # Example: Extract the close title title
                        name = driver.find_element(By.XPATH, '//h1[@class= "heading-style-h2 display-inline"]').text
                    except NoSuchElementException:
                        name = ""
                    # Example: Extract the Built By
                    try:
                        built_by = driver.find_element(By.XPATH, '//div[@class="integrations-individual_built-by-company"]/div[2]').text 
                    except NoSuchElementException:

                        built_by = ""
                    print(built_by)
                    try:
                        icon_element = driver.find_element(By.XPATH, '//div[@class ="integrations-header_icon" ]')
                        # Taking path from the bg_image
                        background_image = icon_element.value_of_css_property("background-image")
                        # Split the string and check if there are at least two elements
                        split_background = background_image.split('"')
                        if len(split_background) >= 2:
                            icon_url = split_background[1]
                        else:
                            icon_url = ""
                    except NoSuchElementException:
                        icon_url = ""
                    # Scrape description
                    try:
                        description = driver.find_element(By.XPATH, '//div[@class = "mj-guide-rich-text margin-0 w-richtext"]').text
                    except NoSuchElementException:
                        description = ""
                    try:
                        xxx_url = driver.find_element(By.LINK_TEXT, 'Connect with Close')
                        url = xxx_url.get_attribute('href')
                    except NoSuchElementException:
                        url= ""
                    
                    data.append({
                        'Listing URL':list_page_url,
                        'ListingName':name, 
                        'Built By':built_by, 
                        'Icon URL':icon_url,
                        'Description':description,
                        'Connect with xxx URL':url,
                        'ListingRank':rank,
                        'ListingScrapeDate':scraping_date
                    })


                    driver.close()
                    driver.switch_to.window(driver.window_handles[-1])
                    # Update the list of links inside the loop
                    a_links = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="w-dyn-list"]//div[@class="w-dyn-item"]//a[@class="integrations_category-grid-item w-inline-block"]')))

                except Exception as e:
                    # Handle exceptions here (e.g., skip non-clickable links)
                    print(f"Skipping link due to exception: {str(e)}")
                    continue
                    a_links = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="w-dyn-list"]//div[@class="w-dyn-item"]//a[@class="integrations_category-grid-item w-inline-block"]')))
            driver.close()
            driver.switch_to.window(driver.window_handles[-1])
        except Exception as e:
                    # Handle exceptions here (e.g., skip non-clickable links)
            print(f"Skipping link due to exception: {str(e)}")
            continue
            main_links = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="w-dyn-list"]//div[@role="list"]//div[@role="listitem"]//a[@class="integrations_menu-link"]')))
    # Write data to CSV file in bulk
    with open('close.csv', 'a', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        for row in data:
            writer.writerow(row)

finally:
    driver.quit()