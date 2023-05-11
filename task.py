import urllib.request
import re
from openpyxl import Workbook
from RPA.Browser.Selenium import Selenium
import json

# Define variables
url = "https://www.nytimes.com/"

# Variables for identify objects in Web. All finders are XPath
buttonSearch="//button[@Class='css-tkwi90 e1iflr850']"
textArea="//input[@Class='css-1j26cud']"
buttonGo="//button[@Class='css-1gudca6 e1iflr852']"
titlesIdXpath="//h4[@Class='css-2fgx4k']"
datesIdXpath="//span[@Class='css-17ubb9w']"
descriptionsIdXpath="//p[@Class='css-16nhkrn']"
imagesIdXpath="//img[@Class='css-rq4mmj']"
sectionButton="//button[@Class='css-4d08fs']"
labelSections="//label[@Class='css-1a8ayg6']"

# Selenium lib
browser_lib = Selenium()

# Define headers for Excel
headers=["Title", "Date", "Description", "Picture filename", "Count Phrase", "T or F contains $"]

# Define main functions

def load_config_from_json(file_path):
    # Load and config JSON for Work Items
    with open(file_path, 'r') as json_file:
        data = json.load(json_file)
    return data

def go_and_search(url, buttonSearch, textArea, search_phrase, buttonGo):
    # Open a browser
    browser_lib.open_available_browser(url)

    # Click on the search button to enable field
    browser_lib.click_button(buttonSearch)
    browser_lib.input_text(textArea,search_phrase)
    browser_lib.click_button(buttonGo)

def filter_category(news_category):
    # Apply filter if available
    checkboxes = browser_lib.find_elements("css:.css-1qtb2wd label.css-1a8ayg6")

    for checkbox in checkboxes:
        checkbox_text = browser_lib.get_text(checkbox)
        if news_category.lower() in checkbox_text.lower():
            browser_lib.click(checkbox)
            break

# def filter_date(num_months):
    # if num_months != None:
    #     browser_lib.click_element(dateRange)
    #     browser_lib.click_button(specificDate)

def extract_elements_titles():
    # Extract titles 
    browser_lib.wait_until_page_contains_element(titlesIdXpath)
    titles= browser_lib.get_webelements(titlesIdXpath)
    return titles

def extract_elements_dates():
    # Extract dates
    browser_lib.wait_until_page_contains_element(datesIdXpath)
    dates= browser_lib.get_webelements(datesIdXpath)
    return dates

def extract_elements_descriptions():
    # Extract descriptions
    browser_lib.wait_until_page_contains_element(descriptionsIdXpath)
    descriptions = browser_lib.get_webelements(descriptionsIdXpath)
    return descriptions

def extract_elements_images():
    # Extract images
    browser_lib.wait_until_page_contains_element(imagesIdXpath)
    images = browser_lib.get_webelements(imagesIdXpath)
    return images

def open_excel(titles, dates, descriptions, images, search_phrase):

    # Create Excel
    wb = Workbook()

    # Active Excel
    ws = wb.active

    # Write headers
    for col, head in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.value= head

    # Write titles
    for row, element in enumerate(titles, start=2):
        cell = ws.cell(row=row, column=1)
        cell.value= element.text
        count = count_phrase(search_phrase, element.text)
        cell_count = ws.cell(row=row, column=5)
        accumulate_count(cell_count, count)
        match=find_money_formats(element.text)
        cell_bool = ws.cell(row=row, column=6)
        assign_boolean_value(match, cell_bool)

    # Write dates    
    for row, element in enumerate(dates, start=2):
        cell = ws.cell(row=row, column=2)
        cell.value= element.text

    # Write descriptions
    for row, element in enumerate(descriptions, start=2):
        cell = ws.cell(row=row, column=3)
        cell.value= element.text
        count = count_phrase(search_phrase, element.text)
        cell_count = ws.cell(row=row, column=5)
        accumulate_count(cell_count, count)
        match=find_money_formats(element.text)
        cell_bool = ws.cell(row=row, column=6)
        assign_boolean_value(match, cell_bool)

    # Write images names
    for row, element in enumerate(images, start=2):

        url_image = element.get_attribute("src")

        name_file = f"image_{row}.jpg"
        path_download = f"output/{name_file}"
        urllib.request.urlretrieve(url_image, path_download)

        cell = ws.cell(row=row, column=4)
        cell.value = name_file

    # Save Excel file
    wb.save("output/list_news.xlsx")

def close_browser():
    # Cerrar el navegador
    browser_lib.close_all_browsers()

# Define inner functions
# It is to count the number of times the keyword appears in the texts. 
def count_phrase(phrase, text):
    count = text.lower().count(phrase.lower())
    return count

# It is to accumulate the amount already counted and the new one.
def accumulate_count (cell, count):

    current_count = cell.value
    if current_count is None:
        current_count = 0

    new_count = current_count + count

    cell.value = new_count

# It is to find the correct $ format in the texts.
def find_money_formats(text):
    pattern = r'\$[\d,]+(\.\d+)?|\d+ dollars|\d+ USD'
    matches = re.findall(pattern, text)
    return matches

# It is to assign true or false if the keyword appears.
def assign_boolean_value(bool_value, cell):
    if bool_value:
        if cell.value is None or cell.value is False:
            cell.value = True
    else:
        cell.value = False

# Call main functions
def main():

    config = load_config_from_json('work_items/config.json')

    search_phrase = config.get('search_phrase', 'time')
    news_category = config.get('news_category', '')
    num_months = config.get('num_months', 0)

    go_and_search(url, buttonSearch, textArea, search_phrase, buttonGo)
    if news_category != "":
        filter_category(news_category)
    # if num_months is not None:
    #     filter_date(num_months)
    titles = extract_elements_titles()
    dates = extract_elements_dates()
    descriptions = extract_elements_descriptions()
    images = extract_elements_images()
    open_excel(titles, dates, descriptions, images, search_phrase)
    close_browser()

# Module and call main()
if __name__ == "__main__":
    main()