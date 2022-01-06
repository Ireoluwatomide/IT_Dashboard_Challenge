"""This automates the process of extracting data from (http://itdashboard.gov/)"""

# Import the libraries
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from datetime import timedelta
from RPA.PDF import PDF
from time import sleep
import os

# +
OUTPUT_DIR = "output"

if not os.path.exists(OUTPUT_DIR):
    os.mkdir(OUTPUT_DIR)
browser = Selenium()
excel = Files()
pdf = PDF()
browser.set_download_directory(os.path.join(os.getcwd(), f"{OUTPUT_DIR}"))


# -

def dive_into_website(url, locator):

    try:
        browser.open_available_browser(url)
    except Exception:
        print("Unable to open the browser\n")
    else:
        print("Successfully opend the browser\n")

    try:
        sleep(1)
        browser.click_element(locator)
    except Exception:
        print("Unable to dive-in\n")
    else:
        print("Successfully dived-in\n")


def create_workbook(workbook_name):

    try:
        excel.create_workbook()
    except Exception:
        print("Failed to create workbook\n")
    else:
        print("Workbook created\n")


def create_worksheet(worksheet_name):

    # Create the Agencies worksheet
    try:
        excel.rename_worksheet('Sheet', worksheet_name)
        excel.set_worksheet_value(row=1, column=1, value="S/N")
        excel.set_worksheet_value(row=1, column=2, value="Agency")
        excel.set_worksheet_value(row=1, column=3, value="FY2021 Spending ($)")
    except Exception:
        print("Could not create worksheet\n")
    else:
        print("Worksheet created\n")


def scrape_agency_data_populate_records(worksheet_name):

    # Scrape the agency data from the website
    try:
        data = browser.find_elements(
            '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        serial = []
        agency = []
        spending = []
        num = 1

        for item in data:

            agency_data = item.text.split('\n')
            spending_amount = (agency_data[2]).split("$")
            agency.append(agency_data[0])
            spending.append(spending_amount[1])
            serial.append(num)
            num += 1
    except Exception:
        print("Unable to scrape the data\n")
    else:
        print("Successfully scrapped the data\n")

    # Populate the data into the worksheet
    try:
        records = {"serial": serial, "agency": agency,  "FY2021": spending}
        excel.append_rows_to_worksheet(records, worksheet_name)
    except Exception:
        print("Unable to populate the records to worksheet\n")
    else:
        print("Successfully populated the worksheet\n")


def get_individual_investment_colnames():

    agency_elements = browser.find_elements(
            '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')

    agency_urls = []

    for element in agency_elements:
        url = browser.find_element(element).find_element_by_tag_name("a").get_attribute("href")
        agency_urls.append(url)

    print(agency_urls, "\n")

    # browser.go_to(agency_urls[0])
    browser.go_to("https://itdashboard.gov/drupal/summary/429")
    sleep(1)

    column_names = []

    while True:
        try:
            colnames = browser.find_element(
                '//table[@class="datasource-table usa-table-borderless dataTable no-footer"]'
            ).find_element_by_tag_name(
                "thead").find_elements_by_tag_name("tr")[1].find_elements_by_tag_name("th")
            if colnames:
                break
        except Exception:
            sleep(1)

    for col in colnames:
        column_names.append(col.text)

    records = {"UII_Ids": [column_names[0]],
               "Bureau": [column_names[1]],
               "Investment_Title": [column_names[2]],
               "Total_FY2021": [column_names[3]],
               "Type": [column_names[4]],
               "CIO_Rating": [column_names[5]],
               "Num_of_project": [column_names[6]],
               "Pdf_flag": ["PDF Flag", ]}

    excel.create_worksheet("Individual Investment")
    excel.append_rows_to_worksheet(records, "Individual Investment")


def scrape_individual_investment():

    # Navigate the website to get the total number of enteries
    entries = browser.find_element('//*[@id="investments-table-object_info"]')
    entries_data = entries.text.split(" ")
    total_entries = int(entries_data[-2])

    # Navigate through the website to load the entire table
    browser.find_element('//*[@id="investments-table-object_length"]/label/select').click()
    browser.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
    browser.wait_until_page_contains_element(
        f'//*[@id="investments-table-object"]/tbody/tr[{total_entries}]/td[1]',
        timeout=timedelta(seconds=10))

    # Define the column names: This is where the records will be appended to
    UII_Ids = []
    Bureau = []
    Investment_Title = []
    Total_FY2021 = []
    Type = []
    CIO_Rating = []
    Num_of_project = []
    Pdf_match = []

    # Iterate through the total enteries to scrape the individul investment table
    for uii in range(1, total_entries + 1):
        uii_ids = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[1]')
        UII_Ids.append(uii_ids.text)

        try:
            bureau = browser.find_element(
                f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[2]').text
            investment_title = browser.find_element(
                f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[3]').text
            total_FY2021 = browser.find_element(
                f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[4]').text
            agency_type = browser.find_element(
                f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[5]').text
            CIO_rating = browser.find_element(
                f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[6]').text
            num_of_project = browser.find_element(
                f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[7]').text
        except Exception:
            bureau = ""
            investment_title = ""
            total_FY2021 = ""
            type_agency = ""
            CIO_rating = ""
            num_of_project = ""

        Bureau.append(bureau)
        Investment_Title.append(investment_title)
        Total_FY2021.append(total_FY2021)
        Type.append(agency_type)
        CIO_Rating.append(CIO_rating)
        Num_of_project.append(num_of_project)

    data = {"uii": UII_Ids,
            "bureau": Bureau,
            "company": Investment_Title,
            "FY2021": Total_FY2021,
            "agency_type": Type,
            "CIO rating": CIO_Rating,
            "# of project": Num_of_project}

    excel.append_rows_to_worksheet(data, "Individual Investment")


def download_business_case_pdf():

    entries = browser.find_element('//*[@id="investments-table-object_info"]')
    entries_data = entries.text.split(" ")
    total_entries = int(entries_data[-2])

    # Save the links on the website in a list called "links"
    links = []

    for uii in range(1, total_entries + 1):

        try:
            link = browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[1]').find_element_by_tag_name(
                    "a").get_attribute("href")
        except Exception:
            link = ""
        links.append(link)

    # Iterate through the links list to download the business case PDF
    for link in links:
        try:
            browser.go_to(link)
            browser.wait_until_page_contains_element('//div[@id="business-case-pdf"]',
                                                     timeout=timedelta(seconds=10))
            browser.find_element('//div[@id="business-case-pdf"]').click()
            sleep(6)
        except Exception:
            link = ""


def compare_uii_and_investment_title(workbook_name):

    # Read worksheet to extract the uii and investment titles
    investment_records = excel.read_worksheet("Individual Investment", header = True)

    uiis = []
    investment_titles = []

    for records in investment_records:

        investment_title = records['Investment Title']
        uii = records['UII']

        uiis.append(uii)
        investment_titles.append(investment_title)

    # Compare uii and investment titles
    pdf_flag = []

    for uii, investment_title in zip(uiis, investment_titles):
        try:
            pdf.extract_pages_from_pdf(source_path=f"output/{uii}.pdf",
                                       output_path=f"output/new{uii}.pdf",
                                       pages=1)
        except Exception:
            flag = "No link"
            pdf_flag.append(flag)

        try:
            text = pdf.get_text_from_pdf(f"output/new{uii}.pdf")
        except Exception:
            continue

        if investment_title in text[1]:
            flag = "True"
            pdf_flag.append(flag)
        else:
            flag = "False"
            pdf_flag.append(flag)

        os.remove(f"output/new{uii}.pdf")

    # Append the pdf_flags to the worksheet and save the workbook
    row = 2
    
    for flag in pdf_flag:
        excel.set_worksheet_value(
            row=row, column=8, value=flag, name="Individual Investment")
        row += 1
    
    excel.save_workbook(workbook_name)


def minimal_task():
    print("Completed.")
