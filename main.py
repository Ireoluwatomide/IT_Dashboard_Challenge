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


# -

class ITDashboard:
    agency_urls = []
    column_names = []
    links = []
    records = {}

    def __init__(self):
        self.browser = Selenium()
        self.browser.set_download_directory(os.path.join(os.getcwd(), f"{OUTPUT_DIR}"))
        self.excel = Files()
        self.pdf = PDF()

    def dive_into_website(self):

        self.browser.open_available_browser("https://itdashboard.gov")
        self.browser.wait_until_page_contains_element("xpath://*[@id='node-23']")
        self.browser.click_element("xpath://*[@id='node-23']")
        sleep(3)

    def scrape_agency_data(self):
        self.dive_into_website()

        # Scrape the agency data from the website
        agency_elements = self.browser.find_elements(
                '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        serial = ["S/N", ]
        agency = ["Agency", ]
        spending = ["FY2021 Spending ($)"]
        num = 1

        for element in agency_elements:
            url = self.browser.find_element(element).find_element_by_tag_name("a").get_attribute("href")
            self.agency_urls.append(url)
            agency_data = element.text.split('\n')
            spending_amount = (agency_data[2]).split("$")
            agency.append(agency_data[0])
            spending.append(spending_amount[1])
            serial.append(num)
            num += 1

        # Populate the data into the worksheet
        records = {"serial": serial, "agency": agency,  "FY2021": spending}
        self.excel.create_workbook("output/IT_Dashboard.xlsx")
        self.excel.rename_worksheet('Sheet', "Agencies")
        self.excel.append_rows_to_worksheet(records, "Agencies")

    def get_individual_investment_colnames(self):
        self.browser.go_to(self.agency_urls[-1])
        sleep(1)

        while True:
            try:
                colnames = self.browser.find_element(
                    '//table[@class="datasource-table usa-table-borderless dataTable no-footer"]'
                ).find_element_by_tag_name(
                    "thead").find_elements_by_tag_name("tr")[1].find_elements_by_tag_name("th")
                if colnames:
                    break
            except Exception:
                sleep(1)

        for col in colnames:
            self.column_names.append(col.text)

    def scrape_individual_investment(self):
        self.get_individual_investment_colnames()
        # Navigate the website to get the total number of enteries
        entries = self.browser.find_element('//*[@id="investments-table-object_info"]')
        entries_data = entries.text.split(" ")
        total_entries = int(entries_data[-2])
        self.total_entry = total_entries

        # Navigate through the website to load the entire table
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select').click()
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
        self.browser.wait_until_page_contains_element(
            f'//*[@id="investments-table-object"]/tbody/tr[{total_entries}]/td[1]',
            timeout=timedelta(seconds=10))

        # Define the column names: This is where the records will be appended to
        UII_Ids = [self.column_names[0], ]
        Bureau = [self.column_names[1], ]
        Investment_Title = [self.column_names[2], ]
        Total_FY2021 = [self.column_names[3], ]
        Type = [self.column_names[4], ]
        CIO_Rating = [self.column_names[5], ]
        Num_of_project = [self.column_names[6], ]

        # Iterate through the total enteries to scrape the individul investment table
        for uii in range(1, total_entries + 1):
            uii_ids = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[1]')
            UII_Ids.append(uii_ids.text)

            try:
                bureau = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[2]').text
                investment_title = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[3]').text
                total_FY2021 = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[4]').text
                agency_type = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[5]').text
                CIO_rating = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[6]').text
                num_of_project = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[7]').text
            except Exception:
                bureau = ""
                investment_title = ""
                total_FY2021 = ""
                type_agency = ""
                CIO_rating = ""
                num_of_project = ""
            try:
                link = self.browser.find_element(
                        f'//*[@id="investments-table-object"]/tbody/tr[{uii}]/td[1]').find_element_by_tag_name(
                        "a").get_attribute("href")
            except Exception:
                link = ""
            self.links.append(link)

            Bureau.append(bureau)
            Investment_Title.append(investment_title)
            Total_FY2021.append(total_FY2021)
            Type.append(agency_type)
            CIO_Rating.append(CIO_rating)
            Num_of_project.append(num_of_project)
        self.records = {"uii": UII_Ids,
                        "bureau": Bureau,
                        "investment_title": Investment_Title,
                        "FY2021": Total_FY2021,
                        "agency_type": Type,
                        "CIO rating": CIO_Rating,
                        "# of project": Num_of_project}

    def compare_uii_and_title(self):
        pdf_flag = ["PDF_Flag", ]
        uiis = self.records["uii"][1:]
        investment_titles = self.records["investment_title"][1:]
        # Iterate through the links list to download the business case PDF
        for link in self.links:
            try:
                self.browser.go_to(link)
                self.browser.wait_until_element_is_visible('//div[@id="business-case-pdf"]',
                                                      timeout=timedelta(seconds=5))
                self.browser.find_element('//div[@id="business-case-pdf"]').click()
                sleep(5)
                
            except Exception:
                link = ""
        for uii, investment_title in zip(uiis, investment_titles):
            try:
                self.pdf.extract_pages_from_pdf(source_path=f"output/{uii}.pdf",
                                                output_path=f"output/new{uii}.pdf",
                                                pages=1)
            except Exception:
                flag = "No link"
                pdf_flag.append(flag)
            try:
                text = self.pdf.get_text_from_pdf(f"output/new{uii}.pdf")
            except Exception:
                continue

            if investment_title in text[1]:
                flag = "True"
                pdf_flag.append(flag)
            else:
                flag = "False"
                pdf_flag.append(flag)
            os.remove(f"output/new{uii}.pdf")
        self.records["PDF_Flag"] = pdf_flag

    def save_workbook(self):
        self.excel.create_worksheet("Individual Investment")
        self.excel.append_rows_to_worksheet(self.records, "Individual Investment")
        self.excel.save_workbook("output/IT_Dashboard.xlsx")
        self.browser.close_all_browsers()
