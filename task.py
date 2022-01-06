from main import *
from config import *


def main():
    try:
        dive_into_website(url, locator)
        create_workbook(workbook_name)
        create_worksheet(worksheet_name)
        scrape_agency_data_populate_records(worksheet_name)
        get_individual_investment_colnames()
        scrape_individual_investment()
        download_business_case_pdf()
        browser.close_all_browsers()
        compare_uii_and_investment_title(workbook_name)
    finally:
        minimal_task()


if __name__ == "__main__":
    main()
