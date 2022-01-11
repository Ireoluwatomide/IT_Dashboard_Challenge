from main import ITDashboard

robot = ITDashboard()


def main():
    robot.scrape_agency_data()
    robot.scrape_individual_investment()
    robot.compare_uii_and_title()
    robot.save_workbook()


if __name__ == "__main__":
    main()
