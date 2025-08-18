# Biding_Data_Extracion

This project contains a Python script that automates the process of scraping bid and contract information from the official Delaware State Bid Solicitation Directory. 
<br> It uses Selenium to navigate the website, extract data for various bid statuses (Open, Recently Closed, Awarded), and saves the collected information into a structured Excel file.

## Features
1. **Automated Browsing** : Navigates through different bid status tabs and paginated results automatically.

2. **Comprehensive Data Extraction** : Scrapes both summary data from the main tables and detailed information from modal pop-ups for each bid.

3. **Multi-Status Scraping** : Gathers data for 'Open', 'Recently Closed', and 'Awarded' bids in a single run.

4. **Structured Output** : Exports all collected data into a clean, well-organized Excel file (.xlsx) for easy analysis and use.

5. **Modular Code** : The script is broken down into logical functions for readability, maintenance, and scalability.

6. **Error Handling** : Includes mechanisms to handle common web scraping issues like timeouts and missing elements.

## Requirements
To run this script, you will need the following installed:

`Python 3.x`

`pandas` library

`selenium` library

`Google Chrome browser`

`ChromeDriver`: The version must correspond to your installed Google Chrome version. Download here.

## Install Python Dependencies:
It's recommended to use a virtual environment.

### Create and activate a virtual environment (optional but recommended)
```
python -m venv venv
```

```
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
```
## Install the required libraries
```
pip install pandas selenium
```

Setup ChromeDriver:

- Download the correct version of ChromeDriver for your operating system and Chrome browser version.

- Place the chromedriver executable in the same directory as the Python script, or add its location to your system's PATH.

## Usage
- **Review Configuration** :
\tOpen the main.py script and check the configuration variables at the top. You can change BASE_URL if the website address changes or OUTPUT_FILENAME to name the output file differently.

## Configuration
`BASE_URL` = "https://bids.delaware.gov"
`OUTPUT_FILENAME` = 'delaware_all_bids.xlsx'

**Run the Script:** 
Execute the script from your terminal:

`python main.py`

**Output:** 
The script will print its progress in the console, indicating which status it is currently processing and when it navigates to new pages. Once completed, it will generate an Excel file named Delaware_Bids_Export.xlsx (or your custom name) in the same directory.

## Project Structure
The script is organized into several key functions:

- `initialize_driver()`: Sets up and returns the Selenium Chrome WebDriver.

- `navigate_to_website()`: Navigates to the target URL.

- `extract_bid_details()`: Focuses on scraping data from the detailed modal pop-up for a single bid.

- `process_bids_for_status()`: The core function that handles a specific bid status tab. It iterates through all pages for that status, extracts summary data from the table, triggers the detailed extraction, and handles pagination.

- `save_to_excel()`: Converts the final list of scraped data into a pandas DataFrame and saves it to an Excel file.

- `main()`: The main execution function that orchestrates the entire process from initialization to saving the final file.
