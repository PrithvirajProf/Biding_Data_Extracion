import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- Configuration ---
BASE_URL = "https://bids.delaware.gov"
OUTPUT_FILENAME = 'Delaware_Bids_Export.xlsx'

# --- Core Functions ---

def initialize_driver():
    """This function
    Initializes and returns a Chrome WebDriver instance.
    """
    print("Initializing Chrome WebDriver...")
    return webdriver.Chrome()

def navigate_to_website(driver, url):
    """Navigates the driver to the specified URL."""
    print(f"Navigating to {url}...")
    driver.get(url)
    # Wait for the page to load, adjust time as needed
    time.sleep(3)

def extract_bid_details(driver):
    """
    Extracts specific details from the bid detail pop-up.
    """
    details = {}
    try:
        # Wait for the modal header to be visible to ensure the pop-up has loaded
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".modal-header h2"))
        )
        
        # --- Scrape individual data points ---
        # The selectors below are examples. They must be updated to match the actual website's structure.
        details['header'] = driver.find_element(By.CSS_SELECTOR, ".modal-header h2").text
        details['solicitation_ad_date'] = driver.find_element(By.XPATH, "//div[contains(@class, 'solicitation')]").text
        details['deadline'] = driver.find_element(By.XPATH, "//div[contains(@class, 'deadline')]").text
        details['contact_name'] = driver.find_element(By.XPATH, "//span[contains(@class, 'contact-name')]").text
        details['contact_email'] = driver.find_element(By.XPATH, "//a[contains(@class, 'contact-email')]").text
        
        # Extract all document links available in the modal
        docs = driver.find_elements(By.CSS_SELECTOR, ".document-link a")
        details['document_links'] = [doc.get_attribute('href') for doc in docs]

    except (NoSuchElementException, TimeoutException) as e:
        print(f"Could not extract all details from modal: {e}")
    
    return details

def process_bids_for_status(driver, status, all_bids_list):
    """
    Clicks on a bid status tab (e.g., 'Open') and scrapes all bids from all pages for that status.

        driver: The Selenium WebDriver instance.
        status (str): The name of the status tab to process (e.g., 'Open', 'Awarded').
        all_bids_list (list): The list where scraped bid data will be appended.
    """
    print(f"\n--- Processing bids for status: {status} ---")
    try:
        # Find and click the corresponding status tab
        tab_xpath = f"//a[contains(text(), '{status}')]"
        tab = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, tab_xpath)))
        tab.click()
        # Wait for the table content to refresh
        time.sleep(2)
    except TimeoutException:
        print(f"Could not find or click the '{status}' tab.")
        return

    # Loop through all pages for the current status
    while True:
        # Wait for the main bids table to be present
        try:
            table_selector = "table.bids-table tbody"
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, table_selector)))
            rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} tr")
            print(f"Found {len(rows)} bids on the current page.")
        except TimeoutException:
            print("Bid table not found on the page.")
            break

        # Process each row in the table
        for row in rows:
            try:
                cols = row.find_elements(By.TAG_NAME, "td")
                # Ensure the row has enough columns to avoid errors
                if len(cols) < 6:
                    continue

                # Scrape data from the main table
                bid_data = {
                    'Contract Number': cols[0].text,
                    'Contract Title': cols[1].text,
                    'Open Date': cols[2].text,
                    'Deadline Date': cols[3].text,
                    'Agency Code': cols[4].text,
                    'UNSPSC': cols[5].text,
                    'Status': status
                }
                
                # Click the row to open the details modal
                row.click()
                time.sleep(1) # Allow modal to open
                
                # Extract detailed data from the modal and merge it
                detail_data = extract_bid_details(driver)
                bid_data.update(detail_data)
                
                # Close the modal to return to the table
                # This selector needs to be accurate for the modal's close button
                close_btn = driver.find_elements(By.CSS_SELECTOR, ".modal-close")
                if close_btn:
                    close_btn[0].click()
                    time.sleep(1) # Allow modal to close
                
                all_bids_list.append(bid_data)

            except Exception as e:
                print(f"Error processing a row: {e}")
                # Try to close a potentially stuck modal before continuing
                try:
                    close_btn = driver.find_elements(By.CSS_SELECTOR, ".modal-close")
                    if close_btn:
                        close_btn[0].click()
                except Exception as close_e:
                    print(f"Could not close stray modal: {close_e}")


        # Check for a "Next" page button that is not disabled
        try:
            next_page_selector = "a.page-link:not(.disabled)" # Example selector for a clickable next button
            next_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, next_page_selector)))
            next_btn.click()
            print("Navigating to the next page...")
            time.sleep(2) # Wait for next page to load
        except (TimeoutException, NoSuchElementException):
            print("No more pages found for this status.")
            break # Exit the while loop if no "Next" button is found

def save_to_excel(data, filename):
    """
    Converts a list of dictionaries to a pandas DataFrame and saves it as an Excel file.

    Args:
        data (list): A list of dictionaries, where each dictionary is a bid.
        filename (str): The name of the output Excel file.
    """
    if not data:
        print("No data was collected to export.")
        return
        
    print(f"\nExporting {len(data)} bids to {filename}...")
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print("Data export complete.")


# --- Main Execution ---

def main():
    """
    Main function to orchestrate the web scraping process.
    """
    # List to store data from all bid statuses
    all_bids_data = []
    driver = None  # Initialize driver to None

    try:
        # Initialize WebDriver and navigate to the website
        driver = initialize_driver()
        navigate_to_website(driver, BASE_URL)
        
        # Define the statuses to scrape
        statuses_to_process = ['Open', 'Recently Closed', 'Awarded']
        
        # Loop through each status and scrape the data
        for status in statuses_to_process:
            process_bids_for_status(driver, status, all_bids_data)
        
        # Save all collected data to an Excel file
        save_to_excel(all_bids_data, OUTPUT_FILENAME)

    except Exception as e:
        print(f"\nAn unexpected error occurred in the main process: {e}")

    finally:
        # Ensure the WebDriver is closed properly
        if driver:
            print("Closing WebDriver.")
            driver.quit()

# Entry point for the script
if __name__ == "__main__":
    main()