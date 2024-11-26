import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
# Define function to get longest and shortest options
def get_longest_and_shortest_options(keyword):
    # Initialize WebDriver (ensure you set the correct path to your WebDriver)
    driver = webdriver.Chrome(executable_path=r"C:/Users/Navid\Downloads/chromedriver-win64")
    driver.get("https://www.google.com")
    
    # Find the search box and enter the keyword
    search_box = driver.find_element(By.NAME, "q")
    search_box.clear()
    search_box.send_keys(keyword)
    search_box.send_keys(Keys.RETURN)
    
    time.sleep(2)  # Wait for the results to load
    
    # Extract search suggestions from Google's dropdown (if available)
    try:
        suggestions = driver.find_elements(By.CSS_SELECTOR, "li span")
        options = [s.text for s in suggestions if s.text.strip()]
    except Exception as e:
        options = []
    
    driver.quit()
    
    if options:
        # Find the longest and shortest options
        longest_option = max(options, key=len)
        shortest_option = min(options, key=len)
        return longest_option, shortest_option
    else:
        return None, None

# Define the main function
def process_excel_and_update(file_path):
    # Get the current day of the week
    current_day = datetime.now().strftime('%A')  # E.g., 'Monday'
    
    # Load the Excel file
    excel_data = pd.ExcelFile(file_path)
    
    # Check if the sheet for the current day exists
    if current_day not in excel_data.sheet_names:
        print(f"No data available for {current_day}")
        return
    
    # Read the sheet for the current day
    day_data = excel_data.parse(sheet_name=current_day)
    
    # Prepare a new DataFrame to store results
    results = []
    for index, row in day_data.iterrows():
        keyword = row.get("Keyword", "")
        if keyword:
            print(f"Processing keyword: {keyword}")
            longest, shortest = get_longest_and_shortest_options(keyword)
            results.append({
                "Keyword": keyword,
                "Longest Option": longest,
                "Shortest Option": shortest
            })
    
    # Write the results back to the Excel file
    results_df = pd.DataFrame(results)
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        results_df.to_excel(writer, sheet_name=f"{current_day}_Results", index=False)
    
    print("Processing complete!")

# Path to your Excel file
file_path = "C:/Users/Navid/Downloads/Script.xlsx"
# Call the main function
process_excel_and_update(file_path)

