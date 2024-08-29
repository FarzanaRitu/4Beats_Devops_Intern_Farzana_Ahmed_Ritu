import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Load the Excel file and get today's sheet
file_path = 'excel_file.xlsx'  # Replace with your file path
today = datetime.today().strftime('%A')  # E.g., 'Sunday', 'Monday', etc.
excel_file = pd.ExcelFile(file_path)

if today in excel_file.sheet_names:
    # Skip the first row and use the second row as header
    df = pd.read_excel(file_path, sheet_name=today, skiprows=1)
    print(f"Opened sheet: {today}")
else:
    print(f"No sheet found for today: {today}")
    raise ValueError(f"No sheet found for {today} in {file_path}")

# Display the first few rows to understand the structure
print("First few rows of the corrected sheet:")
print(df.head())

# Rename the columns properly, keeping the first two columns intact and renaming the rest
df.columns = ['Index', 'Keyword', 'Search Term', 'Longest Option', 'Shortest Option']

# Set up Selenium
driver = webdriver.Chrome()
driver.get("http://www.google.com")
print("ChromeDriver started successfully.")

# Process each search term in the DataFrame
for i, search_term in enumerate(df['Search Term']):
    try:
        # Wait for the search box to be present
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
        search_box.clear()
        search_box.send_keys(search_term)
        time.sleep(3)  # Wait for autocomplete options to load
        
        # Capture autocomplete suggestions
        suggestions = driver.find_elements(By.XPATH, "//li[@role='presentation']//span")
        options = [suggestion.text for suggestion in suggestions if suggestion.text.strip()]
        
        if options:
            # Identify the shortest and longest options by character count
            shortest = min(options, key=len)
            longest = max(options, key=len)
            
            # Update the DataFrame
            df.at[i, 'Shortest Option'] = shortest
            df.at[i, 'Longest Option'] = longest
            
            print(f"Search Term: {search_term} | Shortest: {shortest} | Longest: {longest}")
        else:
            print(f"No suggestions found for search term: {search_term}")

    except Exception as e:
        print(f"An error occurred while processing {search_term}: {e}")

# Save the updated DataFrame back to the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=today, index=False)
    print(f"Updated data saved to sheet '{today}' in {file_path}")

# Closing the browser
driver.quit()
