from playwright.sync_api import sync_playwright
from lxml import html
import time
import pandas as pd  # For saving data to Excel
import re
import os

# Function to scrape Google Maps data for a specific city
def scrape_google_maps(city_name):
    # Initialize a list to store the extracted data
    extracted_data = []
    processed_urls = set()  # Track processed URLs

    with sync_playwright() as p:
        # Launch Edge browser
        browser = p.chromium.launch(
            executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            headless=False
        )
        page = browser.new_page()

        # Modify the URL dynamically for each city
        search_url = f"https://www.google.com/maps/search/destinasi+wisata+di+{city_name.lower()}"

        # Retry logic for page.goto
        for attempt in range(3):  # Try up to 3 times
            try:
                print(f"Attempting to navigate to {search_url} (Attempt {attempt + 1})...")
                page.goto(search_url, timeout=30000)  # Set timeout to 30 seconds
                break  # Exit the loop if successful
            except TimeoutError:
                print(f"Timeout while loading {search_url}. Retrying...")
                if attempt == 2:  # On the third attempt, raise the error
                    raise

        # Wait for the body to load
        page.wait_for_selector("body")
        time.sleep(5)

        # Step 1: Change the language to English (US)
        try:
            menu_button = page.query_selector("//button[@aria-label='Menu']")
            if menu_button:
                print("Opening the menu...")
                menu_button.click()
                time.sleep(2)

                language_button = page.query_selector('//button[@class="aAaxGf T2ozWe"]')
                if language_button:
                    print("Opening the language settings...")
                    language_button.click()
                    time.sleep(1)

                    english_option = page.query_selector('//a[contains(@href, "hl=en")]')
                    if english_option:
                        print("Selecting English...")
                        english_option.click()
                        time.sleep(7)
                    else:
                        print("English option not found. Skipping language change.")
                else:
                    print("Language button not found. Skipping language change.")
            else:
                print("Menu button not found. Skipping language change.")
        except Exception as e:
            print("Error during language change:", e)

        # Step 3: Loop through clickable items and extract data
        try:
            while True:
                clickable_items = page.query_selector_all("//div[@role='feed']//a[@aria-label and starts-with(@href, 'https://www.google.com/maps')]")
                if not clickable_items:
                    print("No more clickable items found.")
                    break

                for index, item in enumerate(clickable_items):
                    href = item.get_attribute("href")
                    if href in processed_urls:
                        # Skip already processed items
                        continue

                    print(f"Clicking item {index + 1}...")
                    item.click()
                    processed_urls.add(href)

                    # Wait for content to load
                    time.sleep(3)

                    # Extract page content
                    page_source = page.content()
                    tree = html.fromstring(page_source)

                    # Extract coordinates from the URL
                    current_url = page.url
                    match = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", current_url)
                    if match:
                        latitude = match.group(1)
                        longitude = match.group(2)
                    else:
                        latitude = None
                        longitude = None

                    # Extract the desired data using safer queries
                    place = tree.xpath("//h1[@class='DUwDvf lfPIob']/text()")
                    overview = tree.xpath("//div[@class='y0K5Df']//div[contains(@class,'PYvSYb')]/text()")
                    category = tree.xpath("//button[contains(@jsaction, 'category')]/text()")
                    address = tree.xpath("//button[@class='CsEnBe' and contains(@aria-label,'Address')]//div[contains(@class,'Io6YTe')]/text()")
                    price = tree.xpath("(//div[@class='drwWxc'])[1]/text()")
                    rating = tree.xpath("//div[contains(@class,'F7nice')]//span[contains(@aria-hidden,'true')]/text()")
                    review_count = tree.xpath('//span[contains(@aria-label, "reviews")]/text()')

                    # Safely handle empty lists and set defaults
                    place = place[0] if place else None
                    overview = overview[0] if overview else None
                    category = category[0] if category else None
                    address = address[0] if address else None
                    price = price[0] if price else "Free"  # Default to "free" if no price is found
                    rating = rating[0] if rating else None
                    review_count = review_count[0] if review_count else None

                    # If price is "free", set it to 0
                    price = 0 if price.lower() == "Free" else price

                    # Clean and structure the extracted data
                    data_entry = {
                        "Place": place,
                        "Overview": overview,
                        "Category": category,
                        "Address": address,
                        "City": city_name,
                        "Price": price,
                        "Rating": rating,
                        "Review_count": review_count,
                        "Latitude": latitude,
                        "Longitude": longitude,
                    }
                    print(f"Extracted data: {data_entry}")

                    # Add to extracted data list
                    extracted_data.append(data_entry)

                    time.sleep(1)

        except Exception as e:
            print("Error during clicking or extracting data:", e)

        # Clean up
        browser.close()

    # Step 4: Save the extracted data to an Excel file
    if extracted_data:
        # Convert the extracted data to a DataFrame
        new_data_df = pd.DataFrame(extracted_data)
        file_path = f"dataset_{city_name.lower()}.xlsx"

        # Check if the file already exists
        if os.path.exists(file_path):
            print(f"File '{file_path}' already exists. Appending new data...")
            existing_data_df = pd.read_excel(file_path)
            # Append the new data and remove duplicates
            combined_data_df = pd.concat([existing_data_df, new_data_df], ignore_index=True).drop_duplicates()
        else:
            print(f"File '{file_path}' does not exist. Creating a new file...")
            combined_data_df = new_data_df

        # Save the combined data to the file
        combined_data_df.to_excel(file_path, index=False)
        print(f"Data saved to '{file_path}'.")


# Step 5: Read cities from Excel and run the scraper for each city
city_data = pd.read_excel('kota_di_indonesia.xlsx')  # Replace with the correct file path

# Assuming the city names are in a column named 'Kota'
city_list = city_data['Kota'].tolist()

# Loop through the cities and scrape data for each
for city_name in city_list:
    print(f"Scraping data for {city_name}...")
    city_name = city_name.strip()  # Remove any leading/trailing spaces
    scrape_google_maps(city_name)  # Pass the city_name as an argument
