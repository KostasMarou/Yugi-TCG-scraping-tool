#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import re
from collections import defaultdict

# Load the data from the Excel file  FIX PATH TO YOUR FILE (the actual file is available on my Git-hub)
file_path = r"C:\Users\.........\fullYugiohListScrape.xlsx"
df = pd.read_excel(file_path)
card_names = df.iloc[:, 0].tolist()

# Set up Edge WebDriver
driver = webdriver.Edge()
driver.get("https://www.cardmarket.com/en/YuGiOh")
time.sleep(3)

# Dictionary to store the lowest-price sellers for each card
card_data = {}

# Function to clean card names by removing special characters
def clean_name(name):
    return re.sub(r"[^a-zA-Z0-9\s]", "", name)

# Loop through the first 10 cards in the card_names list
for card_name in card_names[:15]:
    cleaned_name = clean_name(card_name)
    print(f"Processing card: {card_name} (cleaned name: {cleaned_name})")
    
    try:
        # Find and use the search bar
        search_bar = driver.find_element(By.NAME, "searchString")
        search_bar.clear()
        search_bar.send_keys(card_name)
        search_bar.send_keys(Keys.RETURN)
        time.sleep(3)

        # Retry finding the link a couple of times if not successful on the first try
        for _ in range(3):
            try:
                card_link = driver.find_element(By.XPATH, f"//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{cleaned_name.lower()}')]")
                card_link.click()
                print(f"Clicked on the link for card: {card_name}")
                time.sleep(3)
                break
            except:
                print(f"Retry finding link for card: {card_name}")
                time.sleep(2)
        else:
            print(f"Could not find a link for card: {card_name}")
            continue

        # Find and click "Show Offers" if available
        try:
            show_offers_link = driver.find_element(By.LINK_TEXT, "Show Offers")
            show_offers_link.click()
            print("Clicked on 'Show Offers' link.")
            time.sleep(3)
        except:
            print("Could not find the 'Show Offers' link.")
            continue

        # Initialize list to store lowest-price sellers for this card
        sellers_prices = []
        lowest_price = None  # Variable to track the lowest price for this card
        
        # Find all rows with the 'article-row' class
        rows = driver.find_elements(By.CLASS_NAME, "article-row")
        for row in rows:
            try:
                seller = row.find_element(By.CSS_SELECTOR, ".col-sellerProductInfo .seller-name span a").text
                price_text = row.find_element(By.CSS_SELECTOR, ".price-container .fw-bold").text
                price = float(price_text.replace(",", ".").replace(" €", ""))  # Convert price to float
                
                # Check if this price is the lowest for the current card
                if lowest_price is None or price < lowest_price:
                    # Found a new lowest price, reset the list
                    lowest_price = price
                    sellers_prices = [(seller, price_text)]
                elif price == lowest_price:
                    # Same as current lowest price, add this seller
                    sellers_prices.append((seller, price_text))
            except Exception as e:
                print(f"Error extracting data from row: {e}")

        # Save data for the current card
        card_data[card_name] = sellers_prices
        print(f"Data collected for {card_name}: {sellers_prices}")

        # Go back to the homepage to search for the next card
        driver.get("https://www.cardmarket.com/en/YuGiOh")
        time.sleep(3)

    except Exception as e:
        print(f"An error occurred while processing card '{card_name}': {e}")



# Step 2: Identify the seller with the most low-priced cards
seller_count = defaultdict(int)  # Dictionary to count the occurrences of each seller

# Count each seller's appearance across all cards
for card, sellers in card_data.items():
    for seller, price in sellers:
        seller_count[seller] += 1

# Find the seller with the most low-priced cards
best_seller = max(seller_count, key=seller_count.get)
print(f"Seller with the most low-priced cards: {best_seller}")

# Step 3: Create a list of all cards and prices for the best seller
best_seller_cards = [(card, price) for card, sellers in card_data.items() for seller, price in sellers if seller == best_seller]
print(f"Cards sold by {best_seller} at the lowest price: {best_seller_cards}")


#NOW SAVING DATA IN A TABLE [FIX PATH]
best_seller_df = pd.DataFrame(best_seller_cards, columns=["Card Name", "Price"])

# Define the file path for the output Excel file
output_file_path = r"C:\Users\...........\FilesName.xlsx"

# Save the DataFrame to an Excel file
best_seller_df.to_excel(output_file_path, index=False)

print(f"Excel file has been created at: {output_file_path}")


# In[ ]:




