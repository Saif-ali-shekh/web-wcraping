import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import re
import signal

# Define a function to handle interruption and save data
def save_and_exit(signal, frame):
    global data, chunk_count
    if data:
        save_to_excel(data, f'before_exit_{chunk_count}.xlsx')
    # driver.quit()
    print("Data saved and browser closed.")
    exit(0)

signal.signal(signal.SIGINT, save_and_exit)

# Define a function to scroll and fetch data
def fetch_college_data(driver, num_chunks=200):
    global data, chunk_count
    data = []
    last_height = driver.execute_script("return document.body.scrollHeight")
    chunk_count = 0
    total_scrolls = 0

    while True:
        try:
            # Scroll down in chunks
            for _ in range(num_chunks):
                driver.execute_script("window.scrollBy(0, window.innerHeight);")
                time.sleep(1)  # Wait to load the page
                print("itr",_)
                total_scrolls += 1

            # Extract HTML and parse it with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            college_list = soup.find_all('tbody', class_="jsx-4033392124 jsx-1933831621")  # tbody fetched

            for college_group in college_list:
                colls = [tr for tr in college_group.find_all('tr', recursive=False)]  # all colleges

                for college in colls:
                    try:
                        university_url = college.find('div', class_="jsx-3749532717 clg-name-address").find('a', class_="jsx-3749532717 college_name underline-on-hover")['href']
                        university_full_url = 'https://collegedunia.com' + university_url
                        university_name = college.find('div', class_="jsx-3749532717 clg-name-address").find('a', class_="jsx-3749532717 college_name underline-on-hover").text.strip()
                        un_name_split = re.split(r'[-,]', university_name)
                        if len(un_name_split) > 2:
                            university_name = un_name_split[1]
                        elif len(un_name_split) == 2:
                            university_name = un_name_split[0]

                        college_type = university_url.split('/')[1].strip()
                        course_fees = college.find('td', class_='jsx-3749532717 col-fees').find('span').text.strip()
                        rating = college.find('td', class_='jsx-3749532717 col-reviews').find('span').text.strip()
                        college_ranking = college.find('td', class_='jsx-3749532717 col-ranking').find('span', class_='jsx-2794970405 rank-span no-break').text.strip()
                        city_state = college.find('div', class_='jsx-3749532717 clg-name-address').find('span', class_='jsx-3749532717 pr-1 location').text.strip()
                        city, state = city_state.split(', ')

                        data.append({
                            'University URL': university_full_url,
                            'University Name': university_name,
                            'College Type': college_type,
                            'Course Fees': course_fees,
                            'Rating': rating,
                            'College Ranking': college_ranking,
                            'City': city,
                            'State': state
                        })

                    except Exception as e:
                        print(f"Error parsing college data: {e}")

            # Save data to an Excel file after every chunk
            save_to_excel(data, f'chunk_{chunk_count + 1}.xlsx')
            data = []
            chunk_count += 1
            print(f'Chunk {chunk_count} saved.')

            # Check if the end of the page is reached
            time.sleep(2)  # Wait to load the page
            
            new_height = driver.execute_script("return document.body.scrollHeight")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        except Exception as e:
            print(f"Error during scrolling or data extraction: {e}")

    print(f"Total scrolls: {total_scrolls}")

# Save data to an Excel file
def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    try:
        existing_df = pd.read_excel(filename)
        df = pd.concat([existing_df, df]).drop_duplicates(subset=['University URL']).reset_index(drop=True)
    except FileNotFoundError:
        pass
    df.to_excel(filename, index=False)

# Set up Selenium WebDriver
driver_path = r'C:\Users\Administrator\Downloads\chromedriver-win64\chromedriver.exe'

# Initialize the WebDriver with Service object
service = Service(driver_path)
driver = webdriver.Chrome(service=service)
driver.get('https://collegedunia.com/india-colleges')

# Fetch and save college data
fetch_college_data(driver)

# Close the browser
driver.quit()
