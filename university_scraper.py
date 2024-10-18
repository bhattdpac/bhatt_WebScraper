# Import necessary libraries
import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

# Step 1: Request the web page
url = 'https://www.harvard.edu/programs/?degree_levels=graduate'
response = requests.get(url)

# Check if the request was successful
if response.status_code == 200:
    html_content = response.content
else:
    print(f"Failed to retrieve the page: {response.status_code}")
    exit()

# Step 2: Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Step 3: Find the program listings
# Assuming programs are in list items with class 'program-item'
programs = soup.find_all('li', class_='program-item')

# Step 4: Extract program data (Program name, school, URL)
program_data = []

for program in programs:
    try:
        # Extract program name from h3 tag with class 'program-title'
        program_name = program.find('h3', class_='program-title').get_text(strip=True)
        
        # Extract the department offering the program from div with class 'program-school'
        program_department = program.find('div', class_='program-school').get_text(strip=True)
        
        # Extract the URL to the program page from 'a' tag with class 'program-title-link'
        program_url = program.find('a', class_='program-title-link')['href']
        program_url = f"https://www.harvard.edu{program_url}"  # Construct full URL

    except AttributeError as e:
        # Print error message if data extraction fails for a program
        print(f"Error extracting data: {e}")
        continue

    # Append the extracted data to the list
    program_data.append([program_name, program_department, program_url])

# Step 5: Convert the data to a pandas DataFrame
df_programs = pd.DataFrame(program_data, columns=['Program Name', 'School/Department', 'Program URL'])

# Step 6: Create output directory if it doesn't exist
output_dir = 'output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Step 7: Save the data into an Excel file
output_path = os.path.join(output_dir, 'harvard_graduate_programs.xlsx')
df_programs.to_excel(output_path, index=False)

# Print confirmation message
print(f"Scraped data saved to {output_path}")