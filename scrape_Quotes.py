from bs4 import BeautifulSoup
import requests
import openpyxl

# Initialize Excel workbook and sheet
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Quotes'
sheet.append(['Quote', 'Author', 'Tags'])

# Define headers with a User-Agent to simulate a browser request
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

try:
    # Fetch the quotes page with headers
    source = requests.get('http://quotes.toscrape.com', headers=headers)
    source.raise_for_status()  # Check if request was successful

    # Parse HTML with BeautifulSoup
    soup = BeautifulSoup(source.text, 'html.parser')
    quotes = soup.find_all('div', class_='quote')

    # Extract and print quote details, adding them to the Excel sheet
    for quote in quotes:
        text = quote.find('span', class_='text').text
        author = quote.find('small', class_='author').text
        tags = ', '.join(tag.text for tag in quote.find_all('a', class_='tag'))

        print(text, author, tags)  # Print details for verification
        sheet.append([text, author, tags])

except Exception as e:
    print(f"An error occurred: {e}")

# Save the workbook to a file
excel.save('Quotes.xlsx')