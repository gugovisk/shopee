from bs4 import BeautifulSoup
from openpyxl import Workbook

# Read the HTML file
with open('C:/Users/User/workspace/gustavo/projetos/shoppe/coleta/Shopee10.html', 'r', encoding='utf-8') as file:
    html_content = file.read()

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Create a Workbook and select the active worksheet
wb = Workbook()
ws = wb.active
ws.title = "Shopee Products"

# Define the headers for the Excel file
headers = ["Product_Name", "Product_Price", "Off", "Product_Rating", "Solds", "Local"]
ws.append(headers)

# Find all relevant divs
divs = soup.find_all('div', class_='flex flex-col bg-white cursor-pointer h-full')

# Loop through each div and extract the necessary information
for div in divs:
    # Try to find the Product Name
    product_name_tag = div.find('div', class_='whitespace-normal line-clamp-2 break-words min-h-[2.5rem] text-sm')
    product_name = product_name_tag.get_text(strip=True) if product_name_tag else " "
    
    # Try to find the Product Price
    product_price_tag = div.find('span', class_='font-medium text-base/5 truncate')
    product_price = product_price_tag.get_text(strip=True) if product_price_tag else " "

     # Try to find the Off
    off_tag = div.find('div', class_='text-shopee-primary font-medium bg-shopee-pink py-0.5 px-1 text-sp10/3 h-4 rounded-[2px] shrink-0 mr-1')
    off = off_tag.get_text(strip=True) if off_tag else " "
    
    # Try to find the Product Rating
    product_rating_tag = div.find('div', class_='text-shopee-black87 text-xs/sp14 flex-none')
    product_rating = product_rating_tag.get_text(strip=True) if product_rating_tag else " "

    # Try to find the Solds
    solds_tag = div.find('div', class_='truncate text-shopee-black87 text-xs min-h-4')
    solds = solds_tag.get_text(strip=True) if solds_tag else " "

    # Try to find the Local
    local_tag = div.find('span', class_='ml-[3px] align-middle')
    local = local_tag.get_text(strip=True) if local_tag else " "
    
    # Write the extracted information to the Excel file
    ws.append([product_name, product_price, off, product_rating, solds, local])

# Save the workbook to a file
output_file = 'C:/Users/User/workspace/gustavo/projetos/shoppe/coleta/Shopee_products10.xlsx'
wb.save(output_file)

print(f"Data saved to {output_file}")
