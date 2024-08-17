import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwt
import traceback

def fetch_products(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.content
    else:
        return

def get_link_product(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    products = soup.find_all('div', class_='product')
    link_data = []
    for product in products:
        link = product.find('h3', class_='woocommerce-loop-product__title').find('a').get('href')
        link_data.append({'url': link})
    return link_data

def fetch_all_products():
    all_products = []
    page = 0

    while True:
        try:
            page += 1
            print(f"Fetching page {page}...")
            html_content = fetch_products(f"https://www.lusicas.com/product-category/hermes/hermes-bags/page/{page}?only_posts=1")
            # html_content = fetch_products(f"https://www.lusicas.com/product-category/best-selling/page/3?only_posts=1")
            link_products = get_link_product(html_content)
            for link_product in link_products:
                title, sale, categories, sku_code, price, image, sub_image, description_html, attributes_str = get_info_product(link_product.get('url'))
                all_products.append({'sku': sku_code, 'image': image, 'branch_id': '', 
                                    'supplier_id': '', 'price': price, 'cost': price,
                                    'stock': 0, 'minimum': 1, 'weight_class': 'g',
                                    'weight': 0, 'length_class': 'cm', 'length': 0,
                                    'width': 0, 'height': 0, 'kind': 0,
                                    'tax_id': 0, 'status': 1, 'alias': sku_code,
                                    'sort': 0, 'upc': '', 'ean': '',
                                    'jan': '', 'isbn': '', 'mpn': '',
                                    'categories': categories, 'sub-images': sub_image, 'store_id': 1,
                                    'sale': sale, 'name': title, 'content': description_html, 'attributes': attributes_str})
        except Exception as e:
            print(f"An error occurred: {e}")
            traceback.print_exc()
            break        

    return all_products

def get_info_product(url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')

        # get title product
        title = soup.find('h1', class_='product_title').text

        # get sale price
        sale_str = soup.find('div', class_='woocommerce-product-gallery').find('label', class_='label-sale')
        sale = 0
        if sale_str:
            sale = sale_str.text.split('%')[0]

        # get list category
        categories_crawl = soup.find('span', class_='posted_in').find_all('a')
        categories = ''
        count = 0
        for category in categories_crawl:
            count +=1
            categories += category.text
            if count < len(categories_crawl):
                categories += ','

        # get sku code
        sku_code = soup.find('span', class_='sku').text.replace("\t", "").replace("\n", "").replace("\r", "").replace('LUX', 'LXNL')

        # get price
        price_html = soup.find('p', class_='price').find_all('span', class_='woocommerce-Price-amount')
        if (len(price_html) == 1):
            price = price_html[0].text.split("$")[1]
        else:
            price = price_html[1].text.split("$")[1]
        
        # get image
        images_html = soup.find_all('div', class_='woocommerce-product-gallery__image')
        count_img = 0
        image = ''
        sub_image = ''
        for image_html in images_html:
            count_img += 1
            response = requests.get(image_html.get('data-thumb'))
            if (count_img == 1):
                image = f"/data/product/{sku_code}_{count_img}.jpg"
            else:
                sub_image += f"/data/product/{sku_code}_{count_img}.jpg"

            if (count_img < len(images_html)):
                sub_image += ','

            if response.status_code == 200:
                with open(f"images/{sku_code}_{count_img}.jpg", 'wb') as file:
                    file.write(response.content)
        
        # get image size
        image_size_url = ''
        if soup.find('div', class_='woocommerce-product-details__short-description') is not None:
            if soup.find('div', class_='woocommerce-product-details__short-description').find('img'):
                image_size_url = soup.find('div', class_='woocommerce-product-details__short-description').find('img').get('src')
                response = requests.get(image_size_url)
                if response.status_code == 200:
                        sub_image += f",/data/product/{sku_code}_{count_img+1}.jpg"
                        with open(f"images/{sku_code}_{count_img+1}.jpg", 'wb') as file:
                            file.write(response.content)

        # get description
        description_html = ''
        if soup.find('div', class_='woocommerce-Tabs-panel--description') is not None:
            description_html = soup.find('div', class_='woocommerce-Tabs-panel--description').decode_contents()

        # get attributes
        attributes_str = ''
        if soup.find('div', class_='product-pa_shoes-size-swatch') is not None:
            attributes_html = soup.find('div', class_='product-pa_shoes-size-swatch').find_all('option')[1:-1]
            count_attribute = 0
            for attribute_html in attributes_html:
                count_attribute += 1
                attributes_str += attribute_html.text

                if (count_attribute < len(attributes_html)):
                    attributes_str += ','

        return title, sale, categories, sku_code, price, image, sub_image, description_html, attributes_str

    else:
        return False

# get_info_product('https://www.lusicas.com/product/louis-vuitton-lock-it-flat-mule-monogram-brown/')
# Fetch all products
all_product_data = fetch_all_products()

# Save to Excel
df = pd.DataFrame(all_product_data)

# Create a workbook and add a worksheet
workbook = xlwt.Workbook('products.xls')
worksheet = workbook.add_sheet('sheet1',cell_overwrite_ok=True)

# Create dictionary of formats for each column
number_format = xlwt.easyxf(num_format_str='0.00')
cols_to_format = {0:number_format}

for z, value in enumerate(df.columns):
    worksheet.write(0, z, value)

# Iterate over the data and write it out row by row
for x, y in df.iterrows():
    for z, value in enumerate(y):
        if z in cols_to_format.keys():
            worksheet.write(x + 1, z, value, cols_to_format[z])
        else: ## Save with no format
             worksheet.write(x + 1, z, value)

# Save/output the workbook
workbook.save('products.xls')

print("Data has been successfully saved to products.xls")