from bs4 import BeautifulSoup
import requests
import json
import xlwt
from xlwt import Workbook

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')

response = requests.get("https://shop.adidas.jp/products/HB9386/")
soup = BeautifulSoup(response.content, 'html.parser')

bread_crumb_category = ""
for bread_crumb_item in soup.select(".breadcrumb_wrap .breadcrumbList .breadcrumbListItem:not(:first-child)"):
    bread_crumb_category = bread_crumb_category + bread_crumb_item.get_text(strip=True) + "/"

sheet1.write(0, 0, "Bread crumb category")
sheet1.write(1, 0, bread_crumb_category)

image_urls = ""
for item in soup.select(".slider-frame .slider-list .slider-slide"):
    image_urls = image_urls + item.find("img").get("src") + "\n"

sheet1.write(0, 1, "Side image urls")
sheet1.write(1, 1, image_urls)


article_purchase = soup.select(".articlePurchaseBox")[0]
article_info = article_purchase.select(".articleInformation")[0]
article_header = article_info.select(".articleNameHeader")[0]

category = article_header.select(".articleOtherLabel")[0].get_text(strip=True) + \
           " " + article_header.select(".groupName")[0].get_text(strip=True)
sheet1.write(0, 2, "Category")
sheet1.write(1, 2, category)


product_name = article_header.select("h1.itemTitle")[0].get_text(strip=True)
sheet1.write(0, 3, "Product Name")
sheet1.write(1, 3, product_name)

pricing = article_info.select("div.articlePrice")[0].get_text(strip=True)
sheet1.write(0, 4, 'Pricing')
sheet1.write(1, 4, pricing)

add_to_cart_form = article_purchase.select(".addToCartForm")[0]

available_size = ""
for size_li in add_to_cart_form.select("ul li"):
    text = size_li.get_text(strip=True)
    if text:
        available_size = available_size + text + " | "
sheet1.write(0, 5, 'Available Size')
sheet1.write(1, 5, available_size)

sense_of_the_size = ""
for sense_item in add_to_cart_form.select(".sizeFitBar .label span"):
    sense_of_the_size = sense_of_the_size + sense_item.get_text(strip=True) + " | "
sheet1.write(0, 6, 'Sense of the size')
sheet1.write(1, 6, sense_of_the_size)

script_data = soup.find(id="__NEXT_DATA__").get_text(strip=True)
product_data = json.loads(script_data)["props"]["pageProps"]["apis"]\
    ["pdpInitialProps"]["detailApi"]["product"]
coordinates_article = product_data["article"]["coordinates"]['articles']
coordinated_product = ""
for article in coordinates_article:
    product_name = article.get("name")
    pricing = article["price"]["current"]["withTax"]
    product_number = article["articleCode"]
    image_url = article.get("image")
    product_page_url = f"https://shop.adidas.jp/products/{product_number}/"
    coordinated_product += product_name + "\n" + pricing + "\n" + \
                           product_number + "\n" + image_url + \
                           "\n" + product_page_url + "\n\n"
sheet1.write(0, 7, 'Coordinated product')
sheet1.write(1, 7, coordinated_product)

article_promotion = soup.select("div.pdpContainer .articlePromotion")[0]
title_of_description = article_promotion.select("h4.heading")[0].get_text(strip=True)
sheet1.write(0, 8, 'Title of description')
sheet1.write(1, 8, title_of_description)

general_description_of_product = article_promotion.select(
    "div .description .details .commentItem-mainText")[0].get_text(strip=True)
sheet1.write(0, 9, 'General description of product')
sheet1.write(1, 9, general_description_of_product)

general_description_itemization = ""
for item in article_promotion.select("div .description .articleFeatures li"):
    general_description_itemization += item.get_text(strip=True) + "\n"
sheet1.write(0, 10, 'General description itemization')
sheet1.write(1, 10, general_description_itemization)

model_code = soup.find(id="vs-product-id").get("value")
size_chart = requests.get(f"https://shop.adidas.jp/f/v1/pub/size_chart/{model_code}")
header_data = size_chart.json().get("size_chart")['0'].get("header")['0']
body_data = size_chart.json().get("size_chart")['0'].get("body")

table_size_information = ""
for header_index in range(len(header_data)):
    record = ""
    header_val = header_data[f'{header_index}']['value']
    record += header_val + "|"
    for body_index in range(len(body_data)):
        body_val = body_data[f'{body_index}'][f'{header_index}']['value']
        record += body_val + "|"
    record += "\n"
    table_size_information += record
sheet1.write(0, 11, 'Table size information')
sheet1.write(1, 11, table_size_information)

review = product_data["model"]["review"]
reviewSeoLd = review["reviewSeoLd"]

rating = review["fitbarScore"]
number_of_reviews = review["reviewCount"]
recommended_rate = str(review["ratingAvg"])+"%"

review_rating_rate = str(rating) + "\n" + str(number_of_reviews) + "\n" + str(recommended_rate) + "\n"
sheet1.write(0, 12, 'Review/rating/Rate')
sheet1.write(1, 12, review_rating_rate)


review_info = ""
for review in reviewSeoLd:
    date = review["datePublished"]
    rating = review["reviewRating"]["ratingValue"]
    review_title = review["name"]
    review_descriptions = review["reviewBody"]
    review_info += date + "\n" + rating + "\n" + review_title + "\n" + review_descriptions + "\n\n"

sheet1.write(0, 13, 'Review information')
sheet1.write(1, 13, review_info)

wb.save('output.xls')



