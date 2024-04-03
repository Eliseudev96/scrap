from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
from spreadsheet import saveInSpreadsheet, verifyProducts, verifyImage, cleanTranslate
from getInfoProducts import getInfoProduct
import urllib
import time

start_time = time.time()


def requestsNewProducts():
    print("Capturing products links...")

    # Criando a planilha antes do loop principal
    saveInSpreadsheet([])

    html = urlopen("https://www.vevor.de/")
    bs = BeautifulSoup(html, "html5lib")

    main_categories = bs.find_all(class_="headerCate_item")

    for main_category in main_categories:
        title = main_category.find(class_="headerCate_text")

        sub_categories = main_category.find(class_="headerCate_childColumn")
        categories = sub_categories.find_all(class_="headerCate_blockItem")

        for category in categories:
            anchor = category.find("a")
            category_link = "https://www.vevor.de{}".format(anchor["href"])
            category_link = urllib.parse.quote(category_link, safe=":/")

            # Processa cada categoria e salva diretamente na planilha
            processCategory(category_link)


def processCategory(category_link):
    html = urlopen(category_link)
    bs = BeautifulSoup(html, "html5lib")

    qtyPage = bs.find_all(class_="gPage_item")
    pl = bs.find_all(class_="compItem_topWrap")

    product_links = []

    for product in pl:
        anchor_product = product.find("a")
        product_links.append(anchor_product["href"])

    # Salva os links diretamente na planilha
    saveInSpreadsheet(product_links)

    if len(qtyPage) > 1:
        for i in range(2, len(qtyPage) + 1):
            link = category_link + "?page={}".format(str(i))
            req = Request(
                link,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.43"
                },
            )

            html = urlopen(req)
            bs = BeautifulSoup(html, "html5lib")

            pl = bs.find_all(class_="compItem_topWrap")

            for product in pl:
                anchor_product = product.find("a")
                product_links.append(anchor_product["href"])

            # Salva os links diretamente na planilha
            saveInSpreadsheet(product_links)


def updateProductsFromXLSX():
    print("ESTIMATED TIME: 1h30min\n")
    verifyProducts(2)


def fixWrongRows():
    for i in range(0, 4):
        verifyImage()


print("[1] - Extract new products (products.xlsx must be empty!)\n")
print("[2] - Update products\n")
result = int(input("> "))

if result == 1:
    requestsNewProducts()
    print("Fixing wrong rows...")
    fixWrongRows()
    print("Cleaning translates...")
    cleanTranslate()
elif result == 2:
    print("Updating the products in the worksheet products.xlsx...")
    updateProductsFromXLSX()
    print("Fixing wrong rows...")
    fixWrongRows()
    print("Cleaning translates...")
    cleanTranslate()

print("\nCompleted in %s seconds ---" % (time.time() - start_time))
time.sleep(3600)
