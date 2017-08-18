import threading
import xlwt
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

WEBDRIVER_PATH = "E:\GitHub\\fit-crawler\dependencies\chromedriver.exe"

class Product:
    def __init__(self):
        self.name = ""
        self.variation = ""
        self.reference = ""
        self.price = 0
        self.url = ""

class Crawler:
    def __init__(self):
        self.driver = webdriver.Chrome(WEBDRIVER_PATH)
        self.categories = self.parse_categories()
        self.base_product_urls = []
        self.full_product_urls = []
        self.products = []

        self.manage_threading()
        self.extract_references()
        self.parse_specs()
        self.export_to_xls()

    def parse_categories(self):
        categories = ['Monitor de Atividade', 'Corrida', 'Ciclismo', 'Multidesporto', 'Natação', 'Caminhadas']
        self.driver.get("https://www.garmin.com/pt-PT/")

        for i, ctg in enumerate(categories):
            parse = self.driver.find_element_by_xpath("//a[contains(text(), '" + ctg + "')]")
            yield parse.get_attribute('href')

    def parse_products(self, url):
        thread_driver = webdriver.Chrome(WEBDRIVER_PATH)
        thread_driver.get(url)

        for i, val in enumerate(thread_driver.find_elements_by_xpath("//*[@data-product-id]")):
            self.base_product_urls.append(val.get_attribute('href'))
        thread_driver.close()

    def parse_specs(self):

        for i, val in enumerate(self.full_product_urls):
            self.driver.get(val)
            product = Product()

            try:
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="js__product__price__main"]')))
            except TimeoutException:
                print("DIDNT FIND PRICE SHIT")

            product.name = self.driver.find_element_by_xpath('//*[@id="js__product__meta"]/h1').text
            # TODO: Redundancy.
            product.reference = self.driver.find_element_by_xpath('//*[@id="js__product__meta"]/h3/span[@class="app__product__info__part-number--light"]').text
            product.price = self.driver.find_element_by_xpath('//*[@id="js__product__price__main"]/span[1]').text

            try:
                product.variation = self.driver.find_element_by_xpath('//*[@id="js__product__meta"]/h2').text
            except NoSuchElementException:
                pass

            print(product.name + ": " + product.variation + " | " + product.price + " | " + product.reference)

            self.products.append(product)

    def extract_references(self):
        for val in self.base_product_urls:
            self.driver.get(val)

            button_urls = self.driver.find_elements_by_xpath("//*[@data-sku]") # Button check.
            drop_urls = self.driver.find_elements_by_xpath("//*[@class='app__product__filters__select__list']/option") # Dropdown check.

            # TODO: Clean this function.
            for val in button_urls:
                url = val.get_attribute("href")
                if url not in self.full_product_urls:
                    self.full_product_urls.append(url)

            for val in drop_urls:
                url = val.get_attribute("value")
                if url not in self.full_product_urls:
                    self.full_product_urls.append("https://buy.garmin.com" + url)

    def manage_threading(self):
        t_ids = []

        for val in self.categories:
            t = threading.Thread(target=self.parse_products, args=(val,))
            t_ids.append(t)
            t.start()

        for i, val in enumerate(t_ids):
            val.join()

    def export_to_xls(self):
        book = xlwt.Workbook(encoding="utf-8")
        sheet = book.add_sheet("Garmin", cell_overwrite_ok=True)
        h_style = xlwt.easyxf("font: bold on")

        parameters = ["Name", "Variation", "Price", "Reference"]
        for i, var in enumerate(parameters):
            sheet.write(1, i+1, parameters[i], h_style)

        for i, var in enumerate(self.products):
            sheet.write(i+2, 1, self.products[i].name)
            sheet.write(i+2, 2, self.products[i].variation)
            sheet.write(i+2, 3, self.products[i].price)
            sheet.write(i+2, 4, self.products[i].reference)

        book.save("output.xls")

crawler = Crawler()
