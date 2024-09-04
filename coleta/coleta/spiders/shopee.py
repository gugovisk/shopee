import scrapy


class ShopeeSpider(scrapy.Spider):
    name = "shopee"
    allowed_domains = ["shopee.com.br"]
    start_urls = ["https://shopee.com.br/collections/5193896"]

    def parse(self, response):
        pass
