# -*- coding: utf-8 -*-
import csv
import glob
import os.path
from openpyxl import Workbook

from scrapy.spiders import Spider
from scrapy.http import Request


class AppsSpider(Spider):
    name = 'apps'
    allowed_domains = ['play.google.com']
    start_urls = ['http://play.google.com/store/apps/']

    def __init__(self, category=None):
        self.category = category

    def parse(self, response):
        # 1. Get category
        if self.category:
            category_url = response.xpath('//a[contains(text(), "' + self.category + '")]/@href').extract_first()
            absolute_category_url = response.urljoin(category_url)
            yield Request(absolute_category_url, callback=self.parse_category)
        else:
            categories = response.xpath('//a[@class="r2Osbf"]//@href').extract()
            for category in categories:
                absolute_category_url = response.urljoin(category)
                yield Request(absolute_category_url, callback=self.parse_category)

    def parse_category(self, response):
        # 2. Get all section of category
        category_name = response.xpath('//div[@jsname="j4gsHd"]//span/text()').extract_first()

        sections = response.xpath('//div[@jsname="O2DNWb"]/parent::div')

        # 3. Get all app of category
        for section in sections:
            absolute_section_link = response.urljoin(section.xpath('.//div/div[1]/a/@href').extract_first())
            section_name = section.xpath('.//div/div[1]/a/h2/text()').extract_first()

            yield Request(absolute_section_link,
                          meta={
                              'category_url': response.url,
                              'category_name': category_name,
                              'section_name': section_name,
                          },
                          callback=self.parse_apps)

        # 5. Find all sections
        # Find all section then parse_category

    def parse_apps(self, response):
        # 4. Get all information of app
        category_name = response.meta['category_name']
        category_url = response.meta['category_url']
        section_name = response.meta['section_name']
        section_url = response.url

        app_links = response.xpath('//div[@class="b8cIId ReQCgd Q9MA7b"]/a/@href').extract()
        for app_link in app_links:
            absolute_app_link = response.urljoin(app_link)
            yield Request(absolute_app_link,
                          meta={
                              'category_url': category_url,
                              'category_name': category_name,
                              'section_name': section_name,
                              'section_url': section_url,
                          },
                          callback=self.parse_app)

    def parse_app(self, response):
        category_url = response.meta['category_url']
        category_name = response.meta['category_name']
        section_name = response.meta['section_name']
        section_url = response.meta['section_url']

        image = response.xpath('//div[@class="xSyT2c"]/img/@src').extract_first()
        title = response.xpath('//h1//text()').extract_first()
        developer = response.xpath('//span[@class="T32cc UAO9ie"]//text()').extract_first()
        gender = response.xpath('//span[@class="T32cc UAO9ie"]//text()').extract_first()
        rating = response.xpath('//span[@class="AYi5wd TBRnV"]//text()').extract_first()
        price = response.xpath('//*[@itemprop="price"]/@content').extract_first()
        average_rating = response.xpath('//div[@class="BHMmbe"]//text()').extract_first()
        gallery = response.xpath('//div[@class="SgoUSc"]//img/@src').extract()
        description = response.xpath('//div[@jsname="sngebd"]//text()').extract_first()

        eligible = response.xpath(
            '//div[contains(text(),"Eligible for Family Library")]/following-sibling::span//text()').extract_first()
        updated = response.xpath('//div[contains(text(),"Updated")]/following-sibling::span//text()').extract_first()
        size = response.xpath('//div[contains(text(),"Size")]/following-sibling::span//text()').extract_first()
        installs = response.xpath('//div[contains(text(),"Installs")]/following-sibling::span//text()').extract_first()
        current_version = response.xpath(
            '//div[contains(text(),"Current Version")]/following-sibling::span//text()').extract_first()
        content_rating = response.xpath(
            '//div[contains(text(),"Content Rating")]/following-sibling::span//text()').extract_first()
        interactive_elements = response.xpath(
            '//div[contains(text(),"Interactive Elements")]/following-sibling::span//text()').extract_first()
        in_app_products = response.xpath(
            '//div[contains(text(),"In-app Products")]/following-sibling::span//text()').extract_first()
        permissions = response.xpath(
            '//div[contains(text(),"Permissions")]/following-sibling::span//text()').extract_first()
        report = response.xpath('//div[contains(text(),"Report")]/following-sibling::span//text()').extract_first()
        offered_by = response.xpath(
            '//div[contains(text(),"Offered By")]/following-sibling::span//text()').extract_first()

        yield {
            'category_name': category_name,
            'category_url': category_url,
            'section_name': section_name,
            'section_url': section_url,
            'image': image,
            'title': title,
            'developer': developer,
            'gender': gender,
            'rating': rating,
            'price': price,
            'average_rating': average_rating,
            'gallery': gallery,
            'description': description,
            'eligible': eligible,
            'updated': updated,
            'size': size,
            'installs': installs,
            'current_version': current_version,
            'content_rating': content_rating,
            'interactive_elements': interactive_elements,
            'in_app_products': in_app_products,
            'permissions': permissions,
            'report': report,
            'offered_by': offered_by,
        }

    def close(spider, reason):
        csv_file = max(glob.iglob('*.csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r', encoding="utf8") as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
