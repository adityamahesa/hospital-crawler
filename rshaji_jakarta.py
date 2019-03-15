import os
import re
import scrapy
from w3lib.html import remove_tags
from scrapy.utils.project import get_project_settings
from openpyxl import Workbook


class RSHajiJakartaItem(scrapy.Item):
    nama_dokter = scrapy.Field()
    poliklinik = scrapy.Field()
    jadwal_senin = scrapy.Field()
    jadwal_selasa = scrapy.Field()
    jadwal_rabu = scrapy.Field()
    jadwal_kamis = scrapy.Field()
    jadwal_jumat = scrapy.Field()
    jadwal_sabtu = scrapy.Field()


class RSHajiJakartaSpider(scrapy.Spider):
    name = 'rshaji-jakarta'
    start_urls = ['https://www.rshaji-jakarta.com/dokter']

    def parse(self, response):
        daftar_url_dokter = response.css('a.dokter-popup-detail ::attr(href)').extract()
        for url_dokter in daftar_url_dokter:
            yield scrapy.Request(
                'https://www.rshaji-jakarta.com/' + url_dokter,
                self.parse_doctor_data,
            )

    def parse_doctor_data(self, response):
        nama_dokter = response.css('div.dokter-content-title h1 ::text').extract()[0]
        poliklinik = response.css('h2.dokter-h2 ::text').extract()[0]
        daftar_jadwal = response.css('table.jadwal-dokter').extract()
        # Menghilangkan tag HTML
        daftar_jadwal = [remove_tags(jadwal) for jadwal in daftar_jadwal]
        # Menghapus spasi yang berlebih
        daftar_jadwal = [re.sub(r' +', ' ', jadwal) for jadwal in daftar_jadwal]
        # Mengganti &amp; menjadi simbol &
        daftar_jadwal = [re.sub(r'(&amp;)', '&', jadwal) for jadwal in daftar_jadwal]
        # Menghapus newline beserta spasi nya jika memungkinkan
        daftar_jadwal = [re.sub(r'\n( +)?', '', jadwal) for jadwal in daftar_jadwal]
        # Menghilangkan spasi setelah tanda (
        daftar_jadwal = [re.sub(r'\( ', '(', jadwal) for jadwal in daftar_jadwal]
        # Memberi spasi sebelum tanda (
        daftar_jadwal = [re.sub(r'(?<=[^\s])(?=[\(])', ' ', jadwal) for jadwal in daftar_jadwal]
        # Menghilangkan spasi sebelum tanda )
        daftar_jadwal = [re.sub(r' \)', ')', jadwal) for jadwal in daftar_jadwal]
        # Memberi tanda koma dan spasi setelah tanda )
        daftar_jadwal = [re.sub(r'(?<=[\)])(?=[^\s])', ', ', jadwal) for jadwal in daftar_jadwal]
        # Memastikan adanya spasi setelah koma
        daftar_jadwal = [re.sub(r'(?<=[\,])(?=[^\s])', ' ', jadwal) for jadwal in daftar_jadwal]
        
        item = RSHajiJakartaItem()
        item['nama_dokter'] = nama_dokter
        item['poliklinik'] = poliklinik
        item['jadwal_senin'] = daftar_jadwal[0]
        item['jadwal_selasa'] = daftar_jadwal[1]
        item['jadwal_rabu'] = daftar_jadwal[2]
        item['jadwal_kamis'] = daftar_jadwal[3]
        item['jadwal_jumat'] = daftar_jadwal[4]
        item['jadwal_sabtu'] = daftar_jadwal[5]
        
        yield item


class RSHajiJakartaPipeline(object):
    daftar_nama_kolom = [
        'Nama Dokter',
        'Nama Poliklinik',
        'Senin',
        'Selasa',
        'Rabu',
        'Kamis',
        'Jumat',
        'Sabtu',
    ]
    data = []

    def __init__(self):
        if not self.data:
            self.data.append(self.daftar_nama_kolom)

    def process_item(self, item, spider):
        if spider.name == 'rshaji-jakarta':
            self.data.append([
                item['nama_dokter'],
                item['poliklinik'],
                item['jadwal_senin'],
                item['jadwal_selasa'],
                item['jadwal_rabu'],
                item['jadwal_kamis'],
                item['jadwal_jumat'],
                item['jadwal_sabtu'],
            ])
    
    def close_spider(self, spider):
        work_book = Workbook()
        work_sheet = work_book.active
        work_sheet.title = 'Data'
        for record in self.data:
            work_sheet.append(record)
        
        dirname = 'saved'
        filename = 'rshaji_jakarta.xlsx'
        if not os.path.isdir(dirname):
            os.makedirs(dirname)
        work_book.save(os.path.join(dirname, filename))


settings = get_project_settings()
settings['ITEM_PIPELINES'] = {'rshaji_jakarta.RSHajiJakartaPipeline': 300}


if __name__ == '__main__':
    from scrapy.crawler import CrawlerProcess

    process = CrawlerProcess(settings)
    process.crawl(RSHajiJakartaSpider)
    process.start()
