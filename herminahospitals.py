import os
import json
import re

import scrapy
from scrapy.utils.project import get_project_settings
from openpyxl import Workbook


class HerminaHospitalsItem(scrapy.Item):
    dokter_id = scrapy.Field()
    dokter_nama = scrapy.Field()
    dokter_kategori = scrapy.Field()
    cabang_id = scrapy.Field()
    cabang_nama = scrapy.Field()
    jadwal_hari = scrapy.Field()
    jadwal_mulai = scrapy.Field()
    jadwal_selesai = scrapy.Field()


class HerminaHospitalsSpider(scrapy.Spider):
    name = 'herminahospitals'
    start_urls = ['http://services.herminahospitals.com/home/jadwal']
    branches = {}
    token = ''

    def parse(self, response):
        self.token = response.xpath('//input[@name="_token"]/@value').extract_first()
        branch_ids = response.xpath('//select[@name="branch_id"]/option/@value').extract()
        branch_ids = branch_ids[1:]
        branch_name = response.xpath('//select[@name="branch_id"]/option/text()').extract()
        branch_name = branch_name[1:]
        self.branches = {branch_ids[i]:branch_name[i] for i, e in enumerate(branch_ids)}
        for branch_id in self.branches.keys():
            payload = {
                '_token': self.token,
                'param': {
                    'branch_id': branch_id,
                },
            }
            payload = json.dumps(payload)
            yield scrapy.Request(
                'http://services.herminahospitals.com/home/spesialis/get-list',
                callback=self.parse_specialist,
                method='POST',
                headers={
                    'Content-Type': 'application/json',
                },
                body=payload,
                meta={
                    'branch_id': branch_id,
                },
            )
            
    
    def parse_specialist(self, response):
        specialists = json.loads(response.body)
        for specialist in specialists:
            payload = {
                '_token': self.token,
                'param': {
                    'branch_id': response.meta['branch_id'],
                    'specialist_id': specialist['spesialis_cd'],
                    'name': '',
                }
            }
            payload = json.dumps(payload)
            yield scrapy.Request(
                'http://services.herminahospitals.com/home/dokter/get-dokter-list',
                callback=self.parse_doctor_url,
                method='POST',
                headers={
                    'Content-Type': 'application/json',
                },
                body=payload,
                meta={
                    'branch_id': response.meta['branch_id'],
                },
            )

    def parse_doctor_url(self, response):
        doctor_urls = re.findall(r'\"http\:(.+?)\"', response.body.decode())
        doctor_urls = ['http:' + re.sub(r'\\', '', i) for i in doctor_urls]
        for doctor_url in doctor_urls:
            yield scrapy.Request(
                doctor_url,
                callback=self.parse_doctor_schedule,
                meta={
                    'branch_id': response.meta['branch_id'],
                },
            )
    
    def parse_doctor_schedule(self, response):
        branch_id = response.meta['branch_id']
        branch_name = self.branches[branch_id]
        doctor_code = response.xpath('//input[@id="dr_cd"]/@value').extract_first()
        doctor_name = response.xpath('//td[@class="nama-dokter"]/text()').extract_first()
        doctor_category = response.xpath('//td[@class="kategori-dokter"]/text()').extract_first()
        schedules_day = re.findall(r'\[ DayEnum\.(.+?) \]', response.body.decode())
        schedules_start = re.findall(r'start\: \'(.+?)\'\,', response.body.decode())
        schedules_end = re.findall(r'end\: \'(.+?)\'\,', response.body.decode())
        if doctor_code:
            for i, _ in enumerate(schedules_day):
                item = HerminaHospitalsItem()
                item['dokter_id'] = doctor_code
                item['dokter_nama'] = doctor_name
                item['dokter_kategori'] = doctor_category
                item['cabang_id'] = branch_id
                item['cabang_nama'] = branch_name
                item['jadwal_hari'] = schedules_day[i]
                item['jadwal_mulai'] = schedules_start[i]
                item['jadwal_selesai'] = schedules_end[i]

                yield item


class HerminaHospitalsPipeline(object):
    daftar_nama_kolom= [
        'ID Dokter',
        'Nama',
        'Kategori',
        'ID Cabang',
        'Cabang',
        'Hari',
        'Mulai',
        'Selesai',
    ]
    data = []

    def __init__(self):
        if not self.data:
            self.data.append(self.daftar_nama_kolom)

    def process_item(self, item, spider):
        if spider.name == 'herminahospitals':
            record = [
                item['dokter_id'],
                item['dokter_nama'],
                item['dokter_kategori'],
                item['cabang_id'],
                item['cabang_nama'],
                item['jadwal_hari'],
                item['jadwal_mulai'],
                item['jadwal_selesai'],
            ]
            if record not in self.data:
                self.data.append(record)

    def close_spider(self, spider):
        work_book = Workbook()
        work_sheet = work_book.active
        work_sheet.title = 'Data'
        for record in self.data:
            work_sheet.append(record)
        
        dirname = 'saved'
        filename = 'herminahospitals.xlsx'
        if not os.path.isdir(dirname):
            os.makedirs(dirname)
        work_book.save(os.path.join(dirname, filename))


settings = get_project_settings()
settings['ITEM_PIPELINES'] = {'herminahospitals.HerminaHospitalsPipeline': 300}


if __name__ == '__main__':
    from scrapy.crawler import CrawlerProcess

    process = CrawlerProcess(settings)
    process.crawl(HerminaHospitalsSpider)
    process.start()
