import datetime
import xlrd
import requests
import json, re, random
from openpyxl import Workbook
import pandas as pd

class AliExpress_Order():

    def __init__(self, product_id, file_path):
        self.product_id = product_id
        self.file_path = file_path
        self.ua_list = [
            'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1468.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.1; rv:21.0) Gecko/20130328 Firefox/21.0',
            'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.17 (KHTML, like Gecko) Chrome/24.0.1312.60 Safari/537.17',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2225.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.2; Win64; x64; rv:16.0.1) Gecko/20121011 Firefox/21.0.1',
            'Mozilla/5.0 (X11; CrOS i686 4319.74.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.57 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.0 Safari/537.36',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.2117.157 Safari/537.36',
            'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.116 Safari/537.36 Mozilla/5.0 (iPad; U; CPU OS 3_2 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Version/4.0.4 Mobile/7B334b Safari/531.21.10',
            'Mozilla/5.0 (Windows x86; rv:19.0) Gecko/20100101 Firefox/19.0',
            'Mozilla/5.0 (X11; CrOS i686 3912.101.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.116 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.15 (KHTML, like Gecko) Chrome/24.0.1295.0 Safari/537.15',
            'Mozilla/5.0 (Windows NT 5.1; rv:21.0) Gecko/20130331 Firefox/21.0',
            'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.16 Safari/537.36',
            'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; Media Center PC 6.0; InfoPath.2; MS-RTC LM 8',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1664.3 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.2; Win64; x64;) Gecko/20100101 Firefox/20.0',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2049.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.62 Safari/537.36',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.137 Safari/4E423F',
            'Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.14 (KHTML, like Gecko) Chrome/24.0.1292.0 Safari/537.14',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1664.3 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.0; WOW64; rv:24.0) Gecko/20100101 Firefox/24.0',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.3; rv:36.0) Gecko/20100101 Firefox/36.0',
            'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; yie8)',
            'Mozilla/5.0 (compatible; MSIE 10.0; Macintosh; Intel Mac OS X 10_7_3; Trident/6.0)',
            'Mozilla/5.0 (Windows NT 4.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2049.0 Safari/537.36',
            'Mozilla/5.0 (compatible; MSIE 7.0; Windows NT 5.0; Trident/4.0; FBSMTWB; .NET CLR 2.0.34861; .NET CLR 3.0.3746.3218; .NET CLR 3.5.33652; msn OptimizedIE8;ENUS)'
        ]
        self.api_base = "https://feedback.aliexpress.com/display/evaluationProductDetailAjaxService.htm?page=1&productId={}&type=default"
        self.api_page_base = "https://feedback.aliexpress.com/display/evaluationProductDetailAjaxService.htm?page={}&productId={}&type=default"
        self.wb = Workbook()
        self.sheet = self.wb.active
        self.sheet.title = self.product_id

    def get_page_url(self):
        url = self.api_base.format(product_id)
        headers = {
            'User-Agent': random.choice(self.ua_list)
        }
        response = requests.get(url, headers=headers).content.decode()
        # 获取json字符串
        json_str = re.match('.*?({.*}).*?', response, re.S).group(1)
        # 获取json数据
        # print(json_str)
        page_total = json.loads(json_str)['page']['total']
        for page in range(1, int(page_total) + 1):
            page_url = self.api_page_base.format(page, product_id)
            yield page, page_url

    def get_order_detail(self, page, page_url):
        item = {}
        url = page_url
        headers = {
            'User-Agent': random.choice(self.ua_list)
        }
        response = requests.get(url, headers=headers).content.decode()
        json_str = re.match('.*?({.*}).*?', response, re.S).group(1)
        # 获取json数据
        # print(json_str)
        json_dict = json.loads(json_str)['records']
        # print(json_dict)
        for order in json_dict:
            item['product_id'] = self.product_id
            item['name'] = order['name']
            item['country'] = order['countryCode']
            item['quantity'] = order['quantity']
            item['date'] = str(datetime.datetime.strptime(order['date'], '%d %b %Y %H:%M'))
            item['page'] = page
            yield item

    def save_data(self, data_list):
        self.sheet['A1'].value = 'product_id'
        self.sheet['B1'].value = 'name'
        self.sheet['C1'].value = 'country'
        self.sheet['D1'].value = 'quantity'
        self.sheet['E1'].value = 'date'
        self.sheet['F1'].value = 'page'
        j = 2
        for item in data_list:
            item = eval(item)
            self.sheet['A' + str(j)].value = item['product_id']
            self.sheet['B' + str(j)].value = item['name']
            self.sheet['C' + str(j)].value = item['country']
            self.sheet['D' + str(j)].value = item['quantity']
            self.sheet['E' + str(j)].value = item['date']
            self.sheet['F' + str(j)].value = item['page']
            j = j + 1
        self.wb.save(self.file_path)
        print('保存数据完毕！')

    def statistics(self):
        file_name = self.file_path
        data = pd.read_excel(file_name)  # 读数据，以序列号做为索引
        data1 = data.groupby(by=['country']).agg({'quantity': ['sum', 'count']})
        data['year-month'] = pd.DatetimeIndex(data.date).map(lambda x: 100 * x.year + x.month)
        data2 = data.groupby(by=['year-month']).agg({'quantity': ['sum', 'count']})
        data1.rename(columns={'quantity': '数量'}, inplace=True)
        data1.rename(columns={'sum': '产品数量'}, inplace=True)
        data1.rename(columns={'count': '订单数量'}, inplace=True)
        data2.rename(columns={'quantity': '数量'}, inplace=True)
        data2.rename(columns={'sum': '产品数量'}, inplace=True)
        data2.rename(columns={'count': '订单数量'}, inplace=True)
        print(data2)
        print(data1)
        try:
            with pd.ExcelWriter(file_name) as writer:
                data.to_excel(writer, sheet_name='原数据')
                data1.to_excel(writer, sheet_name='按国家统计订单数量')
                data2.to_excel(writer, sheet_name='按月统计订单数量')
        except:
            pass

    def run(self):
        data_list = list()
        for page, page_url in self.get_page_url():
            for item in self.get_order_detail(page, page_url):
                print(item)
                data_list.append(str(item))
        print('获取数据完毕！！')
        self.save_data(data_list)
        self.statistics()


if __name__=="__main__":
    product_id = '32918525062'
    file_path = 'bb.xlsx'
    aliexpress_order = AliExpress_Order(product_id, file_path)
    aliexpress_order.run()