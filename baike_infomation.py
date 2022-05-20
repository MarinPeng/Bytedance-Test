from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import logging
import time
import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s''- %(levelname)s: %(message)s')


class BakeInfo():
    def __init__(self):
        self.browser = webdriver.Chrome()
        self.search_content = '字节跳动'
        self.url = 'http://www.baidu.com'

    def get_page(self):
        # 打开百度页面
        self.browser.get(self.url)
        # 设置页面大小
        self.browser.set_window_size(1920, 1080)
        # wait = WebDriverWait(self.browser, 10, 0.5)
        # 输入搜索内容
        self.browser.find_element(by=By.ID, value='kw').send_keys(self.search_content)
        self.browser.find_element(by=By.ID, value='su').click()
        # 进入百度百科
        time.sleep(1)
        self.element_timeout(ele_method=By.XPATH, method_value='//a[@aria-label="字节跳动，字节跳动，百度百科"]',
                             error_message='超时--获取元素失败')
        self.browser.find_element(by=By.XPATH, value='//a[@aria-label="字节跳动，字节跳动，百度百科"]').click()
        # 切换选项卡
        baike_window = self.browser.window_handles[-1]
        self.browser.switch_to.window(baike_window)
        try:
            self.element_timeout(ele_method=By.XPATH, method_value='//span[@class="long-title"]/h1',
                                 error_message='超时--获取元素失败')
            page_title = self.browser.find_element(by=By.XPATH, value='//span[@class="long-title"]/h1').text
            print('当前页面的title是{}'.format(page_title))
            # 拖动页面元素
            self.browser.execute_script('document.documentElement.scrollTop=3000')
            time.sleep(1)
            self.info_parse()
        except Exception as e:
            logging.info('获取元素失败，进入页面不符合预期:{}'.format(e))
            return False
        finally:
            self.browser.quit()



    # 数据解析入库
    def info_parse(self):
        div_list = self.browser.find_elements(by=By.XPATH,
                                              value='//div[@class="main-content J-content"]/div[@class="para"]')
        total_year = self.year_calculate()
        result_list = [['时间','事件']]
        index = 2012
        for div in div_list:
            info = div.text.split('，')
            if '北京抖音信息服务有限公司' in info[1]:
                data = self.data_clean(info)
                result_list.append(data)
                break
            for i in range(2012, 2012 + total_year + 1):
                if str(i) in info[0] and i >= index:
                    index = i
                    data = self.data_clean(info)
                    result_list.append(data)
        self.create_excel(result_list)

    # 清洗数据
    def data_clean(self,data_list):
        result = []
        data_list_1 = ','.join(data_list[1:]).split('[')
        data_list_1 = ','.join(data_list_1[:-1])
        result.append(data_list[0])
        result.append(data_list_1)
        return result

    # 写入Excel
    def create_excel(self, data):
        wb = Workbook()
        sheet =wb.active
        sheet.title='字节数据'
        for i in data:
            sheet.append(i)
        wb.save(filename='data_result.xlsx')

    # 元素等待
    def element_timeout(self, ele_method, method_value, error_message):
        wait = WebDriverWait(self.browser, 10, 0.5)
        wait.until(EC.presence_of_all_elements_located((ele_method, method_value)),
                   message='{}'.format(error_message))

    def year_calculate(self, ):
        start_year = datetime.datetime.strptime('2012', '%Y')
        now_year = datetime.datetime.now()
        sum_year = now_year - start_year
        return sum_year.days // 365


if __name__ == '__main__':
    test = BakeInfo()
    test.get_page()
    # test.year_calculate()
