import os
import pandas as pd
import shutil
import sys
import yaml
sys.path[0] = '/home/eanderson/RPA/src'

from datetime import date, timedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep

from infrastructure.drive_upload import upload_file
from infrastructure.email import send_gmail


def browser(from_date, to_date):
    print('Setting up driver...', end=' ')
    # run in headless mode, enable downloads
    options = webdriver.IeOptions()
    options.add_argument('--window-size=1920x1080')
    options.add_argument('--disable-notifications')
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-setuid-sandbox")
    options.add_argument('--verbose')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-software-rasterizer')
    # options.add_argument('--headless')
    driver = webdriver.Ie(executable_path='C:\\Users\\mingus\\AppData\\Local\\IEDriverServer.exe', options=options)
    # params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': 'src/csv'}}
    # driver.execute('send_command', params)
    print('Done.')

    driver.get('https://myevolvacmhc.netsmartcloud.com/')

    # login
    with open('../config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['appleseed']
        pwd = login['pwd']
    driver.find_element_by_xpath(
        '/html/body/form/div[3]/center/table/tbody/tr[4]/td/div/div/div[2]/fieldset/table/tbody/tr[1]/td[2]/input')\
        .send_keys(usr)
    driver.find_element_by_xpath(
        '/html/body/form/div[3]/center/table/tbody/tr[4]/td/div/div/div[2]/fieldset/table/tbody/tr[2]/td[2]/input')\
        .send_keys(pwd)
    driver.find_element_by_xpath(
        '/html/body/form/div[3]/center/table/tbody/tr[4]/td/div/div/div[4]/input[1]').click()

    # navigate to claim details
    driver.find_element_by_xpath("//DIV[@id='tbTaskBar']/DIV/DIV/DIV/UL/LI[32]/A").click()
    finance_frame = driver.find_element_by_id('app_FINANCE')
    driver.switch_to.frame(finance_frame)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("//DIV[@id='contents_rpbNavigation_i10_i0_rtvSubModules']/ul[1]/li[1]/div/span").click()
    sleep(1)
    driver.find_element_by_xpath("//DIV[@id='contents_rpbNavigation_i10_i0_rtvSubModules']/ul[1]/li[1]/ul[1]/li[1]/div/span[2]").click()
    driver.implicitly_wait(10)

    # fill in form info
    internal_frame = driver.find_element_by_class_name('internal_iframe')
    driver.switch_to.frame(internal_frame)
    driver.implicitly_wait(5)
    form_frame = driver.find_element_by_id('iframe_frameset_0')
    driver.switch_to.frame(form_frame)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("//INPUT[@id='from_date']").send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath("//INPUT[@id='thru_date']").send_keys(to_date.strftime('%m/%d/%Y'))


def main():
    print('------------------------------' + date.today().strftime('%Y.%m.%d') + '------------------------------')
    print('Beginning Appleseed Paid Claims RPA...')
    from_date = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
    to_date = date.today().replace(day=1) - timedelta(days=1)

    browser(from_date, to_date)
    # merged_filename = join_datatables()
    # upload_file(merged_filename, '1lbGzRqPGekImmPBr3EXdtsayBQtSMmSl')
    # email_body = "Your monthly ISP due dates report (%s) is ready and available on the Appleseed RPA " \
    #              "Reports shared drive: https://drive.google.com/drive/folders/1lbGzRqPGekImmPBr3EXdtsayBQtSMmSl" \
    #              % merged_filename.split('/')[-1]
    # send_gmail('alester@appleseedcmhc.org', 'KHIT Report Notification', email_body)

    # os.remove('src/csv/treatment_due_dates.csv')
    # os.remove('src/csv/primary_workers.csv')
    # os.remove(merged_filename)

    print('Successfully finished Appleseed Paid Claims RPA!')


if __name__ == '__main__':
    try:
        main()
        # send_gmail('eanderson@khitconsulting.com',
        #            'KHIT Report Notification',
        #            'Successfully finished Appleseed ISP Due Dates RPA!')
    except Exception as e:
        print('System encountered an error running Appleseed Paid Claims RPA: %s' % e)
        email_body = 'System encountered an error running Appleseed MHA Due Dates RPA: %s' % e
        # send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)