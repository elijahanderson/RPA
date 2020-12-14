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
from infrastructure.last_day_of_month import last_day_of_month


def browser():
    from_date = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
    to_date = last_day_of_month(date.today())
    print('Setting up driver...', end=' ')

    # run in headless mode, enable downloads
    options = webdriver.ChromeOptions()
    options.add_argument('--window-size=1920x1080')
    options.add_argument('--disable-notifications')
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-setuid-sandbox")
    options.add_argument('--verbose')
    options.add_experimental_option('prefs', {
        'download.default_directory': 'src/csv',
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing_for_trusted_sources_enabled': False,
        'safebrowsing.enabled': False
    })
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-software-rasterizer')
    # options.add_argument('--headless') # LINUX change to /usr/bin/chromedriver
    driver = webdriver.Chrome(executable_path='C:\\Users\\mingus\\AppData\\Local\\chromedriver.exe',
                              chrome_options=options)
    driver.command_executor._commands['send_command'] = ('POST', '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': 'src/csv'}}
    driver.execute('send_command', params)
    print('Done.')

    driver.get('https://myevolvwmabhsxb.netsmartcloud.com/')

    # login -- LINUX change to src/config/login.yml
    with open('../config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['abhs']
        pwd = login['pwd']
    driver.find_element_by_id('MainContent_MainContent_userName').send_keys(usr)
    driver.find_element_by_id('MainContent_MainContent_btnNext').click()
    driver.find_element_by_id('MainContent_MainContent_password').send_keys(pwd)
    driver.find_element_by_id('MainContent_MainContent_btnLogin').click()

    # navigate to client service entries
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[18]/span').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[1]/ul/li[3]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[2]/ul[1]/li[3]').click()

    iframe1 = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[2]/div/div[16]/div/div/div/div/iframe')
    driver.switch_to.frame(iframe1)
    iframe2 = driver.find_element_by_xpath('/html/body/form/div[3]/div/div[2]/iframe[2]')
    driver.switch_to.frame(iframe2)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[2]/div/input')\
        .send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[3]/div/input') \
        .send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[3]/div[2]/div/span').click()

    # switch back to default content for report selection
    sleep(3)
    driver.implicitly_wait(3)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div[3]/div/div/div[1]/table/tbody/tr[3]/td[1]') \
        .click()

    # switch to parameters iframe and add parameters
    driver.switch_to.frame(iframe1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(iframe2)
    driver.implicitly_wait(5)
    iframe_params = driver \
        .find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[4]/div[11]/div/div/iframe')
    driver.switch_to.frame(iframe_params)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/table/tbody/tr/td[2]/div/input')\
        .send_keys('Supervisor' + Keys.TAB)
    sleep(1)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/table/tbody/tr[1]/td[4]/div/input')\
        .send_keys('Bruns, Amanda' + Keys.TAB)
    driver.implicitly_wait(5)

    # switch back to iframe1 for CSV button
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.frame(iframe1)
    driver.switch_to.frame(iframe2)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/ul/li[17]/a').click()

    sleep(20)


def main():
    print('Beginning ABHS Client Services RPA...')
    browser()
    merged_filename = None
    # upload_file(merged_filename, '1zJLra5w3M9jxRbD3ac5GA4lzu1WRH9AQ')
    # email_body = "Your monthly MHA due dates report (%s) is ready and available on the Appleseed RPA " \
    #              "Reports shared drive: https://drive.google.com/drive/folders/1lbGzRqPGekImmPBr3EXdtsayBQtSMmSl" \
    #              % merged_filename.split('/')[-1]
    # send_gmail('alester@appleseedcmhc.org', 'KHIT Report Notification', email_body)
    #
    # os.remove('src/csv/mha_due_dates.csv')
    # os.remove('src/csv/direct_staff.csv')
    # os.remove(merged_filename)

    print('Successfully finished ABHS Client Services RPA!')


main()
