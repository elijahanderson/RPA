import os
import pandas as pd
import shutil
import sys
import yaml

sys.path[0] = '/home/eanderson/RPA/src'

from datetime import date, datetime, timedelta
from fpdf import FPDF
from selenium import webdriver
from time import sleep
from traceback import print_exc

from infrastructure.drive_upload import upload_folder
from infrastructure.email import send_gmail


def to_excel_sheet():
    print('Beginning CSV modifications...', end=' ')
    print('Done.')


def browser(from_date, to_date):
    print('Setting up driver...', end=' ')

    # run in headless mode, enable downloads
    options = webdriver.ChromeOptions()
    options.add_argument('--window-size=1920x1080')
    options.add_argument('--disable-notifications')
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-setuid-sandbox")
    options.add_argument('--verbose')
    options.add_experimental_option('prefs', {
        'download.default_directory': '../csv',
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing_for_trusted_sources_enabled': False,
        'safebrowsing.enabled': False
    })
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-software-rasterizer')
    # options.add_argument('--headless')
    driver = webdriver.Chrome(executable_path='C:\\Users\\mingus\\AppData\\Local\\chromedriver.exe',
                              chrome_options=options)
    driver.command_executor._commands['send_command'] = ('POST', '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': '../csv'}}
    driver.execute('send_command', params)
    print('Done.')

    driver.get('https://myevolvoaksxb.netsmartcloud.com')

    # login
    with open('../config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['oaks']
        pwd = login['pwd']
    driver.find_element_by_id('MainContent_MainContent_chkDomain').click()
    driver.find_element_by_id('MainContent_MainContent_userName').send_keys(usr)
    driver.find_element_by_id('MainContent_MainContent_btnNext').click()
    driver.find_element_by_id('MainContent_MainContent_password').send_keys(pwd)
    driver.find_element_by_id('MainContent_MainContent_btnLogin').click()

    # download and rename the report
    driver.find_element_by_xpath('/html/body/form/span[5]/span/rdcondelement4/span/a/img').click()
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/recipient_codes.csv')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print(
        '------------------------------ ' + datetime.now().strftime('%Y.%m.%d %H:%M') + '------------------------------'
    )
    print('Beginning Oaks Crisis RPA...')
    from_date = date.today() - timedelta(days=1)
    to_date = date.today()

    browser(from_date, to_date)
    # to_excel_sheet()
    file_path = 'src/%s' % from_date.strftime('%m-%d-%Y')
    # upload_folder(file_path, '1lYsW4yfourbnFYJB3GLh6br7D1_3LOcd')
    # email_body = "Your monthly crisis report for (%s) is ready and available on the Oaks RPA " \
    #              "Reports shared drive: https://drive.google.com/drive/folders/1lYsW4yfourbnFYJB3GLh6br7D1_3LOcd" \
    #              % from_date.strftime('%m-%Y')
    # send_gmail('iweber@fremont.gov', 'KHIT Report Notification', email_body)

    # for filename in os.listdir('src/csv'):
    #     os.remove('src/csv/%s' % filename)
    # shutil.rmtree(folder_path)

    print('Successfully finished Oaks Crisis RPA!')


main()
# try:
#     main()
#     send_gmail('eanderson@khitconsulting.com',
#                'KHIT Report Notification',
#                'Successfully finished Fremont ISL RPA!')
# except Exception as e:
#     print('System encountered an error running Fremont ISL RPA:\n')
#     print_exc()
#     email_body = 'System encountered an error running Fremont ISL RPA: %s' % e
#     send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)
