import os
import pandas as pd
import shutil
import sys
import yaml

sys.path[0] = '/home/eanderson/RPA/src'

from datetime import date, datetime, timedelta
from fpdf import FPDF
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
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
        'download.default_directory': 'D:\\Programming\\KHIT\RPA\\src\\csv',
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
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': 'D:\\Programming\\KHIT\RPA\\src\\csv'}}
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

    # navigate to RPA folder in custom reports
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[19]/span').click()
    driver.find_element_by_xpath('//*[@id="product-header-mega-menu-level1-id"]/li[2]').click()
    driver.find_element_by_xpath('//*[@id="eb7969e8-1e9c-4c86-8c11-f030bfe97b0f"]/li[2]').click()
    cr_frame1 = driver.find_element_by_xpath('//*[@id="MainContent_ctl36"]/iframe')
    driver.switch_to.frame(cr_frame1)
    cr_frame2 = driver.find_element_by_xpath('//*[@id="formset-target-wrapper-id"]/div[2]/iframe')
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_FolderLink_0"]').click()

    # row 17
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_ObjectName_0"]').click()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_1A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_1B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    # download and rename
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('//*[@id="CSV"]').click()
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getctime)
    shutil.move(filename, '../csv/r17.csv')
    # go back to RPA folder
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)

    # row 21
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_ObjectName_1"]').click()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_1A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_1B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    # download and rename
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('//*[@id="CSV"]').click()
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getmtime)
    shutil.move(filename, '../csv/r21.csv')
    # go back to RPA folder
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)

    # row 27
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_ObjectName_2"]').click()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_2A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_2B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    # download and rename
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('//*[@id="CSV"]').click()
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getmtime)
    shutil.move(filename, '../csv/r27.csv')
    # go back to RPA folder
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)

    # rows 18-20
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_ObjectName_3"]').click()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_1A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_1B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_2A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_2B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_3"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_4"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    # download and rename
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('//*[@id="CSV"]').click()
    while len(driver.window_handles) > 2:
        sleep(1)
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getmtime)
    shutil.move(filename, '../csv/r18-20.csv')
    # go back to RPA folder
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)

    # rows 28-29
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_ObjectName_4"]').click()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_1A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_1B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    # download and rename
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('//*[@id="CSV"]').click()
    while len(driver.window_handles) > 2:
        sleep(1)
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getmtime)
    shutil.move(filename, '../csv/r28-29.csv')
    # go back to RPA folder
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)

    # rows 2-5
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_ObjectName_10"]').click()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_1"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_2"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    # download and rename
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('//*[@id="CSV"]').click()
    while len(driver.window_handles) > 2:
        sleep(1)
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getmtime)
    shutil.move(filename, '../csv/r2-5.csv')
    # go back to RPA folder
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)

    # navigate to and download assessment results (rows 6-14 and 22-26)
    driver.find_element_by_xpath('//*[@id="product-header-button-bar-id"]/li[19]/span').click()
    driver.find_element_by_xpath('//*[@id="product-header-mega-menu-level1-id"]/li[3]').click()
    driver.find_element_by_xpath('//*[@id="d66dc6ec-03be-41b2-b808-206523c7e33d"]/li[15]').click()
    assessment_frame1 = driver.find_element_by_xpath('//*[@id="MainContent_ctl36"]/iframe')
    driver.switch_to.frame(assessment_frame1)
    driver.implicitly_wait(5)
    assessment_frame2 = driver.find_element_by_xpath('//*[@id="formset-target-wrapper-id"]/div[2]/iframe[14]')
    driver.switch_to.frame(assessment_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="D-749878CE-6477-4939-AA21-40C231E4F63F-input"]')\
        .send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="D-B28DDAE7-269C-42FA-9C04-53A6EABF445C-input"]')\
        .send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="rightContentFormId-ctrl-7"]/span').click()
    # switch back to default content for report selection
    sleep(1)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div[3]/div/div/div[1]/table/tbody/tr[1]/td[1]')\
        .click()
    # fill out parameters
    driver.switch_to.frame(assessment_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(assessment_frame2)
    driver.implicitly_wait(5)
    param_frame = driver.find_element_by_xpath('//*[@id="rightContentFormId-ctrl-16"]')
    driver.switch_to.frame(param_frame)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr/td[2]/div/input')\
        .send_keys('Assessment/Test' + Keys.TAB)
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[1]/td[4]/div/input')\
        .send_keys('OI - AIS Part 2')
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[1]/td[4]/div/input') \
        .send_keys(Keys.TAB)
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[2]/td[2]/div/input')\
        .send_keys('Question' + Keys.TAB)
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[2]/td[4]/div/input') \
        .send_keys('If Crisis Outreach occurred, please specify location of outreach:' + Keys.TAB)
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[3]/td[2]/div/input') \
        .send_keys('Program Assessing' + Keys.TAB)
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[3]/td[4]/div/input') \
        .send_keys('Crisis Screening Services' + Keys.TAB)
    sleep(1)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.frame(assessment_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(assessment_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="form-toolbar-16"]').click()
    # download and rename report
    while len(os.listdir('../csv')) < 7:
        sleep(1)
    sleep(5)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getctime)
    shutil.move(filename, '../csv/r6-14.csv')

    # rows 22-26
    driver.switch_to.frame(param_frame)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[1]/td[4]/div/input').clear()
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[1]/td[4]/div/input') \
        .send_keys('OI - USTF - Emergency / Screening' + Keys.TAB)
    sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[2]/td[4]/div/input').clear()
    driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr[2]/td[4]/div/input') \
        .send_keys('29. Hospital Discharge from in the Last 30 Days:' + Keys.TAB)
    sleep(1)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.frame(assessment_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(assessment_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="form-toolbar-16"]').click()
    # download and rename report
    dl_wait = True
    while len(os.listdir('../csv')) < 8 and dl_wait:
        for file in os.listdir('../csv'):
            if file.endswith('.crdownload'):
                dl_wait = True
        dl_wait = False
        sleep(1)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getctime)
    shutil.move(filename, '../csv/r22-26.csv')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print(
        '------------------------------ ' + datetime.now().strftime('%Y.%m.%d %H:%M') + ' ------------------------------'
    )
    print('Beginning Oaks Crisis RPA...')
    from_date = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
    to_date = date.today().replace(day=1)

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
