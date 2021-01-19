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

from infrastructure.drive_upload import upload_folder
from infrastructure.email import send_gmail


def fancify(curr_date):
    print('Initiating fancification...', end=' ')
    report_bruns = pd.read_csv('src/csv/report_bruns.csv')
    report_moffett = pd.read_csv('src/csv/report_moffett.csv')
    report_bruns = report_bruns[['full_name', 'id_no', 'program_name', 'event_name', 'staff_name', 'duration',
                                 'actual_date', 'date_entered']]
    report_moffett = report_moffett[['full_name', 'id_no', 'program_name', 'event_name', 'staff_name', 'duration',
                                     'actual_date', 'date_entered']]
    report_bruns['actual_date'] = pd.to_datetime(report_bruns.actual_date)
    report_bruns['date_entered'] = pd.to_datetime(report_bruns.date_entered)
    report_moffett['actual_date'] = pd.to_datetime(report_moffett.actual_date)
    report_moffett['date_entered'] = pd.to_datetime(report_moffett.date_entered)

    report_bruns['date_diff'] = report_bruns['date_entered'] - report_bruns['actual_date']
    report_moffett['date_diff'] = report_moffett['date_entered'] - report_moffett['actual_date']

    report_bruns.sort_values(by=['staff_name', 'actual_date'], inplace=True)
    report_moffett.sort_values(by=['staff_name', 'actual_date'], inplace=True)

    xl_writer_bruns = pd.ExcelWriter('src/csv/report_bruns.xlsx', engine='xlsxwriter')
    xl_writer_moffett = pd.ExcelWriter('src/csv/report_moffett.xlsx', engine='xlsxwriter')

    report_bruns.to_excel(xl_writer_bruns, sheet_name='Bruns', index=False)
    report_moffett.to_excel(xl_writer_moffett, sheet_name='Moffett', index=False)
    workbook_bruns = xl_writer_bruns.book
    workbook_moffett = xl_writer_moffett.book
    worksheet_bruns = xl_writer_bruns.sheets['Bruns']
    worksheet_moffett = xl_writer_moffett.sheets['Moffett']

    red_format_bruns = workbook_bruns.add_format({'bg_color': '#FF4C4C',
                                                  'font_color': '#9C0006'})
    yellow_format_bruns = workbook_bruns.add_format({'bg_color':   '#FFEB9C',
                                                     'font_color': '#9C6500'})
    green_format_bruns = workbook_bruns.add_format({'bg_color':   '#C6EFCE',
                                                    'font_color': '#006100'})
    worksheet_bruns.conditional_format('I2:I'+str(len(report_bruns)+1), {'type': 'cell',
                                                                  'criteria': '>',
                                                                  'value': 3.5,
                                                                  'format': red_format_bruns})
    worksheet_bruns.conditional_format('I2:I' + str(len(report_bruns)+1), {'type': 'cell',
                                                                         'criteria': 'between',
                                                                         'minimum': 1.5,
                                                                         'maximum': 3.5,
                                                                         'format': yellow_format_bruns})
    worksheet_bruns.conditional_format('I2:I' + str(len(report_bruns)+1), {'type': 'cell',
                                                                         'criteria': '<',
                                                                         'value': 1.5,
                                                                         'format': green_format_bruns})

    red_format_moffett = workbook_moffett.add_format({'bg_color': '#FF4C4C',
                                                  'font_color': '#9C0006'})
    yellow_format_moffett = workbook_moffett.add_format({'bg_color': '#FFEB9C',
                                                     'font_color': '#9C6500'})
    green_format_moffett = workbook_moffett.add_format({'bg_color': '#C6EFCE',
                                                    'font_color': '#006100'})
    worksheet_moffett.conditional_format('I2:I' + str(len(report_moffett) + 1), {'type': 'cell',
                                                                             'criteria': '>',
                                                                             'value': 3.5,
                                                                             'format': red_format_moffett})
    worksheet_moffett.conditional_format('I2:I' + str(len(report_moffett) + 1), {'type': 'cell',
                                                                             'criteria': 'between',
                                                                             'minimum': 1.5,
                                                                             'maximum': 3.5,
                                                                             'format': yellow_format_moffett})
    worksheet_moffett.conditional_format('I2:I' + str(len(report_moffett) + 1), {'type': 'cell',
                                                                             'criteria': '<',
                                                                             'value': 1.5,
                                                                             'format': green_format_moffett})

    xl_writer_bruns.save()
    xl_writer_moffett.save()

    # create folder to upload and move xl sheets to it
    path = 'src/csv/' + curr_date.strftime('%m-%Y') + ' ABHS Service Entry Reports'
    os.mkdir(path)
    shutil.move('src/csv/report_bruns.xlsx', path)
    shutil.move('src/csv/report_moffett.xlsx', path)

    print('Fancified.')
    return path


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
        'download.default_directory': 'src/csv',
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing_for_trusted_sources_enabled': False,
        'safebrowsing.enabled': False
    })
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-software-rasterizer')
    options.add_argument('--headless')
    driver = webdriver.Chrome(executable_path='/usr/bin/chromedriver',
                              chrome_options=options)
    driver.command_executor._commands['send_command'] = ('POST', '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': 'src/csv'}}
    driver.execute('send_command', params)
    print('Done.')

    driver.get('https://myevolvwmabhsxb.netsmartcloud.com/')

    # login
    with open('src/config/login.yml', 'r') as yml:
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

    # switch to parameters iframe and add parameters to download Amanda's client info
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

    # rename the downloaded file
    sleep(5)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/report_bruns.csv')

    # download the CSV for Kelly's client info
    driver.switch_to.frame(iframe_params)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/table/tbody/tr[1]/td[4]/div/input') \
        .send_keys('Moffett, Kelly' + Keys.TAB)
    sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.frame(iframe1)
    driver.switch_to.frame(iframe2)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/ul/li[17]/a').click()
    sleep(5)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/report_moffett.csv')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print('------------------------------ ' + date.today().strftime('%Y.%m.%d' '%H:%M') + ' ------------------------------')
    print('Beginning ABHS Client Services RPA...')
    from_date = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
    to_date = date.today().replace(day=1)

    browser(from_date, to_date)
    folder_path = fancify(from_date)
    upload_folder(folder_path, '1h_Mym7ocK5lJ_-a4eZzShQf4DGm6HA8C')
    email_body = "Your monthly service entry reports (%s) are ready and available on the ABHS RPA " \
                 "Reports shared drive: https://drive.google.com/drive/folders/1h_Mym7ocK5lJ_-a4eZzShQf4DGm6HA8C" \
                 % folder_path.split('/')[-1]
    send_gmail('amanda.bruns@wmabhs.org', 'KHIT Report Notification', email_body)
    send_gmail('kelly.moffett-place@wmabhs.org', 'KHIT Report Notification', email_body)

    os.remove('src/csv/report_bruns.csv')
    os.remove('src/csv/report_moffett.csv')
    shutil.rmtree(folder_path)

    print('Successfully finished ABHS Client Services RPA!')


try:
    main()
    send_gmail('eanderson@khitconsulting.com',
               'KHIT Report Notification',
               'Successfully finished ABHS Client Services RPA!')
except Exception as e:
    print('System encountered an error running ABHS Service Entry RPA: %s' % e)
    email_body = 'System encountered an error running ABHS Service Entry RPA: %s' % e
    send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)
