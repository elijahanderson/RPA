import os
import pandas as pd
import shutil
import sys
import yaml

sys.path[0] = '/home/eanderson/RPA/src'

from datetime import date, datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep

from infrastructure.drive_upload import upload_folder
from infrastructure.email import send_gmail


def fancify(curr_date, qtr):
    print('Initiating fancification...', end=' ')
    path = f"src/csv/{curr_date.strftime('%Y')}q{qtr}"
    os.mkdir(path)
    
    q_services = pd.read_csv('src/csv/quarterly_services.csv')
    q_services['service_date'] = pd.to_datetime(q_services.service_date)
    # for each month, separate medicaid from non-medicaid, save both as different csvs in path
    for month, frame in q_services.groupby(pd.Grouper(key='service_date', freq='M')):
        frame['duration_client'] = frame['duration_client'].apply(
                lambda dur: '{:02d}:{:02d}'.format(*divmod(int(dur), 60)))
        frame['duration_worker'] = frame['duration_worker'].apply(
                lambda dur: '{:02d}:{:02d}'.format(*divmod(int(dur), 60)))
        non_medicaid = frame[frame['medicaid_number'].isna()]
        non_medicaid.drop_duplicates(subset=non_medicaid.columns.difference(
            ['duration_client', 'duration_worker', 'event_name']), inplace=True)
        non_medicaid.sort_values(by=['full_name', 'service_date'], ignore_index=True, inplace=True)
        frame.dropna(subset=['medicaid_number'], inplace=True)
        frame.drop_duplicates(subset=frame.columns.difference(
            ['duration_client', 'duration_worker', 'event_name']), inplace=True)
        frame.sort_values(by=['full_name', 'service_date'], ignore_index=True, inplace=True)
        frame.to_csv(f"src/csv/{month.strftime('%b')}-{curr_date.strftime('%Y')}_medicaid.csv", index=False)
        non_medicaid.to_csv(f"src/csv/{month.strftime('%b')}-{curr_date.strftime('%Y')}_non_medicaid.csv", index=False)
        shutil.move(f"src/csv/{month.strftime('%b')}-{curr_date.strftime('%Y')}_medicaid.csv", path)
        shutil.move(f"src/csv/{month.strftime('%b')}-{curr_date.strftime('%Y')}_non_medicaid.csv", path)

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
    #options.add_argument("--remote-debugging-port=9222")
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

    driver.get('https://myevolvaacogxb.netsmartcloud.com/')

    # login
    with open('src/config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['aacog']
        pwd = login['pwd']
    driver.find_element_by_id('MainContent_MainContent_userName').send_keys(usr)
    driver.find_element_by_id('MainContent_MainContent_chkDomain').click()
    driver.find_element_by_id('MainContent_MainContent_btnNext').click()
    driver.find_element_by_id('MainContent_MainContent_password').send_keys(pwd)
    driver.find_element_by_id('MainContent_MainContent_btnLogin').click()

    # navigate to client service entries
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('//*[@id="product-header-button-bar-id"]/li[18]/span').click()
    driver.find_element_by_xpath('//*[@id="product-header-mega-menu-level1-id"]/li[3]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[2]/ul[1]/li[2]').click()

    iframe1 = driver.find_element_by_xpath('//*[@id="MainContent_ctl35"]/iframe')
    driver.switch_to.frame(iframe1)
    iframe2 = driver.find_element_by_xpath('//*[@id="formset-target-wrapper-id"]/div[2]/iframe')
    driver.switch_to.frame(iframe2)
    driver.find_element_by_xpath('//*[@id="grdMain_FolderLink_0"]').click()
    driver.implicitly_wait(10)
    sleep(3)
    driver.find_element_by_xpath('//*[@id="grdMain_ObjectName_0"]').click()

    # new tab
    driver.implicitly_wait(10)
    driver.switch_to.default_content()
    driver.implicitly_wait(10)
    driver.switch_to.default_content()
    driver.implicitly_wait(10)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(10)
    driver.find_element_by_xpath('//*[@id="RP1_1A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_1B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    driver.implicitly_wait(10)
    driver.find_element_by_xpath('//*[@id="CSV"]').click()

    # rename the downloaded file
    sleep(90)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/quarterly_services.csv')
    
    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print('------------------------------ ' + date.today().strftime('%Y.%m.%d %H:%M') + ' ------------------------------')
    print('Beginning AACOG Client Services RPA Quarterly...')
    year = int(input('Enter year: '))
    qtr = int(input('Enter quarter: '))
    from_date = datetime(year, 3 * qtr - 2, 1)
    to_date = datetime(year+1, 1, 1)
    if not qtr == 4:
        to_date = datetime(year, 3 * qtr + 1, 1)
    print(f"Q{qtr}, {from_date.strftime('%Y.%m.%d')} - {to_date.strftime('%Y.%m.%d')}")
    browser(from_date, to_date)
    folder_path = fancify(from_date, qtr)
    upload_folder(folder_path, '1n1GPvhbcI2hkvx3DviVG82wrdrWPQhjE')
    email_body = f"Hey Amber, the AACOG Quarterly Client Services reports for {year}q{qtr}have been generated " \
                 "and are available in the AACOG shared drive: " \
                 "https://drive.google.com/drive/folders/1n1GPvhbcI2hkvx3DviVG82wrdrWPQhjE"
    send_gmail('adelaney@khitconsulting.com', 'KHIT Report Notification', email_body)

    os.remove('src/csv/quarterly_services.csv')
    shutil.rmtree(folder_path)

    print('Successfully finished AACOG Client Services Quarterly RPA!')


if __name__ == '__main__':
    try:
        main()
        send_gmail('eanderson@khitconsulting.com',
                   'KHIT Report Notification',
                   'Successfully finished AACOG Client Services Quarterly RPA!')
    except Exception as e:
        print('System encountered an error running AACOG Client Services Quarterly RPA: %s' % e)
        email_body = 'System encountered an error running AACOG Client Services Quarterly RPA: %s' % e
        send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)
