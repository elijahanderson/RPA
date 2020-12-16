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


def join_datatables():
    mha_csv = pd.read_csv('src/csv/mha_due_dates.csv')
    staff_csv = pd.read_csv('src/csv/direct_staff.csv')

    mha_csv = mha_csv.rename(columns={'Full Name': 'name'})
    mha_csv['name'] = mha_csv['name'].str.strip()
    staff_csv['name'] = staff_csv['name'].str.strip()
    staff_csv = staff_csv[['name', 'id_no', 'worker_name', 'worker_role']]

    merged = mha_csv.merge(staff_csv, on='name')
    merged['due_date'] = pd.to_datetime(merged.due_date)
    merged.sort_values(by=['name', 'due_date'], inplace=True, ascending=[True, True])

    filename = 'src/csv/' + str((date.today().replace(day=1) + timedelta(days=31)).month) + '-' + \
               str((date.today().replace(day=1) + timedelta(days=62)).month) + '-' + \
               str((date.today().replace(day=1) + timedelta(days=62)).year) + '_mha_due_dates.csv'

    merged.to_csv(filename, index=False)
    return filename


def browser():
    from_date = date.today().replace(day=28) + timedelta(days=4)
    to_date = from_date.replace(day=28) + timedelta(days=4)
    to_date = last_day_of_month(to_date)
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

    driver.get('https://myevolvacmhcxb.netsmartcloud.com/')

    # login
    with open('src/config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['appleseed']
        pwd = login['pwd']
    driver.find_element_by_id('MainContent_MainContent_userName').send_keys(usr)
    driver.find_element_by_id('MainContent_MainContent_password').send_keys(pwd)
    driver.find_element_by_id('MainContent_MainContent_btnLogin').click()

    # navigate to worker case loads (for clients' direct staff)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[18]/span').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[1]/ul/li[3]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[2]/ul[2]/li[25]').click()

    # navigate into the iframes and fill out the form
    iframe1 = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[2]/div/div[16]/div/div/div/div/iframe')
    driver.switch_to.frame(iframe1)
    driver.implicitly_wait(5)
    iframe2 = driver.find_element_by_xpath('/html/body/form/div[3]/div/div[2]/iframe[24]')
    driver.switch_to.frame(iframe2)
    driver.implicitly_wait(15)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[2]/div/input') \
        .send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[3]/div[2]/div/span').click()

    # switch back to default content for report selection
    sleep(3)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div[3]/div/div/div[1]/table/tbody/tr[3]/td[1]') \
        .click()

    # switch to parameters iframe
    driver.switch_to.frame(iframe1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(iframe2)
    driver.implicitly_wait(5)
    iframe_params = driver \
        .find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[4]/div[4]/div/div/iframe')
    driver.switch_to.frame(iframe_params)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/table/tbody/tr/td[2]/div/input') \
        .send_keys('Program')
    sleep(1)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/table/tbody/tr/td[2]/div/input') \
        .send_keys(Keys.TAB)
    sleep(3)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/table/tbody/tr[1]/td[4]/div/input') \
        .send_keys('Outpatient Mental Health' + Keys.TAB)
    sleep(2)

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
    shutil.move(filename, 'src/csv/direct_staff.csv')

    # switch to default content to download the MHA custom report
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()

    # navigate to custom reports
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[18]/span').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[1]/ul/li[2]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[2]/ul[1]/li[2]').click()

    # switch to appropriate iframe
    driver.implicitly_wait(15)
    cr_iframe1 = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[2]/div/div[16]/div/div/div/div/iframe')
    driver.switch_to.frame(cr_iframe1)
    driver.implicitly_wait(5)
    cr_iframe2 = driver.find_element_by_xpath('/html/body/form/div[3]/div/div[2]/iframe')
    driver.switch_to.frame(cr_iframe2)
    driver.find_element_by_xpath(
        '/html/body/form/table/tbody/tr/td/table/tbody/tr/td/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[2]/a').click()
    sleep(1)
    driver.find_element_by_xpath(
        '/html/body/form/table/tbody/tr/td/table/tbody/tr/td/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[2]/a').click()
    sleep(1)

    # new tab
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/mainbody/span/span/div/table/tbody/tr[1]/td/span/table/tbody/tr[1]/td['
        '2]/span[1]/input[1]'
    ).send_keys(from_date.replace(day=1).strftime('%m/%d/%Y'))
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/mainbody/span/span/div/table/tbody/tr[1]/td/span/table/tbody/tr[1]/td['
        '2]/span[2]/input[1]'
    ) .send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/mainbody/span/span/div/table/tbody/tr[2]/td/a[1]/input').click()
    driver.implicitly_wait(10)
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/span[1]/rdcondelement10/span/rdcondelement9/span/a/img').click()

    # rename the downloaded file
    sleep(5)  # wait for download
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/mha_due_dates.csv')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print('------------------------------' + date.today().strftime('%Y.%m.%d') + '------------------------------')
    print('Beginning Appleseed MHA Due Dates RPA...')
    browser()
    merged_filename = join_datatables()
    upload_file(merged_filename, '1zJLra5w3M9jxRbD3ac5GA4lzu1WRH9AQ')
    email_body = "Your monthly MHA due dates report (%s) is ready and available on the Appleseed RPA " \
                 "Reports shared drive: https://drive.google.com/drive/folders/1lbGzRqPGekImmPBr3EXdtsayBQtSMmSl" \
                 % merged_filename.split('/')[-1]
    send_gmail('alester@appleseedcmhc.org', 'KHIT Report Notification', email_body)

    os.remove('src/csv/mha_due_dates.csv')
    os.remove('src/csv/direct_staff.csv')
    os.remove(merged_filename)

    print('Successfully finished Appleseed MHA Due Dates RPA!')


if __name__ == '__main__':
    try:
        main()
        send_gmail('eanderson@khitconsulting.com',
                   'KHIT Report Notification',
                   'Successfully finished Appleseed MHA Due Dates RPA!')
    except Exception as e:
        print('System encountered an error running Appleseed MHA Due Dates RPA: %s' % e)
        email_body = 'System encountered an error running Appleseed MHA Due Dates RPA: %s' % e
        send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)

