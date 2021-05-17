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
from infrastructure.last_day_of_month import last_day_of_month


def join_dfs(df1, df2, filename):
    df1['name'] = df1['name'].str.strip()
    df2['name'] = df2['name'].str.strip()

    merged = df1.merge(df2, on='name')
    merged.drop(['people_id', 'id_no', 'ssn_number', 'ssi_number', 'urn_no', 'dob',
                 'phone_day', 'phone_evening', 'aka', 'intake_date', 'discharge_date',
                 'client_status', 'ipd', 'service_track_id', 'gender', 'medicaid_number',
                 'current_location', 'program_enrollment_event_id', 'program_info_id',
                 'worker_assignment_id', 'worker_start', 'worker_end', 'staff_id',
                 'supervisor_id', 'supervisor_name_y', 'is_primary_worker', 'managing_office_id',
                 'managing_office', 'worker_number', 'unit_number', 'worker_unit', 'prg_days',
                 'last_date_serv', 'total_dist', 'grand_total_dist', 'worker_role', 'program_name'], axis=1,
                inplace=True)
    merged = merged.rename(columns={'worker_name': 'primary_worker'})
    merged['expiration_date'] = pd.to_datetime(merged.expiration_date)
    merged.sort_values(by=['primary_worker', 'expiration_date'], inplace=True, ascending=[True, True])

    merged.to_csv(filename, index=False)
    return filename


def join_datatables(from_date, to_date):
    tp_df = pd.read_csv('src/csv/treatment_due_dates.csv')
    pri_df = pd.read_csv('src/csv/primary_workers.csv')
    all_df = pd.read_csv('src/csv/all_workers.csv')

    p_file = join_dfs(tp_df, pri_df, 'src/csv/primary_workers.csv')
    p_file = join_dfs(tp_df, all_df, 'src/csv/all_workers.csv')

    folder_path = 'src/csv/' + from_date.strftime('%b.%Y') + '-' + to_date.strftime('%b.%Y') + ' ISP Due Dates'
    os.mkdir(folder_path)
    shutil.move('src/csv/primary_workers.csv', folder_path)
    shutil.move('src/csv/all_workers.csv', folder_path)
    return folder_path


def browser(from_date, to_date):
    print('Setting up driver...', end=' ')
    # run in headless mode, enable downloads
    options = webdriver.ChromeOptions()
    options.add_argument('--window-size=1920x1080')
    options.add_argument('--disable-notifications')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-setuid-sandbox')
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

    # navigate to worker case loads (for clients' primary workers)
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

    driver.implicitly_wait(5)
    driver.switch_to.frame(iframe1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(iframe2)
    driver.implicitly_wait(5)
    # switch to parameters iframe
    iframe_params = driver \
        .find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[4]/div[4]/div/div/iframe')
    driver.switch_to.frame(iframe_params)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/table/tbody/tr/td[2]/div/input') \
        .send_keys('Program' + Keys.TAB)
    sleep(4)
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
    shutil.move(filename, 'src/csv/all_workers.csv')

    # primary workers only download
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[4]/div[2]/span/input').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/ul/li[17]/a').click()
    sleep(5)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/primary_workers.csv')

    # switch to default content to download the treatment plan custom report
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
        '/html/body/form/table/tbody/tr/td/table/tbody/tr/td/div[1]/div[1]/div[2]/table/tbody/tr[3]/td[2]/a').click()
    sleep(1)

    # new tab
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/mainbody/span/span/div/table/tbody/tr[1]/td/span/table/tbody/tr[1]/td[2]/'
        'span[1]/input[1]'
    ).send_keys(from_date.replace(day=1).strftime('%m/%d/%Y'))
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/mainbody/span/span/div/table/tbody/tr[1]/td/span/table/tbody/tr[1]/td[2]/'
        'span[2]/input[1]'
    ).send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/mainbody/span/span/div/table/tbody/tr[2]/td/a[1]/input').click()
    driver.implicitly_wait(10)
    driver.find_element_by_xpath(
        '/html/body/form/span[2]/span[2]/span[1]/rdcondelement10/span/rdcondelement9/span/a/img').click()

    # rename the downloaded file
    sleep(5)  # wait for download
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/treatment_due_dates.csv')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print('------------------------------ ' + date.today().strftime('%Y.%m.%d %H:%M') + ' ------------------------------')
    print('Beginning Appleseed ISP Due Dates RPA...')
    from_date = last_day_of_month(date.today()) + timedelta(days=1)
    to_date = last_day_of_month(from_date) + timedelta(days=1)
    to_date = last_day_of_month(to_date) + timedelta(days=1)

    browser(from_date, to_date)
    folder_path = join_datatables(from_date, to_date)
    upload_folder(folder_path, '1lbGzRqPGekImmPBr3EXdtsayBQtSMmSl')
    email_body = "Hi Amber, \nThe ISP Due Dates reports are ready and available in the shared folder as %s:" \
            "https://drive.google.com/drive/folders/1lbGzRqPGekImmPBr3EXdtsayBQtSMmSl" \
                 % folder_path.split('/')[-1]
    #send_gmail('alester@appleseedcmhc.org', 'KHIT Report Notification', email_body)
    
    os.remove('src/csv/treatment_due_dates.csv')
    shutil.rmtree(folder_path)

    print('Successfully finished Appleseed ISP Due Dates RPA!')


if __name__ == '__main__':
    try:
        main()
        send_gmail('eanderson@khitconsulting.com',
                   'KHIT Report Notification',
                   'Successfully finished Appleseed ISP Due Dates RPA!')
    except Exception as e:
        print('System encountered an error running Appleseed ISP Due Dates RPA: %s' % e)
        email_body = 'System encountered an error running Appleseed ISP Due Dates RPA: %s' % e
        send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)

