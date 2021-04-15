import os
import numpy as np
import pandas as pd
import shutil
import sys
import yaml

sys.path[0] = '/home/eanderson/RPA/src'

from datetime import date, datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from traceback import print_exc

from infrastructure.drive_upload import upload_file
from infrastructure.email import send_gmail


def r2_5(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r2-5.csv')
    if not df.empty:
        today = date.today().strftime('%Y-%m-%d')
        df = df.rename(columns={'Date of Birth': 'dob'})
        df['dob'] = pd.to_datetime(df.dob)
        df['dob'] = df['dob'].apply(lambda dob: (pd.to_datetime(today) - dob) / np.timedelta64(1, 'Y'))
        adults = df[df['dob'] >= 18]
        minors = df[df['dob'] < 18]
        
        crisis_src.loc[1, curr_date] = len(adults)
        crisis_src.loc[2, curr_date] = len(minors)
        crisis_src.loc[3, curr_date] = len(df)

        # sum each row
        for idx, row in crisis_src.loc[0:3, :].iterrows():
            crisis_src.loc[idx, 'SFY 2021 Total'] = row.iloc[4:16].sum()

        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def r6_14(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r6-14.csv')
    if not df.empty:
        # drop all clients less than 18 y/o
        today = date.today().strftime('%Y-%m-%d')
        df["dob"] = pd.to_datetime(df.dob)
        df['dob'] = df['dob'].apply(lambda dob: (pd.to_datetime(today) - dob) / np.timedelta64(1, 'Y'))
        df = df[df['dob'] >= 18]

        crisis_src.loc[4, curr_date] = 0
        crisis_src.loc[5, curr_date] = 0
        crisis_src.loc[6, curr_date] = 0
        crisis_src.loc[7, curr_date] = 0
        crisis_src.loc[8, curr_date] = 0
        crisis_src.loc[11, curr_date] = 0

        for idx, row in df.iterrows():
            if 'Correctional' in row['answers_caption']:
                crisis_src.loc[4, curr_date] = crisis_src.loc[4, curr_date] + 1
            elif 'Nursing' in row['answers_caption']:
                crisis_src.loc[5, curr_date] = crisis_src.loc[5, curr_date] + 1
            elif 'ER' in row['answers_caption']:
                crisis_src.loc[6, curr_date] = crisis_src.loc[6, curr_date] + 1
            elif 'Inpatient' in row['answers_caption']:
                crisis_src.loc[7, curr_date] = crisis_src.loc[7, curr_date] + 1
            elif 'Med Unit' in row['answers_caption']:
                crisis_src.loc[8, curr_date] = crisis_src.loc[8, curr_date] + 1
            elif 'Community' in row['answers_caption']:
                crisis_src.loc[11, curr_date] = crisis_src.loc[11, curr_date] + 1

        # sum the column
        crisis_src.loc[12, curr_date] = crisis_src.loc[4:11, curr_date].sum()

        # sum each row
        for idx, row in crisis_src.loc[4:12, :].iterrows():
            crisis_src.loc[idx, 'SFY 2021 Total'] = row.iloc[4:16].sum()

        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def r17(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r17.csv')
    if not df.empty:
        # drop all clients less than 18 y/o
        today = date.today().strftime('%Y-%m-%d')
        df = df.rename(columns={'Date of Birth': 'dob'})
        df['dob'] = pd.to_datetime(df.dob)
        df['dob'] = df['dob'].apply(lambda dob: (pd.to_datetime(today) - dob) / np.timedelta64(1, 'Y'))
        df = df[df['dob'] >= 18]

        # set row 17 to the length of the df
        crisis_src.loc[15, curr_date] = len(df)
        # sum the row
        crisis_src.loc[15, 'SFY 2021 Total'] = crisis_src.iloc[15, 4:16].sum()

        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def r18_20(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r18-20.csv')
    if not df.empty:
        crisis_src.loc[16, curr_date] = 0
        crisis_src.loc[17, curr_date] = 0
        crisis_src.loc[18, curr_date] = 0
        for idx, row in df.iterrows():
            if row['Row'] == 'Row 18':
                crisis_src.loc[16, curr_date] = crisis_src.loc[16, curr_date] + 1
            elif row['Row'] == 'Row 19':
                crisis_src.loc[17, curr_date] = crisis_src.loc[16, curr_date] + 1
            elif row['Row'] == 'Row 20':
                crisis_src.loc[18, curr_date] = crisis_src.loc[18, curr_date] + 1

        # sum each program
        for idx, row in crisis_src.loc[16:18, :].iterrows():
            crisis_src.loc[idx, 'SFY 2021 Total'] = row.iloc[4:16].sum()

        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def r21(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r21.csv')
    if not df.empty:
        # set row 21 to the length of the df
        crisis_src.loc[19, curr_date] = len(df)
        
        df = pd.read_csv('src/csv/r6-14.csv')
        if not df.empty:
            # add client who had outreach to one of four specified locations
            today = date.today().strftime('%Y-%m-%d')
            df["dob"] = pd.to_datetime(df.dob)
            df['dob'] = df['dob'].apply(lambda dob: (pd.to_datetime(today) - dob) / np.timedelta64(1, 'Y'))
            df = df[df['dob'] >= 18]

            for idx, row in df.iterrows():
                if 'ER Cooper' in row['answers_caption']:
                    crisis_src.loc[19, curr_date] = crisis_src.loc[19, curr_date] + 1
                elif 'ER Jefferson Cherry Hill' in row['answers_caption']:
                    crisis_src.loc[19, curr_date] = crisis_src.loc[19, curr_date] + 1
                elif 'ER Stratford Jefferson' in row['answers_caption']:
                    crisis_src.loc[19, curr_date] = crisis_src.loc[19, curr_date] + 1
                elif 'Kennedy' in row['answers_caption']:
                    crisis_src.loc[19, curr_date] = crisis_src.loc[19, curr_date] + 1
                elif 'ER Virtua Berlin' in row['answers_caption']:
                    crisis_src.loc[19, curr_date] = crisis_src.loc[19, curr_date] + 1
                elif 'ER Virtua our Lady of Lourdes' in row['answers_caption']:
                    crisis_src.loc[19, curr_date] = crisis_src.loc[19, curr_date] + 1
                elif 'ER Virtua Voorhees' in row['answers_caption']:
                    crisis_src.loc[19, curr_date] = crisis_src.loc[19, curr_date] + 1

        # sum the row
        crisis_src.loc[19, 'SFY 2021 Total'] = crisis_src.iloc[19, 4:16].sum()
        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def r22_26(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r22-26.csv', error_bad_lines=False, engine='python')
    if not df.empty:
        # drop all clients less than 18 y/o
        today = date.today().strftime('%Y-%m-%d')
        df['dob'] = pd.to_datetime(df.dob)
        df['dob'] = df['dob'].apply(lambda dob: (pd.to_datetime(today) - dob) / np.timedelta64(1, 'Y'))
        df = df[df['dob'].astype(int) >= 18]
        df = df[df['program_name'] == 'Crisis Screening Services']

        crisis_src.loc[20, curr_date] = 0
        crisis_src.loc[21, curr_date] = 0
        crisis_src.loc[22, curr_date] = 0
        crisis_src.loc[23, curr_date] = 0
        crisis_src.loc[24, curr_date] = 0

        # sort the hospitals into their appropriate row
        for answer in df['answers_caption']:
            if 'STCF' in answer:
                crisis_src.loc[20, curr_date] = crisis_src.loc[20, curr_date] + 1
            if 'Other Involuntary Facility' in answer:
                crisis_src.loc[21, curr_date] = crisis_src.loc[21, curr_date] + 1
            if 'County Hospital' in answer:
                crisis_src.loc[22, curr_date] = crisis_src.loc[22, curr_date] + 1
            if 'State Hospital' in answer:
                crisis_src.loc[23, curr_date] = crisis_src.loc[23, curr_date] + 1
            if 'Voluntary Inpatient Facility' in answer:
                crisis_src.loc[24, curr_date] = crisis_src.loc[24, curr_date] + 1

        # sum each row
        for idx, row in crisis_src.loc[20:24, :].iterrows():
            crisis_src.loc[idx, 'SFY 2021 Total'] = row.iloc[4:16].sum()

        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def r27(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r27.csv')
    if not df.empty:
        # set row 27 to the length of the dataframe
        crisis_src.loc[25, curr_date] = len(df)
        # sum the row
        crisis_src.loc[25, 'SFY 2021 Total'] = crisis_src.iloc[25, 4:16].sum()

        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def r28_29(curr_date, crisis_src, xl_writer):
    df = pd.read_csv('src/csv/r28-29.csv')
    if not df.empty:
        df = df.rename(columns={'Time Between Entry and Adult Intake (Mins)': 'time_diff'})

        crisis_src.loc[26, curr_date] = 0
        crisis_src.loc[27, curr_date] = 0

        for idx, row in df.iterrows():
            if row['SecureStatus'] == 'Secured':
                crisis_src.loc[27, curr_date] = crisis_src.loc[27, curr_date] + 1
            elif row['SecureStatus'] == 'Unsecured':
                crisis_src.loc[26, curr_date] = crisis_src.loc[26, curr_date] + 1

        # sum each program
        for idx, row in crisis_src.loc[26:27, :].iterrows():
            crisis_src.loc[idx, 'SFY 2021 Total'] = row.iloc[4:16].sum()

        crisis_src.to_excel(xl_writer, sheet_name='crisis_src', index=False)
        xl_writer.save()


def to_excel_sheet():
    print('Beginning export to excel spreadsheet...', end=' ')

    curr_date = (date.today().replace(day=1) - timedelta(days=1)).strftime('%b_%Y').lower()
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r2_5(curr_date, crisis_src, xl_writer)
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r6_14(curr_date, crisis_src, xl_writer)
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r17(curr_date, crisis_src, xl_writer)
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r18_20(curr_date, crisis_src, xl_writer)
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r21(curr_date, crisis_src, xl_writer)
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r22_26(curr_date, crisis_src, xl_writer)
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r27(curr_date, crisis_src, xl_writer)
    crisis_src = pd.read_excel('src/csv/crisis_sfy_2021.xlsx', engine='openpyxl')
    xl_writer = pd.ExcelWriter('src/csv/crisis_sfy_2021.xlsx', engine='xlsxwriter')
    r28_29(curr_date, crisis_src, xl_writer)

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

    driver.get('https://myevolvoaksxb.netsmartcloud.com')

    print('Beginning browser automation...', end=' ')

    # login
    with open('src/config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['oaks']
        pwd = login['oaks-pwd']

    driver.find_element_by_id('MainContent_MainContent_chkDomain').click()
    driver.find_element_by_id('MainContent_MainContent_userName').send_keys(usr)
    driver.find_element_by_id('MainContent_MainContent_btnNext').click()
    driver.find_element_by_id('MainContent_MainContent_password').send_keys(pwd)
    driver.find_element_by_id('MainContent_MainContent_btnLogin').click()
    
    # navigate to RPA folder in custom reports
    driver.find_element_by_xpath('//*[@id="product-header-button-bar-id"]/li[19]/span').click()
    driver.find_element_by_xpath('//*[@id="product-header-mega-menu-level1-id"]/li[3]').click()
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(10)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/r17.csv')
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(15)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getmtime)
    print(filename)
    shutil.move(filename, 'src/csv/r21.csv')
    print('------------------ r21 ------------------------')
    [print(f) for f in os.listdir('src/csv')]
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(15)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getmtime)
    shutil.move(filename, 'src/csv/r27.csv')
    print('------------------ r27 ------------------------')
    [print(f) for f in os.listdir('src/csv')]
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(10)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getmtime)
    shutil.move(filename, 'src/csv/r18-20.csv')
    print('------------------ r18-20 ------------------------')
    [print(f) for f in os.listdir('src/csv')]
    sleep(1)
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(15)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getmtime)
    shutil.move(filename, 'src/csv/r28-29.csv')
    print('------------------ r28-29 ------------------------')
    [print(f) for f in os.listdir('src/csv')]
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(10)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getmtime)
    shutil.move(filename, 'src/csv/r2-5.csv')
    print('------------------ r2-5 ------------------------')
    [print(f) for f in os.listdir('src/csv')]
    # go back to RPA folder
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)

    # navigate to and download assessment results (rows 6-14 and 22-26)
    driver.find_element_by_xpath('//*[@id="product-header-button-bar-id"]/li[19]/span').click()
    driver.find_element_by_xpath('//*[@id="product-header-mega-menu-level1-id"]/li[4]').click()
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
        .send_keys((to_date-timedelta(days=1)).strftime('%m/%d/%Y'))
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(10)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/r6-14.csv')
    print('------------------ r6-14 ------------------------')
    [print(f) for f in os.listdir('src/csv')]
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
    while any([f.endswith('.crdownload') for f in os.listdir('src/csv')]):
        sleep(1)
    sleep(10)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/r22-26.csv')
    print('------------------ r22-26 ------------------------')
    [print(f) for f in os.listdir('src/csv')]
    print('Completed.')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print(
        '------------------------------ ' + datetime.now().strftime('%Y.%m.%d %H:%M') +
        ' ------------------------------ ')
    print('Beginning Oaks Crisis RPA...')
    from_date = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
    to_date = date.today().replace(day=1)

    if date.today().weekday() == 2:
        browser(from_date, to_date)
        to_excel_sheet()
        upload_file('src/csv/crisis_sfy_2021.xlsx', '14vjvXL3TIVD366xS08LIIxTBgnYsTns8')
        email_body = "Your monthly crisis report for (%s) is ready and available on the Oaks RPA " \
                     "Reports shared drive: https://drive.google.com/drive/u/0/folders/14vjvXL3TIVD366xS08LIIxTBgnYsTns8" \
                     % from_date.strftime('%b-%Y')
        send_gmail('Sherri.Dunn@oaksintcare.org', 'KHIT Report Notification', email_body)
        send_gmail('Krystle.Jarzyk@oaksintcare.org', 'KHIT Report Notification', email_body)

        for filename in os.listdir('src/csv'):
            if not filename.endswith('.xlsx'):
                os.remove('src/csv/%s' % filename)
        
        send_gmail('eanderson@khitconsulting.com',
                   'KHIT Report Notification',
                   'Successfully finished Oaks Crisis RPA!')
    else:
        print('Today is not Tuesday -- cancelling automation.')
    
    print('Successfully finished Oaks Crisis RPA!')


try:
    main()
except Exception as e:
    print('System encountered an error running Oaks Crisis RPA:\n')
    print_exc()
    email_body = 'System encountered an error running Oaks Crisis RPA: %s' % e
    send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)
