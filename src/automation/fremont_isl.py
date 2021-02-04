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


def create_isl(frame, staff, program_modifier, from_date):
    isl_pdf = FPDF(orientation='L')
    isl_pdf.add_page()
    isl_pdf.set_font(family='Arial', size=11)
    isl_pdf.cell(w=0, h=4, txt='ALAMEDA COUNTY BEHAVIORAL HEALTH CARE - MENTAL HEALTH', align='C', ln=2)
    isl_pdf.cell(w=0, h=5, txt='INDIVIDUAL STAFF LOG', align='C', ln=2)
    isl_pdf.cell(w=0, h=10, txt='REPORTING UNIT:    01EI1 - City of Fremont MMH', align='L', ln=0)
    isl_pdf.cell(w=0, h=10, txt='STAFF NAME:    %s' % staff, align='R', ln=1)
    isl_pdf.cell(w=0, h=10, txt='DATE OF SERVICES:    %s' % from_date.strftime('%m/%d/%y'))
    if not pd.isna(frame['cert_number'].iloc[0]):
        isl_pdf.cell(w=0, h=10, txt='STAFF NO:    %s' % str(int(frame['cert_number'].iloc[0])), align='R', ln=1)
    else:
        isl_pdf.cell(w=0, h=10, txt='STAFF NO: _____', align='R', ln=1)
    isl_pdf.multi_cell(w=0, h=5, txt='CONFIDENTIAL INFORMATION\nCalifornia W & I Code Section 5328\n\n', align='C')
    isl_pdf.dashed_line(95, 38, 205, 38)
    isl_pdf.dashed_line(95, 49, 205, 49)
    isl_pdf.dashed_line(95, 38, 95, 49)
    isl_pdf.dashed_line(205, 38, 205, 49)

    # staff events with clients
    isl_pdf.set_font(family='Arial', size=9)
    isl_pdf.multi_cell(w=30, h=6, txt='Client #\nMedi-Cal', border=1)
    isl_pdf.x = isl_pdf.x + 30
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.multi_cell(w=30, h=6, txt='Client #\nMedicare', border=1)
    isl_pdf.x = isl_pdf.x + 60
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.cell(w=30, h=12, txt='INSYST #', border=1)
    isl_pdf.cell(w=40, h=12, txt='Client Name', border=1)
    isl_pdf.multi_cell(w=40, h=6, txt='County\nProcedure', border=1)
    isl_pdf.x = isl_pdf.x + 170
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.multi_cell(w=20, h=6, txt='Proc#\nDate', border=1)
    isl_pdf.x = isl_pdf.x + 190
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.multi_cell(w=12, h=4, txt='InSyst\nProc.\nCode', border=1)
    isl_pdf.x = isl_pdf.x + 202
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.multi_cell(w=12, h=6, txt='CPT\nCode', border=1)
    isl_pdf.x = isl_pdf.x + 214
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.cell(w=11, h=12, txt='Compl.', border=1)
    isl_pdf.multi_cell(w=10, h=6, txt='Total\nTime', border=1)
    isl_pdf.x = isl_pdf.x + 235
    isl_pdf.y = isl_pdf.y - 12
    vert_col_y = isl_pdf.y + 12
    isl_pdf.multi_cell(w=10, h=6, txt='FTF\nTime', border=1)
    isl_pdf.x = isl_pdf.x + 245
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.multi_cell(w=15, h=6, txt='County\nLoc', border=1)
    isl_pdf.x = isl_pdf.x + 260
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.multi_cell(w=16, h=6, txt='Medicare\nLoc', border=1)
    isl_pdf.x = isl_pdf.x + 276
    isl_pdf.y = isl_pdf.y - 12
    isl_pdf.multi_cell(w=10, h=6, txt='DX\nCode', border=1)

    has_row = False
    for idx, row in frame.iterrows():
        if not pd.isna(row['full_name']):
            vert_col_y = vert_col_y + 12
            if row['vendor_name'] == 'CTYMEDICAL' or row['vendor_name'] == 'STATEMEDICAL':
                isl_pdf.cell(w=30, h=12, txt=row['policy_num'], border=1)
            else:
                isl_pdf.cell(w=30, h=12, txt='', border=1)
            if row['vendor_name'] == 'Medicare':
                isl_pdf.cell(w=30, h=12, txt=row['policy_num'], border=1)
            else:
                isl_pdf.cell(w=30, h=12, txt='', border=1)
            isl_pdf.cell(w=30, h=12, txt=str(int(row['id_no'])), border=1)
            isl_pdf.cell(w=40, h=12, txt=row['full_name'], border=1)
            isl_pdf.cell(w=40, h=12, txt=row['event_name'], border=1)
            isl_pdf.cell(w=20, h=12, txt=row['actual_date'].strftime('%m/%d/%y'), border=1)
            if not pd.isna(row['std_code']):
                isl_pdf.cell(w=12, h=12, txt=str(int(row['std_code'])), border=1)
            else:
                isl_pdf.cell(w=12, h=12, txt='', border=1)
            if not pd.isna(row['sc_code']):
                isl_pdf.cell(w=12, h=12, txt=str(row['sc_code']), border=1)
            else:
                isl_pdf.cell(w=12, h=12, txt='', border=1)
            if not pd.isna(row['COF_Complexitites_2']):
                isl_pdf.cell(w=11, h=12, txt=str(row['COF_Complexitites_2']), border=1)
            else:
                isl_pdf.cell(w=11, h=12, txt='', border=1)
            if pd.isna(row['COF_TOTAL_DURATION']):
                isl_pdf.cell(w=10, h=12, txt='', border=1)
            else:
                isl_pdf.cell(w=10, h=12,
                             txt='{:02d}:{:02d}'.format(*divmod(int(row['COF_TOTAL_DURATION']), 60)),
                             border=1)
            if pd.isna(row['duration_worker']):
                isl_pdf.cell(w=10, h=12, txt='N/A', border=1)
            else:
                isl_pdf.cell(w=10, h=12, txt=row['duration_worker'], border=1)
            if pd.isna(row['general_location']):
                isl_pdf.cell(w=15, h=12, txt='', border=1)
                isl_pdf.cell(w=16, h=12, txt='', border=1)
            else:
                isl_pdf.cell(w=15, h=12, txt=str(int(row['general_location'])), border=1)
                isl_pdf.cell(w=16, h=12, txt=str(get_medicare_loc(int(row['general_location']))), border=1)
            if not pd.isna(row['icd10_code']):
                isl_pdf.cell(w=10, h=12, txt=str(row['icd10_code']), border=1, ln=1)
            else:
                isl_pdf.cell(w=10, h=12, txt='', border=1, ln=1)
            has_row = True
    if not has_row:
        vert_col_y = vert_col_y + 12
        isl_pdf.cell(w=30, h=12, txt='', border=1)
        isl_pdf.cell(w=30, h=12, txt='', border=1)
        isl_pdf.cell(w=30, h=12, txt='', border=1)
        isl_pdf.cell(w=40, h=12, txt='', border=1)
        isl_pdf.cell(w=40, h=12, txt='', border=1)
        isl_pdf.cell(w=20, h=12, txt='', border=1)
        isl_pdf.cell(w=12, h=12, txt='', border=1)
        isl_pdf.cell(w=12, h=12, txt='', border=1)
        isl_pdf.cell(w=11, h=12, txt='', border=1)
        isl_pdf.cell(w=10, h=12, txt='', border=1)
        isl_pdf.cell(w=10, h=12, txt='', border=1)
        isl_pdf.cell(w=15, h=12, txt='', border=1)
        isl_pdf.cell(w=16, h=12, txt='', border=1)
        isl_pdf.cell(w=10, h=12, txt='', border=1, ln=1)

    # staff individual events
    isl_pdf.cell(w=30, h=7, txt='MAA Code', border=1)
    isl_pdf.cell(w=30, h=7, txt='Time', border=1)
    isl_pdf.cell(w=40, h=7, txt='Recipient Code', border=1, ln=1)
    for idx, row in frame.iterrows():
        if pd.isna(row['full_name']):
            if len(row['event_name']) > 15:
                text = row['event_name'][:15] + '...'
                isl_pdf.cell(w=30, h=7, txt=text, border=1)
            else:
                isl_pdf.cell(w=30, h=7, txt=row['event_name'], border=1)
            if pd.isna(row['duration']):
                isl_pdf.cell(w=30, h=7, txt='N/A', border=1)
            else:
                isl_pdf.cell(w=30, h=7, txt='{:02d}:{:02d}'.format(*divmod(int(row['duration']), 60)), border=1)
            if pd.isna(row['rec_code']):
                isl_pdf.cell(w=40, h=7, txt='', border=1, ln=1)
            else:
                isl_pdf.cell(w=40, h=7, txt=str(row['rec_code']), border=1, ln=1)

    isl_pdf.cell(w=30, h=7, txt='Total MAA Time', border=1)
    maa_time = frame['duration'].sum()
    isl_pdf.cell(w=30, h=7, txt='{:02d}:{:02d}'.format(*divmod(int(maa_time), 60)), border=1)

    isl_pdf.x = 215
    isl_pdf.y = vert_col_y
    isl_pdf.cell(w=20, h=10, txt='SUB-TOTAL: ', ln=2, align='R')
    isl_pdf.x = 130
    isl_pdf.multi_cell(w=100, h=5,
                       txt='Enter other staff time already entered in PSP from\n'
                           'Group Attendance Roster or Day Svc Log: ',
                       align='R')
    isl_pdf.x = 130
    isl_pdf.multi_cell(w=100, h=5,
                       txt='Enter your Co-Staff time already entered in PSP\n'
                           'from Primary Staff Log: ',
                       align='R')
    isl_pdf.x = 205
    isl_pdf.cell(w=30, h=10, txt='Enter MAA Time: ', ln=2, align='R')
    isl_pdf.cell(w=30, h=10, txt='Total Paid Time: ', ln=2, align='R')

    isl_pdf.x = 235
    isl_pdf.y = vert_col_y
    subtotal = frame['COF_TOTAL_DURATION'].sum()
    isl_pdf.cell(w=10, h=10, txt='{:02d}:{:02d}'.format(*divmod(int(subtotal), 60)), border=1, ln=2)
    isl_pdf.cell(w=10, h=10, txt='', border=1, ln=2)
    isl_pdf.cell(w=10, h=10, txt='', border=1, ln=2)
    isl_pdf.cell(w=10, h=10, txt='{:02d}:{:02d}'.format(*divmod(int(maa_time), 60)), border=1, ln=2)
    isl_pdf.cell(w=10, h=10, txt='{:02d}:{:02d}'.format(*divmod(int(maa_time + subtotal), 60)), border=1, ln=1)

    isl_pdf.cell(w=0, h=15, txt='I hereby certify, under penalty of perjury, that the information contained in this'
                                ' document is accurate and free from fraudulent claiming.', ln=1)

    isl_pdf.cell(w=150, h=10, txt='Signature:')
    isl_pdf.cell(w=20, h=10, txt='Date:', ln=1)
    isl_pdf.cell(w=200, h=10, txt='_______________________________________________________________________________'
                                  '_________________________________________')
    isl_pdf.output('src/pdf/isl_%s_%s_%s_%s.pdf' %
                   (staff.split(', ')[0].lower(), staff.split(', ')[1].lower(), from_date.strftime('%Y-%m-%d'),
                    num_to_modifier(program_modifier)))
    

def isl(from_date):
    print('Beginning CSV modifications...', end=' ')
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    staff_only = pd.read_csv('src/csv/only_staff.csv')
    staff_ids = pd.read_csv('src/csv/staff_ids.csv')
    clients_only = pd.read_csv('src/csv/clients_only.csv')
    recipient_codes = pd.read_csv('src/csv/recipient_codes.csv')
    other_codes = pd.read_csv('src/csv/other_codes.csv')
    insyst_ids = pd.read_csv('src/csv/client_insyst_ids.csv')
    insurance_info = pd.read_csv('src/csv/insurance_info.csv')

    staff_only = staff_only[['staff_name', 'event_name', 'actual_date', 'duration', 'event_log_id', 'staff_id']]
    staff_only['actual_date'] = pd.to_datetime(staff_only.actual_date)
    staff_only['staff_name'] = staff_only['staff_name'].str.strip()
    needed_staff = ['Weber, Ihande', 'Manjunath, Sudha', 'Nelson, Britta', 'Kapis, Kelly', 'Lau, Michael',
                    'Guiao, Christine', 'Carrell, Ella', 'Tran, Lan Anh', 'Awana, Jaime', 'Broyles, Rachel']
    staff_only = staff_only[staff_only['staff_name'].isin(needed_staff)]
    staff_only.drop_duplicates(subset=['event_log_id'], inplace=True)

    clients_only = clients_only[['full_name', 'id_no', 'staff_name', 'actual_date', 'event_log_id', 'event_name',
                                 'duration', 'general_location', 'program_modifier_code', 'event_category_id',
                                 'staff_id']]
    clients_only['actual_date'] = pd.to_datetime(clients_only.actual_date)
    clients_only = clients_only.rename(columns={'duration': 'duration_worker'})
    clients_only['program_modifier_code'] = clients_only['program_modifier_code'].str.strip()
    clients_only = clients_only[clients_only['program_modifier_code'].isin(['RR', 'SMMH', 'SMHAD'])]
    clients_only = clients_only[clients_only['event_category_id'] == '4b9aebb1-34d7-4a06-b22f-1491fb725d8c']
    clients_only.reset_index(drop=True, inplace=True)
    clients_only['general_location'] = clients_only['general_location'].apply(lambda v: get_loc_code(v))

    insyst_ids = insyst_ids.rename(columns={'Client ID': 'id_no'})
    insyst_ids['id_no'] = insyst_ids['id_no'].astype(int)
    clients_only['id_no'] = clients_only['id_no'].astype(int)
    clients_only = clients_only.merge(insyst_ids, on=['id_no'], how='left')
    clients_only = clients_only.merge(other_codes, on=['event_log_id'], how='left')

    insurance_info.drop_duplicates(inplace=True)
    insurance_info = insurance_info.rename(columns={'Client ID': 'id_no'})
    insurance_info['id_no'] = insurance_info['id_no'].astype(int)
    clients_only = clients_only.merge(insurance_info, on=['id_no'], how='left')

    staff_only = staff_only.merge(recipient_codes, on=['event_log_id'])

    merged = pd.concat([staff_only, clients_only], axis=0, sort=False, ignore_index=True)
    merged = merged.merge(staff_ids, on=['staff_id'], how='left')
    merged['program_modifier_code'] = merged['program_modifier_code'].fillna('maa')
    merged['program_modifier_code'] = merged['program_modifier_code'].apply(lambda v: modifier_to_num(v))
    print(merged)
        
    for key, frame in merged.groupby(['staff_name', 'program_modifier_code']):
        print(frame)
        create_isl(frame, key[0], key[1], from_date)
    
    print('Done.')


def get_loc_code(val):
    val = str(val).strip()
    if val == 'Telehealth':
        return 20
    elif val == 'Phone':
        return 3
    elif val == 'Agency Office':
        return 1
    elif 'Homeless' in val:
        return 10
    elif val == 'Home':
        return 4
    elif not pd.isna(val):
        return 18
    return ''


def get_medicare_loc(val):
    if val == 1:
        return 11
    elif val == 4:
        return 12
    elif val == 10:
        return 4
    elif val == 20:
        return 2
    elif val == 18:
        return 99
    return ''


def modifier_to_num(val):
    val = val.lower()
    if val == 'smmh':
        return 1
    elif val == 'smhad':
        return 2
    elif val == 'rr':
        return 3
    elif val == 'maa':
        return 1
    return ''


def num_to_modifier(val):
    if val == 1:
        return 'smmh-maa'
    elif val == 2:
        return 'smhad'
    elif val == 3:
        return 'rr'
    else:
        return ''


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

    driver.get('https://myevolvcofhsxb.netsmartcloud.com/')

    # login
    with open('src/config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['fremont']
        pwd = login['pwd']
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[1]/div/div[1]/div[1]/input').send_keys(usr)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[1]/div/input[4]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[1]/div/div[1]/div[2]/div/input') \
        .send_keys(pwd)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[1]/div/input[7]').click()

    # navigate to and generate canned staff idv events report (only_staff.csv)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[19]/span').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[1]/ul/li[6]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[2]/ul/li[9]').click()
    cd_frame1 = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[2]/div/div[17]/div/div/div/div/iframe')
    driver.switch_to.frame(cd_frame1)
    cd_frame2 = driver.find_element_by_xpath('/html/body/form/div[3]/div/div[2]/iframe[8]')
    driver.switch_to.frame(cd_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[2]/div/input') \
        .send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[3]/div/input') \
        .send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[3]/div[2]/div/span').click()

    # switch back to default content for report selection
    sleep(1)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div[3]/div/div/div[1]/table/tbody/tr[3]') \
        .click()

    # download and rename the report
    driver.implicitly_wait(5)
    driver.switch_to.frame(cd_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cd_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/ul/li[17]/a').click()
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/only_staff.csv')

    # navigate to and generate canned client services report (clients_only.csv)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[19]/span').click()
    driver.find_element_by_xpath('//*[@id="product-header-mega-menu-level1-id"]/li[3]').click()
    driver.find_element_by_xpath('//*[@id="d66dc6ec-03be-41b2-b808-206523c7e33d"]/li[2]').click()
    cd_frame1 = driver.find_element_by_xpath('//*[@id="MainContent_ctl36"]/iframe')
    driver.switch_to.frame(cd_frame1)
    cd_frame2 = driver.find_element_by_xpath('//*[@id="formset-target-wrapper-id"]/div[2]/iframe[1]')
    driver.switch_to.frame(cd_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[2]/div/input') \
        .send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[3]/div/input') \
        .send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="rightContentFormId-ctrl-5"]/span').click()

    # switch back to default content for report selection
    sleep(1)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/div[2]/div/div/div[2]/div/div[3]/div/div/div[1]/table/tbody/tr[3]') \
        .click()

    # download and rename the report
    driver.implicitly_wait(5)
    driver.switch_to.frame(cd_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cd_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="form-toolbar-16"]').click()
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/clients_only.csv')

    # navigate to custom reports > -RPA-
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[19]/span').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[1]/ul/li[2]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[2]/ul[1]/li[2]').click()
    cr_frame1 = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[2]/div/div[17]/div/div/div/div/iframe')
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    cr_frame2 = driver.find_element_by_xpath('/html/body/form/div[3]/div/div[2]/iframe')
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="grdMain_FolderLink_0"]').click()
    driver.implicitly_wait(5)
    sleep(3)

    # 1 Individual Staff Events Recipient Codes
    driver.find_element_by_id('grdMain_ObjectName_0').click()
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_3A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_3B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()

    # download and rename the report
    driver.implicitly_wait(5)
    driver.find_element_by_id('CSV').click()
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/recipient_codes.csv')

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)
    sleep(3)

    # 2 Staff IDs
    driver.find_element_by_id('grdMain_ObjectName_1').click()
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(10)

    # download and rename the report
    driver.find_element_by_id('CSV').click()
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/staff_ids.csv')

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)
    sleep(3)

    # 3 CPT/Proc/ICD-10 Codes & Complexities
    driver.find_element_by_id('grdMain_ObjectName_2').click()
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('//*[@id="RP1_1A"]').send_keys(from_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="RP1_1B"]').send_keys(to_date.strftime('%m/%d/%Y'))
    driver.find_element_by_xpath('//*[@id="Submit"]').click()

    # download and rename the report
    driver.implicitly_wait(5)
    driver.find_element_by_id('CSV').click()
    driver.implicitly_wait(5)
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/other_codes.csv')

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)
    sleep(3)

    # 4 Client INSYST IDs
    driver.find_element_by_id('grdMain_ObjectName_3').click()
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(10)

    # download and rename the report
    driver.find_element_by_id('CSV').click()
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/client_insyst_ids.csv')

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cr_frame2)
    driver.implicitly_wait(5)
    sleep(3)

    # 5 Client insurance info
    driver.find_element_by_id('grdMain_ObjectName_4').click()
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[-1])
    driver.implicitly_wait(10)

    # download and rename the report
    driver.find_element_by_id('CSV').click()
    sleep(3)
    filename = max(['src/csv' + '/' + f for f in os.listdir('src/csv')], key=os.path.getctime)
    shutil.move(filename, 'src/csv/insurance_info.csv')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def fremont_isl(from_date):
    print('Running ISL report for ' + from_date.strftime('%Y.%m.%d'))
    to_date = from_date + timedelta(days=1)

    # browser(from_date, to_date)
    isl(from_date)
    folder_path = 'src/%s' % from_date.strftime('%Y-%m-%d')
    os.mkdir(folder_path)
    for filename in os.listdir('src/pdf'):
        shutil.move('src/pdf/%s' % filename, folder_path)
    upload_folder(folder_path, '1lYsW4yfourbnFYJB3GLh6br7D1_3LOcd')
    
    for filename in os.listdir('src/csv'):
        if not filename.endswith('.xlsx'):
            os.remove('src/csv/%s' % filename)
    for filename in os.listdir('src/pdf'):
        os.remove('src/pdf/%s' % filename)
    shutil.rmtree(folder_path)


def main():
    print('------------------------------ ' + datetime.now().strftime('%Y.%m.%d %H:%M') +
          ' ------------------------------')
    # only run automation for workdays
    f = open('src/txt/most_recent_from_date.txt', 'r+')
    from_date = date.today() - timedelta(days=5)
    fremont_isl(from_date)
    print('Beginning Fremont ISL RPA (%s)...' % from_date.strftime('%Y.%m.%d'))
    if from_date.weekday() < 6:
        today = date.today()
        # if second workday of month, run automation for the rest of the previous month
        if 1 < int(today.strftime('%d')) <= 5 and today.weekday() < 5:
            print('Second workday of the month -- running remaining reports for previous month')
            from_date = pd.to_datetime(f.read().strip()) + timedelta(days=1)
            from_day = int(from_date.strftime('%d'))
            prev_month_last_day = date.today().replace(day=1) - timedelta(days=1)
            if from_date < prev_month_last_day:
                for i in range(from_day, int(prev_month_last_day.strftime('%d'))+1):
                    if from_date.weekday() < 6:
                        fremont_isl(from_date)
                        f.truncate(0)
                        f.write(from_date.strftime('%Y-%m-%d').strip())
                    from_date = from_date + timedelta(days=1)
            else:
                print('Already ran all reports from previous month.')
        else:
            fremont_isl(from_date)
            f.truncate(0)
            f.write(from_date.strftime('%Y-%m-%d').strip())
    else:
        print('%s is on a Sunday.' % from_date.strftime('%Y.%m.%d'))
        f.truncate(0)
        f.write(from_date.strftime('%Y-%m-%d').strip())
    
    print('Successfully finished Fremont ISL RPA!')


try:
    main()
    send_gmail('eanderson@khitconsulting.com',
               'KHIT Report Notification',
               'Successfully finished Fremont ISL RPA!')
except Exception as e:
    print('System encountered an error running Fremont ISL RPA:\n')
    print_exc()
    email_body = 'System encountered an error running Fremont ISL RPA: %s' % e
    send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)
