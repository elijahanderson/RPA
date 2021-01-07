import os
import pandas as pd
import shutil
import sys
import yaml

sys.path[0] = '/home/eanderson/RPA/src'

from datetime import date, timedelta
from fpdf import FPDF
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep

from infrastructure.drive_upload import upload_folder
from infrastructure.email import send_gmail


def create_isls(df):
    for staff, frame in df.sort_values(['Full Name']).groupby('staff_name'):
        isl_pdf = FPDF(orientation='L')
        isl_pdf.add_page()
        isl_pdf.set_font(family='Arial', size=11)
        isl_pdf.cell(w=0, h=4, txt='ALAMEDA COUNTY BEHAVIORAL HEALTH CARE - MENTAL HEALTH', align='C', ln=2)
        isl_pdf.cell(w=0, h=5, txt='INDIVIDUAL STAFF LOG', align='C', ln=2)
        isl_pdf.cell(w=0, h=10, txt='REPORTING UNIT:    01EI1 - City of Fremont MMH', align='L', ln=0)
        isl_pdf.cell(w=0, h=10, txt='STAFF NAME:    %s' % staff, align='R', ln=1)
        isl_pdf.cell(w=0, h=10, txt='DATE OF SERVICES:    %s' % '12/28/2020')  # TODO -- date.today().strftime('%m/%d/%y'))
        isl_pdf.cell(w=0, h=10, txt='STAFF NO: ________________', align='R', ln=1)
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
        isl_pdf.cell(w=30, h=12, txt='Client Name', border=1)
        isl_pdf.multi_cell(w=30, h=6, txt='County\nProcedure', border=1)
        isl_pdf.x = isl_pdf.x + 150
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=20, h=6, txt='Proc#\nDate', border=1)
        isl_pdf.x = isl_pdf.x + 170
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=15, h=4, txt='InSyst\nProc.\nCode', border=1)
        isl_pdf.x = isl_pdf.x + 185
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=10, h=6, txt='CPT\nCode', border=1)
        isl_pdf.x = isl_pdf.x + 195
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.cell(w=11, h=12, txt='Compl.', border=1)
        isl_pdf.multi_cell(w=10, h=6, txt='Total\nTime', border=1)
        isl_pdf.x = isl_pdf.x + 216
        isl_pdf.y = isl_pdf.y - 12
        vert_col_y = isl_pdf.y + 24
        isl_pdf.multi_cell(w=10, h=6, txt='FTF\nTime', border=1)
        isl_pdf.x = isl_pdf.x + 226
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=15, h=6, txt='County\nLoc', border=1)
        isl_pdf.x = isl_pdf.x + 241
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=16, h=6, txt='Medicare\nLoc', border=1)
        isl_pdf.x = isl_pdf.x + 257
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=10, h=6, txt='DX\nCode', border=1)

        has_row = False
        for idx, row in frame.iterrows():
            if not pd.isna(row['Full Name']):
                if row['vendor_name'] == 'Medi-Cal':
                    isl_pdf.cell(w=30, h=12, txt=row['policy_num'], border=1)
                else:
                    isl_pdf.cell(w=30, h=12, txt='', border=1)
                if row['vendor_name'] == 'Medicare':
                    isl_pdf.cell(w=30, h=12, txt=row['policy_num'], border=1)
                else:
                    isl_pdf.cell(w=30, h=12, txt='', border=1)
                isl_pdf.cell(w=30, h=12, txt=str(int(row['insyst'])), border=1)
                name = row['Full Name'].split(',')
                isl_pdf.multi_cell(w=30, h=6, txt=name[0] + ',\n' + name[1], border=1)
                isl_pdf.x = isl_pdf.x + 120
                isl_pdf.y = isl_pdf.y - 12
                if len(row['event_name']) > 15:
                    isl_pdf.multi_cell(w=30, h=6,
                                       txt=row['event_name'][:len(row['event_name']) // 2] + '-\n' +
                                           row['event_name'][len(row['event_name']) // 2:len(row['event_name'])],
                                       border=1)
                    isl_pdf.x = isl_pdf.x + 150
                    isl_pdf.y = isl_pdf.y - 12
                else:
                    isl_pdf.cell(w=30, h=12, txt=row['event_name'], border=1)
                isl_pdf.cell(w=20, h=12, txt=row['service_date'].strftime('%m/%d/%y'), border=1)
                if not pd.isna(row['insyst_proc_code']):
                    isl_pdf.cell(w=15, h=12, txt=str(int(row['insyst_proc_code'])), border=1)
                else:
                    isl_pdf.cell(w=15, h=12, txt='', border=1)
                if not pd.isna(row['cpt_code']):
                    isl_pdf.cell(w=10, h=12, txt=str(int(row['cpt_code'])), border=1)
                else:
                    isl_pdf.cell(w=10, h=12, txt='', border=1)
                if not pd.isna(row['complexities']):
                    isl_pdf.cell(w=11, h=12, txt=str(row['complexities']), border=1)
                else:
                    isl_pdf.cell(w=11, h=12, txt='', border=1)
                if pd.isna(row['duration_worker']):
                    isl_pdf.cell(w=10, h=12, txt='N/A', border=1)
                else:
                    isl_pdf.cell(w=10, h=12, txt=str(int(row['duration_worker'])), border=1)
                if pd.isna(row['duration_client']):
                    isl_pdf.cell(w=10, h=12, txt='N/A', border=1)
                else:
                    isl_pdf.cell(w=10, h=12, txt=str(int(row['duration_client'])), border=1)
                if pd.isna(row['sc_code']):
                    isl_pdf.cell(w=15, h=12, txt='N/A', border=1)
                else:
                    isl_pdf.cell(w=15, h=12, txt=str(int(row['sc_code'])), border=1)
                isl_pdf.cell(w=16, h=12, txt='', border=1)
                isl_pdf.cell(w=10, h=12, txt='', border=1, ln=1)
                has_row = True
        if not has_row:
            isl_pdf.cell(w=30, h=12, txt='', border=1)
            isl_pdf.cell(w=30, h=12, txt='', border=1)
            isl_pdf.cell(w=30, h=12, txt='', border=1)
            isl_pdf.cell(w=30, h=12, txt='', border=1)
            isl_pdf.cell(w=30, h=12, txt='', border=1)
            isl_pdf.cell(w=20, h=12, txt='', border=1)
            isl_pdf.cell(w=15, h=12, txt='', border=1)
            isl_pdf.cell(w=10, h=12, txt='', border=1)
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
            if pd.isna(row['Full Name']):
                isl_pdf.cell(w=30, h=7, txt=row['event_name'], border=1)
                if pd.isna(row['duration']):
                    isl_pdf.cell(w=30, h=7, txt='N/A', border=1)
                else:
                    isl_pdf.cell(w=30, h=7, txt=str(int(row['duration'])), border=1)
                isl_pdf.cell(w=40, h=7, txt='', border=1, ln=1)
        isl_pdf.cell(w=30, h=7, txt='Total MAA Time', border=1)
        maa_time = frame['duration'].sum()
        isl_pdf.cell(w=30, h=7, txt=str(int(maa_time)), border=1)

        isl_pdf.x = 196
        isl_pdf.y = vert_col_y
        isl_pdf.cell(w=20, h=10, txt='SUB-TOTAL: ', ln=2, align='R')
        isl_pdf.x = 111
        isl_pdf.multi_cell(w=100, h=5,
                           txt='Enter other staff time already entered in PSP from\n'
                               'Group Attendance Roster or Day Svc Log: ',
                           align='R')
        isl_pdf.x = 111
        isl_pdf.multi_cell(w=100, h=5,
                           txt='Enter your Co-Staff time already entered in PSP\n'
                               'from Primary Staff Log: ',
                           align='R')
        isl_pdf.x = 186
        isl_pdf.cell(w=30, h=10, txt='Enter MAA Time: ', ln=2, align='R')
        isl_pdf.cell(w=30, h=10, txt='Total Paid Time: ', ln=2, align='R')

        isl_pdf.x = 216.25
        isl_pdf.y = vert_col_y
        subtotal = frame['duration_worker'].sum()
        isl_pdf.cell(w=9.75, h=10, txt=str(int(subtotal)), border=1, ln=2)
        isl_pdf.cell(w=9.75, h=10, txt='', border=1, ln=2)
        isl_pdf.cell(w=9.75, h=10, txt='', border=1, ln=2)
        isl_pdf.cell(w=9.75, h=10, txt=str(int(maa_time)), border=1, ln=2)
        isl_pdf.cell(w=9.75, h=10, txt=str(int(maa_time + subtotal)), border=1, ln=1)

        isl_pdf.cell(w=0, h=15, txt='I hereby certify, under penalty of perjury, that the information contained in this'
                                    ' document is accurate and free from fraudulent claiming.', ln=1)

        isl_pdf.cell(w=150, h=10, txt='Signature:')
        isl_pdf.cell(w=20, h=10, txt='Date:', ln=1)
        isl_pdf.cell(w=200, h=10, txt='_______________________________________________________________________________'
                                      '_________________________________________')
        # TODO change to src/csv/[].pdf
        isl_pdf.output('../pdf/isl_%s.pdf' % staff.split(',')[0].lower())
    return df


def isl():
    print('Beginning CSV modifications...', end=' ')
    # TODO change to src/csv/[].csv
    staff_only = pd.read_csv('../csv/only_staff.csv')
    clients_only = pd.read_csv('../csv/clients_only.csv')

    staff_only = staff_only[['staff_name', 'event_name', 'actual_date', 'duration']]
    staff_only['actual_date'] = pd.to_datetime(staff_only.actual_date)
    staff_only.drop_duplicates(subset=['actual_date'], inplace=True)
    staff_only['staff_name'] = staff_only['staff_name'].str.strip()

    clients_only['service_date'] = pd.to_datetime(clients_only.service_date)
    clients_only['staff_name'] = clients_only['staff_name'].str.strip()
    # clients_only = clients_only.rename(columns={'staff_name': 'staff_name_co'})

    # to deal with any staff that have both individual events and client events
    merged = pd.concat([staff_only, clients_only], axis=0, ignore_index=True)
    merged.to_csv('../csv/merged.csv')  # TODO change to src/csv/...
    create_isls(merged)

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
        'download.default_directory': 'D:\\Programming\\KHIT\\RPA\\src\\csv',
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
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': 'D:\\Programming\\KHIT\\RPA\\src\\csv'}}
    driver.execute('send_command', params)
    print('Done.')

    driver.get('https://myevolvcofhsxb.netsmartcloud.com/')

    # login
    with open('../config/login.yml', 'r') as yml:
        login = yaml.safe_load(yml)
        usr = login['abhs']
        pwd = login['pwd']
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[1]/div/div[1]/div[1]/input').send_keys(usr)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[1]/div/input[4]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[1]/div/div[1]/div[2]/div/input')\
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
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[2]/div/input')\
        .send_keys('12/28/2020')  # TODO replace with proper date
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[2]/div[3]/div/input')\
        .send_keys('12/29/2020')
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[5]/div/div/div/div[3]/div[2]/div/span').click()

    # switch back to default content for report selection
    sleep(1)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div[3]/div/div/div[1]/table/tbody/tr[3]')\
        .click()

    # download and rename report
    driver.switch_to.frame(cd_frame1)
    driver.implicitly_wait(5)
    driver.switch_to.frame(cd_frame2)
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/ul/li[17]/a').click()
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getctime)
    shutil.move(filename, '../csv/only_staff.csv')

    # navigate to and generate custom ISL report (clients_only.csv)
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/ul/li[19]/span').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[1]/ul/li[2]').click()
    driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[1]/div[5]/div/div[2]/ul[1]/li[2]').click()
    cr_frame1 = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[2]/div/div[17]/div/div/div/div/iframe')
    driver.switch_to.frame(cr_frame1)
    cr_frame2 = driver.find_element_by_xpath('/html/body/form/div[3]/div/div[2]/iframe')
    driver.switch_to.frame(cr_frame2)
    driver.find_element_by_xpath(
        '/html/body/form/table/tbody/tr/td/table/tbody/tr/td/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]/a').click()
    driver.implicitly_wait(5)
    driver.find_element_by_id('grdMain_ObjectName_0').click()
    driver.switch_to.default_content()
    driver.switch_to.default_content()
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element_by_xpath(
        '/html/body/form/div[2]/span/div/table/tbody/tr[1]/td/span/table/tbody/tr[1]/td[2]/span/input[1]')\
        .send_keys('12/28/2020')  # TODO replace both dates w/ from/to_date
    driver.find_element_by_xpath(
        '/html/body/form/div[2]/span/div/table/tbody/tr[1]/td/span/table/tbody/tr[1]/td[3]/span/input[1]') \
        .send_keys('12/29/2020')
    driver.find_element_by_xpath('/html/body/form/div[2]/span/div/table/tbody/tr[2]/td/a[1]/input').click()

    # download and rename the report
    driver.find_element_by_xpath('/html/body/form/span[5]/span/rdcondelement5/span/a/img').click()
    sleep(3)
    filename = max(['../csv' + '/' + f for f in os.listdir('../csv')], key=os.path.getctime)
    shutil.move(filename, '../csv/clients_only.csv')

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print('------------------------------' + date.today().strftime('%Y.%m.%d') + '------------------------------')
    print('Beginning Fremont ISL RPA...')
    from_date = date.today()
    to_date = date.today() + timedelta(days=1)

    browser(from_date, to_date)
    isl()
    folder_path = '../%s' % from_date.strftime('%m-%d-%Y')
    os.mkdir(folder_path)
    for filename in os.listdir('../csv'):
        shutil.move('../csv/%s' % filename, folder_path)
    upload_folder(folder_path, '1lYsW4yfourbnFYJB3GLh6br7D1_3LOcd')
    email_body = "Your daily ISL reports for (%s) are ready and available on the Fremont RPA " \
                 "Reports shared drive: https://drive.google.com/drive/folders/1h_Mym7ocK5lJ_-a4eZzShQf4DGm6HA8C" \
                 % from_date.strftime('%m-%d-%Y')
    send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)

    for filename in os.listdir('../csv'):
        os.remove(filename)
    for filename in os.listdir('../pdf'):
        os.remove(filename)
    shutil.rmtree(folder_path)

    print('Successfully finished Fremont ISL RPA!')


main()
# try:
#     main()
#     send_gmail('eanderson@khitconsulting.com',
#                'KHIT Report Notification',
#                'Successfully finished ABHS Client Services RPA!')
# except Exception as e:
#     print('System encountered an error running ABHS Service Entry RPA: %s' % e)
#     email_body = 'System encountered an error running ABHS Service Entry RPA: %s' % e
#     send_gmail('eanderson@khitconsulting.com', 'KHIT Report Notification', email_body)
