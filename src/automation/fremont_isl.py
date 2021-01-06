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


def has_clients(df):
    for staff, frame in df.groupby('staff_name'):
        print(frame)
        isl_pdf = FPDF(orientation='L')
        isl_pdf.add_page()
        isl_pdf.set_font(family='Arial', size=11)
        isl_pdf.cell(w=0, h=7, txt='ALAMEDA COUNTY BEHAVIORAL HEALTH CARE - MENTAL HEALTH', align='C', ln=2)
        isl_pdf.cell(w=0, h=15, txt='INDIVIDUAL STAFF LOG', align='C', ln=2)
        isl_pdf.cell(w=0, h=10, txt='REPORTING UNIT: 01EI1 - City of Fremont MMH', align='L', ln=0)
        isl_pdf.cell(w=0, h=15, txt='STAFF NAME: %s' % staff, align='R', ln=1)
        isl_pdf.multi_cell(w=0, h=5, txt='CONFIDENTIAL INFORMATION\nCalifornia W & I Code Section 5328\n\n', align='C')
        isl_pdf.dashed_line(90, 46, 210, 46)
        isl_pdf.dashed_line(90, 57, 210, 57)
        isl_pdf.dashed_line(90, 46, 90, 57)
        isl_pdf.dashed_line(210, 46, 210, 57)

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
        isl_pdf.multi_cell(w=30, h=6, txt='County\nProcedure', border=1)
        isl_pdf.x = isl_pdf.x + 160
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=20, h=6, txt='Proc#\nDate', border=1)
        isl_pdf.x = isl_pdf.x + 180
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=15, h=4, txt='InSyst\nProc.\nCode', border=1)
        isl_pdf.x = isl_pdf.x + 195
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=20, h=6, txt='CPT\nCode', border=1)
        isl_pdf.x = isl_pdf.x + 215
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.cell(w=15, h=12, txt='Compl.', border=1)
        isl_pdf.multi_cell(w=15, h=6, txt='Total\nTime', border=1)
        isl_pdf.x = isl_pdf.x + 245
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=15, h=6, txt='FTF\nTime', border=1)
        isl_pdf.x = isl_pdf.x + 260
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=15, h=6, txt='County\nLoc', border=1)
        isl_pdf.x = isl_pdf.x + 275
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=15, h=6, txt='Medicare\nLoc', border=1)
        isl_pdf.x = isl_pdf.x + 290
        isl_pdf.y = isl_pdf.y - 12
        isl_pdf.multi_cell(w=15, h=12, txt='DX\nCode', border=1)
        for idx, row in frame.iterrows():
            if not pd.isna(row['Full Name']):
                if row['vendor_name'] == 'Medi-Cal':
                    isl_pdf.cell(w=40, h=12, txt=row['policy_num'], border=1)
                else:
                    isl_pdf.cell(w=40, h=12, txt='', border=1)
                if row['vendor_name'] == 'Medicare':
                    isl_pdf.cell(w=40, h=12, txt=row['policy_num'], border=1)
                else:
                    isl_pdf.cell(w=40, h=12, txt='', border=1)
                isl_pdf.cell(w=30, h=12, txt=str(row['insyst']), border=1)
                isl_pdf.cell(w=40, h=12, txt=row['Full Name'], border=1)
                isl_pdf.cell(w=40, h=12, txt=row['event_name'], border=1)
                isl_pdf.cell(w=30, h=12, txt=row['service_date'].strftime('%m/%d/%y'), border=1)
                isl_pdf.cell(w=20, h=12, txt='', border=1)
                isl_pdf.cell(w=20, h=12, txt='', border=1)
                isl_pdf.cell(w=15, h=12, txt='', border=1)
                isl_pdf.cell(w=20, h=12, txt=str(row['duration_worker']), border=1)
                isl_pdf.cell(w=20, h=12, txt=str(row['duration_client']), border=1)
                isl_pdf.cell(w=20, h=12, txt=str(row['sc_code']), border=1)
                isl_pdf.cell(w=20, h=12, txt='', border=1)
                isl_pdf.cell(w=20, h=12, txt='', border=1, ln=1)

        # staff individual events
        isl_pdf.set_font(family='Arial', size=11)
        isl_pdf.cell(w=0, h=12, txt='', ln=1)
        isl_pdf.cell(w=40, h=7, txt='MAA Code', border=1)
        isl_pdf.cell(w=40, h=7, txt='Time', border=1)
        isl_pdf.cell(w=40, h=7, txt='Recipient Code', border=1, ln=1)
        for idx, row in frame.iterrows():
            if pd.isna(row['Full Name']):
                isl_pdf.cell(w=40, h=7, txt=row['event_name'], border=1)
                isl_pdf.cell(w=40, h=7, txt=str(row['duration']), border=1)
                isl_pdf.cell(w=40, h=7, txt=row['actual_date'].strftime('%m/%d/%y %I:%M %p'), border=1, ln=1)
        isl_pdf.cell(w=40, h=7, txt='Total MAA Time', border=1)
        isl_pdf.cell(w=40, h=7, txt=str(frame['duration'].sum()), border=1)

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
    both = pd.concat([staff_only, clients_only], axis=0, ignore_index=True)
    both.to_csv('../csv/merged.csv')
    if not both.empty:
        has_clients(both)

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
    # options.add_argument('--headless')
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
        usr = login['abhs']
        pwd = login['pwd']

    print('Exiting chromedriver...', end=' ')
    driver.close()
    driver.quit()
    print('Process killed.')


def main():
    print('------------------------------' + date.today().strftime('%Y.%m.%d') + '------------------------------')
    print('Beginning ABHS Client Services RPA...')
    from_date = date.today()
    to_date = date.today() + timedelta(days=1)

    # browser(from_date, to_date)
    isl()
    # upload_folder(folder_path, '1h_Mym7ocK5lJ_-a4eZzShQf4DGm6HA8C')
    # email_body = "Your monthly service entry reports (%s) are ready and available on the ABHS RPA " \
    #              "Reports shared drive: https://drive.google.com/drive/folders/1h_Mym7ocK5lJ_-a4eZzShQf4DGm6HA8C" \
    #              % folder_path.split('/')[-1]
    # send_gmail('amanda.bruns@wmabhs.org', 'KHIT Report Notification', email_body)
    # send_gmail('kelly.moffett-place@wmabhs.org', 'KHIT Report Notification', email_body)
    #
    # os.remove('src/csv/report_bruns.csv')
    # os.remove('src/csv/report_moffett.csv')

    print('Successfully finished ABHS Client Services RPA!')


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
