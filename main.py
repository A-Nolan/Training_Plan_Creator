from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook


def schedule_to_excel():
    #input users id and password
    user = input('User ID: ')
    password = input('Password: ')
    week_num = input('Week Number: ')

    driver = open_schedule(user, password, week_num)
    schedule_td = driver.find_elements_by_css_selector('#gridbox > table > tbody > tr:nth-child(2) > td > div > div > table > tbody > tr > td')
    index = 0

    wb = Workbook()
    ws = wb.active

    for col in range(1, (int(len(schedule_td) / 11)) + 1):
        for row in range(1, 12):
            info = schedule_td[index].text
            if row == 1:
                name = rearrange_name(info)
                info = name
            curr_cell = ws.cell(col, row)
            curr_cell.value = info
            index += 1

    wb.save('schedule.xlsx')
    print('schedule.xlsx created successfully')

def open_schedule(user, password, week_num):

    options = Options()
    options.add_argument('headless')

    driver = webdriver.Chrome(options=options, executable_path='C:\Program Files (x86)\chromedriver')
    driver.get('https://psschedule.reflexisinc.co.uk/wfmmcdirlprd/ModuleSelection.jsp')

    # Log in
    userid_input_box = driver.find_element_by_name("txtUserID")
    password_input_box = driver.find_element_by_name('txtPassword')
    userid_input_box.send_keys(user)
    password_input_box.send_keys(password)
    login_button = driver.find_element_by_class_name('button-t')
    login_button.click()

    driver.get(f'https://psschedule.reflexisinc.co.uk/wfmmcdirlprd/rws/schedule/schedule_weekly.jsp?sm=0&mm=SCHD&cboYear=2020&cboQuarter=3&cboMonth=7&cboWeek={week_num}&showYearDropDown=Y')

    return driver

def rearrange_name(name):
    name_l = name.split()
    name_l.remove(',')
    name_l.insert(0, name_l.pop())
    return ' '.join(name_l)

schedule_to_excel()