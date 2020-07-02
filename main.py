from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
from openpyxl import load_workbook

MANAGERS = [
    'Aaron Nolan',
    'Laurence Murphy',
    'Sebastian Staszak',
    'Nicole Kelly',
    'Mellissa Kelly',
    'Dumitru Popa',
    'aoife Byrne',
    'Joe Dudley',
    'Katarzyna Hajrych',
    'Stephanie Rooney',
    'Matthew Taylor',
    'Edyta Zawada'
]

CREW_TRAINERS = [
    'Caoimhe Kelly',
    'Emma Power',
    'Barry Doyle',
    'Nicole O\'Brien',
    'Lauren Moore',
    'Martell Cullen',
    'Robyn Moore Keogh',
    'William Nolan',
    'Raphael Vieira'
]

# METHODS TO CREATE A .XLSX FROM MYSCHEDULE
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

    wb2 = Workbook()
    wb2.create_sheet('Crew Members')
    wb2.create_sheet('Crew Trainers')
    ws = wb2.active
    wb2.remove(ws)
    wb2.save('training_planner.xlsx')
    print('training_planner.xlsx created successfully')

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

# METHOD TO CREATE TRAINING PLANNER FILE
def schedule_to_planner():
    schedule_wb = load_workbook('schedule.xlsx')
    planner_wb = load_workbook('training_planner.xlsx')

    schedule_ws = schedule_wb.active
    cm_planner_ws = planner_wb.get_sheet_by_name('Crew Members')
    ct_planner_ws = planner_wb.get_sheet_by_name('Crew Trainers')

    info = []

    for row in schedule_ws.iter_rows(values_only=True):
        info.append([row[0].strip(), row[4].strip(), row[5].strip(), row[6].strip(), row[7].strip(), row[8].strip(), row[9].strip(), row[10].strip()])

    for i in enumerate(info, start=1):
        for i2 in enumerate(info[i], start=1):
            
            if info[i-1][i2-1][0].isnumeric():
                pass
    


# METHODS TO HELP READ IN SOC INFORMATION FROM SUGGESTED_SOCS.XLSX
def add_suggested_socs():
    sugg_soc_wb = load_workbook('Suggested_SOCs.xlsx')
    schedule_wb = load_workbook('schedule.xlsx')

    sugg_soc_ws = sugg_soc_wb.active
    schedule_ws = schedule_wb.active

    soc_dict = {}

    for row in sugg_soc_ws.iter_rows(min_col=2, max_col=5, min_row=2, values_only=True):
        soc_dict[row[0]] = [row[1], row[2], row[3]]
        #print(row)
    print(soc_dict)

    for index, row in enumerate(schedule_ws.iter_rows(values_only=True), start=1):
        if row[0] in soc_dict.keys():
            schedule_ws.cell(row=index, column=12).value = soc_dict[row[0]][0]
            schedule_ws.cell(row=index, column=13).value = soc_dict[row[0]][1]
            schedule_ws.cell(row=index, column=14).value = soc_dict[row[0]][2]

    #schedule_ws.cell(row=1, column=12).value = soc_dict[row[0]][0]
    schedule_wb.save('schedule.xlsx')
    

#schedule_to_excel()
#add_suggested_socs()
schedule_to_planner()