from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# input users id and password
#userid = input('User ID: ')
#password = input('Password: ')

options = Options()
options.add_argument('headless')

driver = webdriver.Chrome(options=options, executable_path='C:\Program Files (x86)\chromedriver')
driver.get('https://psschedule.reflexisinc.co.uk/wfmmcdirlprd/ModuleSelection.jsp')

# Log in
userid_input_box = driver.find_element_by_name("txtUserID")
password_input_box = driver.find_element_by_name('txtPassword')
userid_input_box.send_keys(userid)
password_input_box.send_keys(password)
login_button = driver.find_element_by_class_name('button-t')
login_button.click()

driver.get('https://psschedule.reflexisinc.co.uk/wfmmcdirlprd/rws/schedule/schedule_weekly.jsp?sm=0&mm=SCHD&cboYear=2020&cboQuarter=3&cboMonth=7&cboWeek=27&showYearDropDown=Y')
rows = driver.find_elements_by_css_selector('#gridbox > table > tbody > tr:nth-child(2) > td > div > div > table > tbody > tr')

test = rows[7].text.splitlines()
print(test)


#print(rows[1].text)