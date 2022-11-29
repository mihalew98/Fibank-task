from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import xlsxwriter
import os


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://my.fibank.bg/EBank/public/offices")

time.sleep(10)

offices = driver.find_elements(By.CSS_SELECTOR, value=".sg-office-badge-view")

data = []
for office in offices:
    branch_name = office.find_element(By.CSS_SELECTOR, value="p[bo-bind='item.name'")
    branch_address = office.find_element(By.CSS_SELECTOR, value="p[bo-bind='item.address'")
    branch_workdays = office.find_elements(By.CSS_SELECTOR, value="div.sg-office-work-time dl.dl-horizontal dt")
    branch_workhours = office.find_element(By.CSS_SELECTOR, value="div.sg-office-work-time dl.dl-horizontal dd span")
    branch_phone = office.find_element(By.CSS_SELECTOR, value="p[bo-bind='item.phones[0].phone'")

    branch_workhours_sat = None
    branch_workhours_sun = None
    for workday in branch_workdays:
        if "Събота" in workday.text:
            branch_workhours_sat = branch_workhours.text.split("\n")[1]
        if "Неделя" in workday.text:
            branch_workhours_sun = branch_workhours.text.split("\n")[2]

    if branch_workhours_sat and branch_workhours_sun:
        data.append([branch_name.text, branch_address.text, branch_phone.text, branch_workhours_sat, branch_workhours_sun])
    elif branch_workhours_sat:
        data.append([branch_name.text, branch_address.text, branch_phone.text, branch_workhours_sat, "затворено"])
    elif branch_workhours_sun:
        data.append([branch_name.text, branch_address.text, branch_phone.text, "затворено", branch_workhours_sun])

print(data) #testing

driver.quit()

path = r'C:\PythonApp'
if not os.path.exists(path):
    os.makedirs(path)


workbook = xlsxwriter.Workbook(r"C:\PythonApp\fibank_branches.xlsx")
sheet = workbook.add_worksheet()

sheet.write(0, 0, "Име на офис")
sheet.write(0, 1, "Адрес")
sheet.write(0, 2, "Телефон")
sheet.write(0, 3, "Раб. време събота")
sheet.write(0, 4, "Раб. време неделя")

for i in range(0, len(data)):
    for j in range(0, len(data[i])):
        sheet.write(i+1, j, data[i][j])

workbook.close()