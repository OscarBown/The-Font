from selenium import webdriver
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta

# open chrome

driver = webdriver.Chrome('./chromedriver')
driver.get('https://clients.mindbodyonline.com/ASP/su1.asp?studioid=425764&tg=&vt=&lvl=&stype=&view=&trn=0&page=&catid=&prodid=&date=8%2f18%2f2021&classid=0&prodGroupId=&sSU=&optForwardingLink=&qParam=&justloggedin=&nLgIn=&pMode=0&loc=1')

time.sleep(4)


SignInButton = driver.find_element_by_xpath('//*[@id="staffSignIn"]')
SignInButton.click()

time.sleep(4)

# sign into mindbody

UserName = "XXXXXXXXXX"
name = driver.find_element_by_xpath('//*[@id="username"]')
name.send_keys(UserName)

Password = "XXXXXXXXX"
name2 = driver.find_element_by_xpath('//*[@id="password"]')
name2.send_keys(Password)

SignInButton = driver.find_element_by_xpath('/html/body/div/span/div/div/div/main/div/div/form/section[2]/button/span')
SignInButton.click()

time.sleep(3)

# get to membership page

FindReportButton = driver.find_element_by_xpath('//*[@id="tabs"]/ul/li[7]/a')
FindReportButton.click()

time.sleep(2)

PricingOptionExpirationButton = driver.find_element_by_xpath('//*[@id="79"]/div[1]/a/div/div/p')
PricingOptionExpirationButton.click()

time.sleep(2)

#membership expiration page instructions

MembershipOptionButton = driver.find_element_by_xpath('//*[@id="main-content"]/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/p/b/select[2]')
MembershipOptionButton.click()

OneMonthMembershipButton = driver.find_element_by_xpath('//*[@id="main-content"]/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/p/b/select[2]/option[5]')
OneMonthMembershipButton.click()

NinetyDayRollingButton = driver.find_element_by_xpath('//*[@id="date-arrows"]/b/ul[3]/li[2]/span')
NinetyDayRollingButton.click()

ClearCalendarButton = driver.find_element_by_xpath('//*[@id="main-content"]/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/p/b/input[2]')
ClearCalendarButton.clear()

DateAfterMonth = datetime.today()+ relativedelta(months=1)
dateformat = DateAfterMonth.strftime("%d/%m/%Y")
TodayDate = driver.find_element_by_xpath('//*[@id="main-content"]/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/p/b/input[2]')
TodayDate.send_keys(dateformat)

# get rid of the pop up box
driver.refresh()

# download the excel file
DownloadExcelButton = driver.find_element_by_xpath('//*[@id="main-content"]/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/b/span[1]')
DownloadExcelButton.click()

time.sleep(1)

# close window
driver.quit()

import pandas as pd

# getting the dates for the source file name
DateAfterMonth1 = datetime.today()+ relativedelta(months=1)
dateformat1 = DateAfterMonth1.strftime("%-m-%-d-%Y")

BackTrackDate = datetime.today()- timedelta(89)
dateformat2 = BackTrackDate.strftime("%-m-%-d-%Y")

# for some reason mindbody excel file is in html format
# so have to bring in data as html and then convert into xlsx
xlsdata = pd.read_html(r"/Users/oscarbown/Downloads/Series Expirations " +dateformat2+ " - " +dateformat1+ ".xls")[0] #use r before absolute file path 
xlsdata.to_excel('/Users/oscarbown/Desktop/The Font/Python automation/Monthly membership- date & No. of visits.xlsx')

# remove old file
os.remove("/Users/oscarbown/Downloads/Series Expirations " +dateformat2+ " - " +dateformat+ ".xls")

# test if file exists
if os.path.exists('/Users/oscarbown/Desktop/The Font/RAW DATA for Tableau/Monthly membership- date & No. of visits.xlsx'):
    print ('file successfully downloaded')
