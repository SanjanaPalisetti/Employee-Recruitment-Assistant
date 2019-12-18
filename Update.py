import os
import xlwt
from xlwt import Workbook
import xlrd
from Functions import * 
from selenium.webdriver.support.ui import Select



# INTIALIZING VARIABLES
username = "anithav@cerium-systems.com"
password = "Cerium@sony"
location = "/Users/sanjanapalisetti/Desktop/Desk/Codes/Recruitment Info.xls"
update_max = 100
dict_candidates_date_posted = {}
dict_candidates_status = {}
dict_candidates_title = {}
dict_candidates_code = {}
dict_candidates_date_submitted = {}
dict_candidates_reason = {}
dict_candidates_experience = {}
dict_candidates_update = {}
dict_jobs_date_posted = {}
dict_jobs_title = {}
dict_jobs_experience = {}
dict_jobs_code = {}
dict_jobs_update = {}




# CHECKING IF EXCEL SHEETS EXIST
if(not os.path.exists(location)):
    result = Excel().Create(location)

result = Excel().ReadingSheets(location)
dict_candidates_date_posted = result[0]
dict_candidates_status = result[1]
dict_candidates_title = result[2]
dict_candidates_code = result[3]
dict_candidates_date_submitted = result[4]
dict_candidates_reason = result[5]
dict_candidates_experience = result[6]
dict_jobs_date_posted = result[7]
dict_jobs_title = result[8]
dict_jobs_experience = result[9]
dict_jobs_code = result[10]
dict_candidates_update = result[11]
dict_jobs_update = result[12]


# GETTING DRIVER AND LOGGING IN
driver = webdriver.Chrome("/chromedriver")
driver.implicitly_wait(30)
driver.get("http://app.talentrackr.com/sony/Vendor/Login.aspx")
login = Login(driver)
login.Username(username)
login.Password(password)
login.Login()


# JOBS TABLE
result = MySubmissionsTable(driver).UpdateJobs(dict_jobs_date_posted,dict_jobs_title,dict_jobs_experience,dict_jobs_code,dict_jobs_update)

dict_jobs_date_posted = result[0]
dict_jobs_title= result[1]
dict_jobs_experience = result[2]
dict_jobs_code = result[3]
dict_jobs_update = result[4]


# WRITING TO EXCEL SHEET
Excel().Write(location,dict_candidates_date_posted,dict_candidates_status,dict_candidates_title,dict_candidates_code,dict_candidates_date_submitted,dict_candidates_reason,dict_candidates_experience,dict_jobs_date_posted,dict_jobs_title,dict_jobs_experience,dict_jobs_code,dict_candidates_update,dict_jobs_update)


# CANDIDATES TABLE 
result = MySubmissionsTable(driver).UpdateCandidates(dict_candidates_date_posted,dict_candidates_status,dict_candidates_title,dict_candidates_code,dict_candidates_date_submitted,dict_candidates_reason,dict_candidates_experience,dict_jobs_date_posted,dict_jobs_title,dict_jobs_experience,dict_jobs_code,dict_candidates_update,update_max)

dict_candidates_date_posted = result[0]
dict_candidates_status = result[1]
dict_candidates_title = result[2]
dict_candidates_code = result[3]
dict_candidates_date_submitted = result[4]
dict_candidates_reason = result[5]
dict_candidates_experience = result[6]
dict_candidates_update = result[7]
            

# WRITING IN EXCEL SHEET
Excel().Write(location,dict_candidates_date_posted,dict_candidates_status,dict_candidates_title,dict_candidates_code,dict_candidates_date_submitted,dict_candidates_reason,dict_candidates_experience,dict_jobs_date_posted,dict_jobs_title,dict_jobs_experience,dict_jobs_code,dict_candidates_update,dict_jobs_update)


# CLOSE DRIVER
driver.close()
