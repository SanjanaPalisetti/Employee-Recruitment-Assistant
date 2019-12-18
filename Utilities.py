from selenium import webdriver
import unittest, time, re
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.by import By
import xlwt 
from xlwt import Workbook 
import os


class Login(object):

    def __init__(self,driver):
        self.driver = driver

    def Username(self,username):
        self.driver.find_element_by_id("ctl00_cphHomeMaster_LoginUser_UserName").clear()
        self.driver.find_element_by_id("ctl00_cphHomeMaster_LoginUser_UserName").send_keys(username)

    def Password(self,password):
        self.driver.find_element_by_id("ctl00_cphHomeMaster_LoginUser_Password").clear()
        self.driver.find_element_by_id("ctl00_cphHomeMaster_LoginUser_Password").send_keys(password)
        
    def Login(self):
        self.driver.find_element_by_id("ctl00_cphHomeMaster_LoginUser_LoginButton").click()

class NavigateToPage(object):

    def __init__(self,driver):
        self.driver = driver

    def MyRequisitions(self):
        self.driver.find_element_by_id("ctl00_ctlApplicationMenu_rptTopMenu_ctl00_lblMenuName").click()
        self.driver.find_element_by_link_text("My Submissions").click()

class MySubmissionsTable(object):

    def __init__(self,driver):

        self.driver = driver

    def Reason(self,title):

        first_part1 = "//*[@id='ctl00_ctl00_cphHomeMaster_cphCandidateMaster_uclCandidateOverView_uclRejection_lsvInterviewRejection_ctrl"
        second_part1 = "_row']/td[2]"
        second_part2 = "_row']/td[5]"
        i=0
        while(1):
            if(self.driver.find_element_by_xpath(first_part1+str(i)+second_part1).text in title):
                return first_part1+str(i)+second_part2
            i=i+1

    def Rejection(self,xpath,table_data,my_sub_reject,page,title):

        if(("reject" in table_data.lower()) or ("rejected" in table_data.lower())):
            wait = ui.WebDriverWait(self.driver,30)
            wait.until(lambda driver: driver.find_element_by_xpath(xpath))
            self.driver.find_element_by_xpath(xpath).click()
            html_source = self.driver.page_source
            
            for i in range(0,4): # ERROR HANDLING
                if("Reason for Rejection" in html_source):
                    break
                self.driver.find_element_by_xpath(xpath).click()
            try:
                self.driver.find_element_by_link_text("Reason for Rejection").click()
            except:
                my_sub_reject.append("Error")
                return my_sub_reject


            html_source = self.driver.page_source
  
            if "Rejected By" in html_source: # IF REASON IS MENTIONED
                try:
                    my_sub_reject.append(self.driver.find_element_by_xpath(MySubmissionsTable(self.driver).Reason(title)).text)
                except:
                    my_sub_reject.append("Reason not mentioned")
            
            else: # IF REASON IS NOT MENTIONED
                my_sub_reject.append("Reason not mentioned")

            self.driver.back()
            if(page!=1):
                self.driver.find_element_by_id("ctl00_cphHomeMaster_uclVendorSubmissions_lsvVendorSubmissions_pagerControl_pager_ctl00_txtSlider").clear()
                self.driver.find_element_by_id("ctl00_cphHomeMaster_uclVendorSubmissions_lsvVendorSubmissions_pagerControl_pager_ctl00_txtSlider").send_keys(str(page))
                self.driver.find_element_by_id("aspnetForm").submit()
                print(f"Staying in page {page}")
                time.sleep(5)
        
        else: # IF ACCEPTED   
            my_sub_reject.append("NA")

        return my_sub_reject


    def NextPage(self,xyz,page,pages,endpage):

        if((xyz+1)%pages==0):
            page+=1
            print("Page value increased")
        '''if((xyz+1)%pages==pages):
            endpage+=1
            self.driver.find_element_by_id("ctl00_cphHomeMaster_uclVendorSubmissions_lsvVendorSubmissions_pagerControl_pager_ctl00_txtSlider").clear()
            self.driver.find_element_by_id("ctl00_cphHomeMaster_uclVendorSubmissions_lsvVendorSubmissions_pagerControl_pager_ctl00_txtSlider").send_keys(str(endpage))
            self.driver.find_element_by_id("aspnetForm").submit()
            time.sleep(3)'''
        if((xyz+1)%pages==0):
            self.driver.find_element_by_xpath('//*[@id="ctl00_cphHomeMaster_uclVendorSubmissions_lsvVendorSubmissions_pagerControl_pager_ctl00_btnNext"]').click()
            print("Clicking button")
            time.sleep(5)
        return page


    def Iteration(self,mysub_rows,pages,page,endpage,first_part,second_part,third_part):

        my_sub_name = []
        my_sub_date_submitted = []
        my_sub_job_title = []
        my_sub_status = []
        my_sub_reject = []
        my_sub_job_id = []
        matrix2 = [[0 for x in range(5)] for y in range(mysub_rows)] 

        for xyz in range(0,mysub_rows):

            for abc in range(1,5):
                        
                if((xyz+1)%pages==0):
                    xyzz=pages-1
                else:
                    xyzz=((xyz+1)%pages)-1

                whole_string = first_part+str(xyzz)+second_part+str(abc)+third_part
                table_data = self.driver.find_element_by_xpath(whole_string).text
                req_id = "ctl00_cphHomeMaster_uclVendorSubmissions_lsvVendorSubmissions_ctrl"+str(xyzz)+"_lnkJobTitle"
                
                if(abc==2): # FOR NAME OF CANDIDATE
                    my_sub_name.append(table_data)
                    print(table_data)
                    xpath = whole_string
                    matrix2[xyz][1]=table_data
                                      
                elif(abc==1):  # FOR DATE OF SUBMISSION
                    my_sub_date_submitted.append(table_data)
                    matrix2[xyz][0]=table_data
              
                elif(abc==3): # FOR JOB TITLE AND FOR JOB ID
                    title = table_data
                    my_sub_job_title.append(table_data)
                    matrix2[xyz][2]=table_data
                    my_sub_job_id.append(OnClick(self.driver).Info(req_id))
                    matrix2[xyz][4]=OnClick(self.driver).Info(req_id)
               
                elif(abc==4): # FOR REJECTION STATUS
                    my_sub_status.append(table_data)
                    matrix2[xyz][3]=table_data

                    # FOR REASON OF REJECTION
                    my_sub_reject = MySubmissionsTable.Rejection(self,xpath,table_data,my_sub_reject,page,title)

                print(f"-----------{xyz+1,abc}-----------")

            # MOVING TO NEXT PAGE
            page = MySubmissionsTable.NextPage(self,xyz,page,pages,endpage)

        result=[my_sub_name,my_sub_date_submitted,my_sub_job_title,my_sub_status,my_sub_reject,my_sub_job_id,matrix2]
        return result


    
class JobOpeningsTable(object):

    def __init__(self,driver):

        self.driver = driver

    
    def Iteration(self,job_openings_rows,first_part,second_part,third_part,pages):

        job_code = {}
        date_posted = {}
        experience = {}
        job_title = {}
        matrix1 = [[0 for x in range(5)] for y in range(job_openings_rows)] 

        for xyz in range(0,job_openings_rows):

            for abc in range(1,5):

                if((xyz+1)%100==0):
                    xyzz=99
                else:
                    xyzz=((xyz+1)%100)-1
                whole_string = first_part+str(xyzz)+second_part+str(abc)+third_part
                table_data = self.driver.find_element_by_xpath(whole_string).text
                req_id = "ctl00_cphHomeMaster_VendorActiveJobOpenings_lsvJobPosting_ctrl"+str(xyzz)+"_lnkJobTitle"
                if(abc==1):
                    var1=table_data                             # DATE POSTED
                elif(abc==2):
                    var2=table_data                             # JOB CODE
                elif(abc==3):
                    var3=table_data                             # JOB NAME
                    var=OnClick(self.driver).Info(req_id)       # JOB ID
                elif(abc==4):
                    var4=table_data                             # EXPERIENCE
                print(f"-----------{xyz,abc}-----------")
            if((xyz+1)%100==0):
                self.driver.find_element_by_xpath("//*[@id='ctl00_cphHomeMaster_uclVendorSubmissions_lsvVendorSubmissions_pagerControl_pager_ctl00_btnNext']").click()
                time.sleep(3)
            date_posted[var2] = var1    # JOB CODE -> DATE POSTED
            job_code[var] = var2        # JOB CODE -> JOB TITLE
            experience[var2] = var4     # JOB CODE -> EXPERIENCE
            job_title[var] = var3       # JOB ID -> JOB TITLE
            matrix1[xyz][0] = var1
            matrix1[xyz][1] = var2
            matrix1[xyz][2] = var3
            matrix1[xyz][3] = var4
            matrix1[xyz][4] = var

        print(job_code,date_posted,experience,matrix1)

        result = [date_posted,job_code,experience,job_title,matrix1]
        return result



class Excel(object):

    def __init__(self,driver):

        self.driver = driver


    def Remove(self):

        if os.path.exists("Recruitment Info.xls"):
            os.remove("Recruitment Info.xls")
        else:
            print("The file does not exist")


    def Write(self,my_sub_name,job_code,my_sub_job_title,date_posted,experience,my_sub_date_submitted,my_sub_status,my_sub_reject,mysub_rows,job_title,my_sub_job_id):

        wb = Workbook()
        sheet1 = wb.add_sheet('sheet1') 
        style = xlwt.easyxf('font: bold 1')
        sheet1.write(0,0,'SERIAL NO',style)
        sheet1.write(0,1,'CANDIDATE NAME',style)
        sheet1.write(0,2,'JOB CODE',style)
        sheet1.write(0,3,'JOB TITLE',style)
        sheet1.write(0,4,'DATE POSTED',style)
        sheet1.write(0,5,'YEARS OF EXPERIENCE',style)
        sheet1.write(0,6,'DATE SUBMITTED',style)
        sheet1.write(0,7,'REJECTION STATUS',style)
        sheet1.write(0,8,'REASON FOR REJECTION',style)
        sheet1.write(0,9,'JOB ID',style)

        for i in range(1,mysub_rows+1):
            print(f"{i}th iteration")
            sheet1.write(i,0,i)
            sheet1.write(i,1,my_sub_name[i-1])
            if(my_sub_job_id[i-1].isdigit()):
                required_job_id = my_sub_job_id[i-1]
            else:
                required_job_id = '2'+my_sub_job_id[i-1][:3]
            sheet1.write(i,2,job_code.get(my_sub_job_id[i-1]))
            #sheet1.write(i,2,required_job_code)
            sheet1.write(i,3,my_sub_job_title[i-1])
            sheet1.write(i,4,date_posted.get(job_code.get(my_sub_job_id[i-1])))
            #sheet1.write(i,4,date_posted.get(required_job_code))
            sheet1.write(i,5,experience.get(job_code.get(my_sub_job_id[i-1])))
            #sheet1.write(i,5,experience.get(required_job_code))
            sheet1.write(i,6,my_sub_date_submitted[i-1])
            sheet1.write(i,7,my_sub_status[i-1])
            sheet1.write(i,8,my_sub_reject[i-1])
            sheet1.write(i,9,required_job_id)
            
        wb.save('Recruitment Info.xls') 

    def Temp1(self,matrix1,matrix2,mysub_rows,job_openings_rows):

        if os.path.exists("Recruitment Tables.xls"):
            #call another function
            return

        wb = Workbook()
        sheet1 = wb.add_sheet('Job Openings')
        sheet2 = wb.add_sheet('My Submissions')
        style = xlwt.easyxf('font: bold 1')
        sheet1.write(0,0,'SERIAL NO',style)
        sheet1.write(0,1,'DATE POSTED',style)
        sheet1.write(0,2,'JOB CODE',style)
        sheet1.write(0,3,'JOB TITLE',style)
        sheet1.write(0,4,'EXPERIENCE',style)
        sheet1.write(0,5,'JOB ID',style)
        sheet2.write(0,0,'SERIAL NO',style)
        sheet2.write(0,1,'DATE SUBMITTED',style)
        sheet2.write(0,2,'CANDIDATE NAME',style)
        sheet2.write(0,3,'JOB TITLE',style)
        sheet2.write(0,4,'STATUS',style)
        sheet2.write(0,5,'JOB ID',style)
        for i in range(1,job_openings_rows):
            sheet1.write(i,0,i)
            sheet1.write(i,1,matrix1[i-1][0])
            sheet1.write(i,2,matrix1[i-1][1])
            sheet1.write(i,3,matrix1[i-1][2])
            sheet1.write(i,4,matrix1[i-1][3])
            sheet1.write(i,5,matrix1[i-1][4])
        for i in range(1,mysub_rows):
            sheet2.write(i,0,i)
            sheet2.write(i,1,matrix2[i-1][0])
            sheet2.write(i,2,matrix2[i-1][1])
            sheet2.write(i,3,matrix2[i-1][2])
            sheet2.write(i,4,matrix2[i-1][3])
            sheet2.write(i,5,matrix2[i-1][4])

        wb.save('Recruitment Tables') 


class OnClick(object):

    def __init__(self,driver):

        self.driver = driver

    
    def Info(self,req_id):

        element = self.driver.find_element_by_id(req_id)
        value = element.get_attribute("onclick")
        value_list = value.split("JID=")
        result = value_list[1][:4]
        print(result)
        return result