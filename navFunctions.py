import numpy as np
import random
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from getpass import getpass
from selenium.webdriver.chrome.options import Options
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from selenium.common.exceptions import NoSuchElementException

##THIS CODE HAS BEEN WRITTEN WITH FILEPATHS MATCHING MAC_OS SYNTAX
##FILEPATHS AND KEYCOMMANDS WILL NEED TO BE TRANSLATED ACCORDINGLY FOR WINDOWS / LINUX OS

username = input("Naviga Username: ")
password = getpass("Naviga Password: ")

commissionsUser = input("NovaTime Username: ")
commissionsPW = getpass("NovaTime Password: ")

options = Options()
options.headless = False
driver = webdriver.Chrome('/Users/The-Intern/Downloads/chromedriver', options=options)


## ZONE IMPORT AUTOMATION ##
def zoneImport(path):
	
	zoneElemXPATHs = {
		"fpath_entry": "/html//input[@id='uImportFilefile0']",
		"test_zoneImport": "/html//input[@id='bTestImport_input']",
		"zoneWindow": "//form[@id='form1']/div[1]/table//td[@class='rwWindowContent']/div/div",
		"get_zoneErrors": "/html//input[@id='wImportErrors_C_tTotalErrors']"
	}
	
	try:
		driver.get("https://ssrtest.navigahub.com/EW/SSR/ad/setup/website_zones_import_by_date")
		time.sleep(3)

		driver.find_element(By.XPATH, zoneElemXPATHs["fpath_entry"]).send_keys(path)
		time.sleep(3)

		driver.find_element(By.XPATH, zoneElemXPATHs["test_zoneImport"]).click()
		time.sleep(3)

		condition = driver.find_element(By.XPATH, zoneElemXPATHs["zoneWindow"]).get_attribute("class")

		good = "rwDialogPopup radalert rwNoIcon"

		errors = driver.find_element(By.XPATH, zoneElemXPATHs["get_zoneErrors"]).get_attribute("value")
	
		print("\n\n----[NAVAD ZONE IMPORT BY DATE RANGE]----")
    
		if (condition == good):
			driver.find_element(By.XPATH, "//form[@id='form1']/div[1]/table//td[@class='rwWindowContent']/div[1]//a[@href='javascript:void(0);']").click()
			driver.find_element(By.XPATH, "/html//input[@id='bTestImport_input']").click()
			print ("IMPORT SUCCESSFUL!")
			time.sleep(3)
      
		else:
			print("ERROR - Failed to import due to records being <100% valid\n" + "Errored Records: " + errors)
			driver.find_element(By.XPATH, "//form[@id='form1']/div[1]/table//td[@class='rwWindowContent']/div[1]//a[@href='javascript:void(0);']").click()
			time.sleep(5)
		
	except NoSuchElementException as exception:
		logOn()
		zoneImport(path)


## AUTHENTICATION ##
def logOn():
	
	actualsElemNames = {
		"UN": "tUserName",
		"PW": "tPassword",
	}
	
	driver.get("https://ssrtest.navigahub.com/EW/SSR/ad/missing_actuals?m=1")
	
	##Authenticate##
	elem = driver.find_element(By.NAME, actualsElemNames["UN"])
	elem.send_keys(username)
	
	elem = driver.find_element(By.NAME, actualsElemNames["PW"])
	elem.send_keys(password)
	time.sleep(1)
	
	elem.send_keys(Keys.RETURN)
	time.sleep(3)
	
	try:
		driver.find_element(By.ID, "dWebGroup_Input")
	except NoSuchElementException as exception:
		logOn()
	
	
##SET QUERY PARAMETERS FOR MISSING ACTUALS TABLE##
def correctMActuals():
	
	try:
		print("\n\n----[MISSING ACTUALS REVISIONS]----")
	
		actualsElemIDs = {
			"product": "dWebGroup_Input",
			"startDate":"dtEndingFrom_dateInput",
			"endDate": "dtUpToDate_dateInput",
			"commitParams": "bGetData_input",
			"get_estimatedQty": "gMonthlyDetail_ctl00_ctl04_nEstimatedQty"
		}
	
		actualsElemXPATHs = {
			"rowCounter_Base": "//div[@id='gCampaigns']/table[@class='rgMasterTable']/tbody",
			"indvRecord_Base": "//div[@id='gCampaigns']/table[@class='rgMasterTable']/tbody/tr[",
			"selectIFrame": "//iframe[@name='wLineDetail']",
			"get_lineData": "//div[@id='gMonthlyDetail_GridHeader']/table/thead/tr/th[4]",
			"lineItemLink": "//div[@id='setupMenu_RadTreeView1']/ul//ul[@class='rtUL']/li[2]/div",
			"send_actualQty": "//div[@id='gMonthlyDetail_GridData']/table//tr[@class='rgRow']//input[@name='gMonthlyDetail$ctl00$ctl04$nActualQty']",
			"get_estimatedAmt": "//div[@id='gMonthlyDetail_GridData']/table//tr[@class='rgRow']//input[@name='gMonthlyDetail$ctl00$ctl04$nEstimatedAmount']",
			"get_actualAmt": "//div[@id='gMonthlyDetail_GridData']/table//tr[@class='rgRow']//input[@name='gMonthlyDetail$ctl00$ctl04$nActualAmount']",
			"saveLineInput": "/html//input[@id='bSaveLine_input']"
		}
	
		inputs = {
			"productValue": "ALL PRODUCTS",
			"startDateValue": "12/12/2022", 
			"endDateValue": "12/14/2022" 
		}
    
	
		driver.get("https://ssrtest.navigahub.com/EW/SSR/ad/missing_actuals?m=1",)
    
		##Clear product params and enter date/product fields##
		productSelector = driver.find_element(By.ID, actualsElemIDs["product"])
	
		productSelector.send_keys(Keys.COMMAND + "a")
		productSelector.send_keys(Keys.BACK_SPACE)
		productSelector.send_keys(inputs["productValue"])
	
		driver.find_element(By.ID, actualsElemIDs["startDate"]).send_keys(inputs["startDateValue"]) #Start Date
		driver.find_element(By.ID, actualsElemIDs["endDate"]).send_keys(inputs["endDateValue"]) #End Date

		##COMMIT QUERY##
		driver.find_element(By.ID, actualsElemIDs["commitParams"]).click()
		time.sleep(3) #KEEP
	
	
		##RECORD ITERATION## --May need to break into sub-functions within loop...could also just move loop to global and move inner block to new function

		countRows = driver.find_element(By.XPATH, actualsElemXPATHs["rowCounter_Base"]).find_elements(By.CSS_SELECTOR, 'tr')
		countRows = len(countRows)
		records = range(1, (int(countRows) + 1))
	
		padPlus = 100
	
		print("Total Records: " + str(countRows - 1))
	
	
		## BEGIN ITERATION THROUGH PRIMARY LIST ##
	
		for i in records:
		
			if countRows < 1:
				print("NO MISSING ACTUALS RETURNED :: EMPTY LIST")
				break
		
			driver.execute_script("window.scrollTo(0, " + str(padPlus) + ")")
			padPlus += 30
			row = driver.find_element(By.XPATH, actualsElemXPATHs["indvRecord_Base"] + str(i) + "]")
		
			if str(row.text) == "-- No Records --":
				print("-- No Records In Missing Actuals --")
				break
		
			recordLink = row.find_element(By.TAG_NAME, "a")
			campID = recordLink.get_attribute("text")
			lineID = driver.find_element(By.XPATH, actualsElemXPATHs["indvRecord_Base"] + str(i) + "]/td[2]").text
			lineID = lineID[0:(len(lineID)-2)]
		
			#print(lineID)
		
			time.sleep(2)
			recordLink.click()
			driver.switch_to.window(driver.window_handles[-1])
		
			clickPrep = driver.find_element(By.XPATH, actualsElemXPATHs["lineItemLink"])
			clickPrep.find_element(By.TAG_NAME, "a").click()
		
			time.sleep(2)
		
			driver.execute_script("window.scrollTo(0, (document.body.scrollHeight)-300)")
		
			time.sleep(2)
		
		
			driver.find_element(By.LINK_TEXT, lineID).click()
		
			iframe = driver.find_element(By.XPATH, actualsElemXPATHs["selectIFrame"])
			driver.switch_to.frame(iframe)
		
			element = str(driver.find_element(By.XPATH, actualsElemXPATHs["get_lineData"]).text)
			estimateQty = driver.find_element(By.ID, actualsElemIDs["get_estimatedQty"])
			actualQty = driver.find_element(By.XPATH, actualsElemXPATHs["send_actualQty"])
		
			actualQty.clear() #may not be necessary
		
			actualQty.send_keys(estimateQty.get_attribute("value"), Keys.CLEAR, Keys.RETURN, Keys.TAB)
			estimatedAmt = driver.find_element(By.XPATH, actualsElemXPATHs["get_estimatedAmt"])
			estimatedAmt = estimatedAmt.get_attribute("value")
			finalAmt = driver.find_element(By.XPATH, actualsElemXPATHs["get_actualAmt"])
			finalAmt = finalAmt.get_attribute("value")
		
			##Test to ensure actual and estimated amounts match##
			if estimatedAmt == finalAmt:
				#time.sleep(2)
				driver.find_element(By.XPATH, actualsElemXPATHs["saveLineInput"]).click()
				print(str(estimateQty.get_attribute("value")) + " quantity :: Submitted")
		
			###### WRAP UP ######
			time.sleep(2) #KEEP
			driver.switch_to.default_content()
			driver.execute_script('window.close()')
			driver.switch_to.window(driver.window_handles[0])
			time.sleep(2) #kEEP
			
	except NoSuchElementException as exception:
		print("Process failed :: Retrying...")
		logOn()
		correctMActuals()


def searchRateCard():
	print("\n\n-------------")
	
	settings = {
		"rcProduct": "Spokesman Review",
		"rcSelection": "All Ratecards",
		"rcID": "118"
	}

	print("[Rate Card Search Results]")
	driver.get("https://ssrtest.navigahub.com/EW/SSR/ad/setup/website_ratecards")
	prod = driver.find_element(By.XPATH, "/html//input[@id='dProduct_Input']")
	prod.send_keys(settings["rcProduct"])
	prod.send_keys(Keys.RETURN)
	time.sleep(4)
	section = driver.find_element(By.XPATH, "/html//input[@id='dSelectionType_Input']")
	section.send_keys(Keys.COMMAND + "a")
	section.send_keys(Keys.BACK_SPACE)
	time.sleep(2)
	section.send_keys(settings["rcSelection"])
	time.sleep(2)
	section.send_keys(Keys.RETURN)
	rateline = driver.find_element(By.XPATH, "/html//input[@id='dRatecardID_Input']")
	rateline.send_keys(settings["rcID"])
	rateline.send_keys(Keys.RETURN)
	
	
def clientAccess(customerID, clientAccessSettings):
	try:
		print("\n\n----[ADMIN SETTINGS: CLIENT ACCESS SETUP]----")
		print("Customer: " + str(customerID))
	
		clientElemXPATHs = {
			"accountIDInput":"/html//input[@id='dAccountID_Input']",
			"drawsInput":"/html//input[@id='chkDraws_input']",
			"returnsInput":"/html//input[@id='chkReturns_input']",
			"invoicesInput":"/html//input[@id='chkInvoices_input']",
			"shortShipmentsInput":"/html//input[@id='chkShortShipments_input']",
			"miscInput":"/html//input[@id='chkMisc_input']",
			"minInvoiceInput":"/html//input[@id='nMinInvoiceAmount']",
			"maxInvoiceInput":"/html//input[@id='nMaxInvoiceAmount']",
			"minShortShipInput":"/html//input[@id='nMinShortShipAmount']",
			"maxShortShipInput":"/html//input[@id='nMaxShortShipAmount']",
			"minMiscInput":"/html//input[@id='nMinMiscAmount']",
			"maxMiscInput":"/html//input[@id='nMaxMiscAmount']",
			"usernameInput":"/html//input[@id='tNewEmail']",
			"passwordInput":"/html//input[@id='tNewPassword']",
			"nameInput":"/html//input[@id='tNewFullName']",
			"submitBtn":"/html//input[@id='bSave_input']"
		}
	
	
		time.sleep(3)
		driver.get("https://ssrtest.navigahub.com/EW/SSR/general/setup/client_access")
		time.sleep(3)
		elem = driver.find_element(By.XPATH, clientElemXPATHs["accountIDInput"])
		elem.send_keys(str(customerID))
		elem.send_keys(Keys.RETURN)
		time.sleep(3)
	
		try:
			driver.find_element(By.XPATH, "//div[@id='gClientAccess']/table[@class='rgMasterTable']//tr[@class='rgRow']/td[1]")
			print("Client Access Already Present")
		except NoSuchElementException as exception:
	
			if clientAccessSettings["allowDraws"] == True:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["drawsInput"])
				elem.send_keys(Keys.SPACE)
		
			if clientAccessSettings["allowReturns"] == True:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["returnsInput"])
				elem.send_keys(Keys.SPACE)
		
			if clientAccessSettings["allowInvoices"] == True:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["invoicesInput"])
				elem.send_keys(Keys.SPACE)
		
			if clientAccessSettings["allowShortShipment"] == True:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["shortShipmentsInput"])
				elem.send_keys(Keys.SPACE)
	
			driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
	
			if clientAccessSettings["allowMiscTransactions"] == True:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["miscInput"])
				elem.send_keys(Keys.SPACE)
		
			if clientAccessSettings["minInvoiceAmount"] > 0:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["minInvoiceInput"])
				elem.send_keys(str(clientAccessSettings["minInvoiceAmount"]))
		
			if clientAccessSettings["maxInvoiceAmount"] > 0:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["maxInvoiceInput"])
				elem.send_keys(str(clientAccessSettings["maxInvoiceAmount"]))
		
			if clientAccessSettings["min_ShortShipmentAmount"] > 0:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["minShortShipInput"])
				elem.send_keys(str(clientAccessSettings["min_ShortShipmentAmount"]))
		
			if clientAccessSettings["max_ShortShipmentAmount"] > 0:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["maxShortShipInput"])
				elem.send_keys(str(clientAccessSettings["max_ShortShipmentAmount"]))
	
			if clientAccessSettings["minMiscAmount"] > 0:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["minMiscInput"])
				elem.send_keys(str(clientAccessSettings["minMiscAmount"]))
		
			if clientAccessSettings["maxMiscAmount"] > 0:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["maxMiscInput"])
				elem.send_keys(str(clientAccessSettings["maxMiscAmount"]))
	
			if True == True:
				elem = driver.find_element(By.XPATH, clientElemXPATHs["usernameInput"])
				elem.send_keys(clientAccessSettings["clientUN"])
		
				elem = driver.find_element(By.XPATH, clientElemXPATHs["passwordInput"])
				elem.send_keys(clientAccessSettings["clientPW"])
		
				elem = driver.find_element(By.XPATH, clientElemXPATHs["nameInput"])
				elem.send_keys(clientAccessSettings["clientName"])
		
				elem = driver.find_element(By.XPATH, clientElemXPATHs["submitBtn"])
				#elem.click()
				print("Client username: " + str(clientAccessSettings["clientUN"]) + "\nClient password: " + str(clientAccessSettings["clientPW"]))
			else:
				print("CLIENT ACCESS WAS NOT CREATED FOR: " + str(customerID))
				
	except NoSuchElementException as exception:
		logOn()
		clientAccess(customerID)


def changeTerritory(accountID, desiredTerr):
	try:
	
		print("\n\n----[REASSIGNED ACCOUNT TERRITORY]----")
	
		driver.get("https://ssrtest.navigahub.com/EW/SSR/general/setup/name_maint_general")
	
		elem = driver.find_element(By.LINK_TEXT, "A/R Setup")
		elem.click()
		elem = driver.find_element(By.XPATH, "/html//input[@id='dAccountID_Input']")
		elem.send_keys(accountID)
		elem.send_keys(Keys.RETURN)
		time.sleep(3)
		elem = driver.find_element(By.XPATH, "/html//input[@id='dTerritory_Input']")
		currentTerr = elem.get_attribute("value")
	
		if currentTerr != desiredTerr:
			print("was: " + currentTerr)
			elem.send_keys(Keys.COMMAND + "a")
			elem.send_keys(Keys.BACK_SPACE)
			elem.send_keys(desiredTerr)
			print("desired: " + desiredTerr)
			elem = driver.find_element(By.XPATH, "/html//input[@id='bSave_input']")
			#elem.click()
			time.sleep(3)
		
		else:
			print("Already desired territory")
		
		driver.execute_script("window.scrollTo(0, 0)")
		elem = driver.find_element(By.LINK_TEXT, "Advertising Setup")
		elem.click()
		elem = driver.find_element(By.XPATH, "/html//input[@id='tTerritory']")
	
		terrCheck = elem.get_attribute("value")
	
		if terrCheck == desiredTerr:
			print("TERRITORY HAS BEEN CHANGED SUCCESSFULLY TO: " + str(terrCheck))
		else:
			print("TERRITORY CHANGE FAILURE ON ADVERTISER: " + str(accountID))
			print("Active Territory: " + str(terrCheck))
			
	except NoSuchElementException as exception:
		logOn()
		changeTerritory(accountID, desiredTerr)
		
def changeTerritory(arr):

	accountID = arr[0]
	desiredTerr = arr[1]
	
	print("\n\n----[REASSIGNED ACCOUNT TERRITORY]----")
	
	driver.get("https://ssrtest.navigahub.com/EW/SSR/general/setup/name_maint_general")
	
	elem = driver.find_element(By.LINK_TEXT, "A/R Setup")
	elem.click()
	elem = driver.find_element(By.XPATH, "/html//input[@id='dAccountID_Input']")
	elem.send_keys(accountID)
	elem.send_keys(Keys.RETURN)
	time.sleep(3)
	elem = driver.find_element(By.XPATH, "/html//input[@id='dTerritory_Input']")
	currentTerr = elem.get_attribute("value")
	
	if currentTerr != desiredTerr:
		print("was: " + currentTerr)
		elem.send_keys(Keys.COMMAND + "a")
		elem.send_keys(Keys.BACK_SPACE)
		elem.send_keys(desiredTerr)
		print("desired: " + desiredTerr)
		elem = driver.find_element(By.XPATH, "/html//input[@id='bSave_input']")
		#elem.click()
		time.sleep(3)
		
	else:
		print("Already desired territory")
		
	driver.execute_script("window.scrollTo(0, 0)")
	elem = driver.find_element(By.LINK_TEXT, "Advertising Setup")
	elem.click()
	elem = driver.find_element(By.XPATH, "/html//input[@id='tTerritory']")
	
	terrCheck = elem.get_attribute("value")
	
	if terrCheck == desiredTerr:
		print("TERRITORY HAS BEEN CHANGED SUCCESSFULLY TO: " + str(terrCheck))
	else:
		print("TERRITORY CHANGE FAILURE ON ADVERTISER: " + str(accountID))
		print("Active Territory: " + str(terrCheck))
		
def brandRepAssign(advertiserID, desiredRep):
	print("\n\n----[REASSIGNED REPS ON BRAND]----")
	print("Advertiser: " + str(advertiserID))
	time.sleep(2)
	driver.get("https://ssrtest.navigahub.com/EW/SSR/ad/setup/brand_digital_reps?a_module=AD")
	elem = driver.find_element(By.XPATH, "/html//input[@id='dAccountID_Input']")
	elem.send_keys(advertiserID)
	elem.send_keys(Keys.RETURN)
	time.sleep(2)
	brandID = driver.find_element(By.XPATH, "/html//input[@id='dBrand_Input']")
	brandID.send_keys("XX")
	brandID.send_keys(Keys.RETURN)
	time.sleep(3)
	getRep = str(driver.find_element(By.XPATH, "//div[@id='gDefaultReps']/table[@class='rgMasterTable']//tr[@class='rgRow']/td[1]").text)
	#time.sleep(3)
	
	if getRep != desiredRep:
		print("Rep on brand was: " + str(getRep))
		print("Target rep on brand is: " + str(desiredRep))
		elem = driver.find_element(By.ID, "gDefaultReps_ctl00_ctl04_gbcEditColumn")
		elem.click()
		time.sleep(2)
		elem = driver.find_element(By.ID, "wAssignments_C_dRep1_Input")
		elem.send_keys(desiredRep)
		elem.send_keys(Keys.RETURN)
		elem = driver.find_element(By.XPATH, "/html//span[@id='wAssignments_C_bSaveLine']")
		elem.click()
		elem = driver.find_element(By.XPATH, "/html//input[@id='wAssignments_C_bSaveLine_input']")
		#elem.click()
	else:
		print("Current rep on brand matches target rep: " + str(getRep))


	
def commLogin():
	
	try:
		driver.get("https://online4.timeanywhere.com/novatime/wslogin.aspx?cid=8EC4C7BD-E576-47B5-87E0-325561A79924")
	
		elem = driver.find_element(By.XPATH, "/html//input[@id='txtUserName']")
		elem.send_keys(commissionsUser)
	
		elem = driver.find_element(By.XPATH, "/html//input[@id='txtPassword']")
		elem.send_keys(commissionsPW)
	
		elem = driver.find_element(By.XPATH, "/html//input[@id='btnLogin']")
		elem.click()
		
	except NoSuchElementException as exception:
		commLogin()
		
		
def postSECommissions():

	try:
		elem = driver.find_element(By.CSS_SELECTOR, ".chevron-tab  tr > td:nth-of-type(2)")
	except NoSuchElementException as exception:
		commLogin()
	
	#try:
	print("\n\n----[SUBMIT FINALIZED COMM PAYROLL]----")
	wb = load_workbook("PyAutoTest.xlsx")
	ws1 = wb.active

	entryDate = "12/24/2022"

	for l in range(2,len(ws1['A'])+1):

		time.sleep(6)

		elem = driver.find_element(By.CSS_SELECTOR, ".chevron-tab  tr > td:nth-of-type(2)")

		elem.find_element(By.TAG_NAME, "a").click()

		time.sleep(4)#4

		iframe = driver.find_element(By.XPATH, "/html//iframe[@id='MainContentFrame']")
		driver.switch_to.frame(iframe)




		LT = ws1['A'+str(l)].value
		AT = ws1['B'+str(l)].value
		#wb.save(filename = "PyAutoTest.xlsx")
		lastName = str(LT).split(",")[0]
		#print(lastName)
		##elem = Select(driver.find_element(By.XPATH, "/html//select[@id='STATUSFILTER']"))
		##elem.select_by_value("-1")

		##time.sleep(3)

		elem = driver.find_element(By.XPATH, "/html//input[@id='txtSearch']")
		elem.send_keys(lastName)

		elem = driver.find_element(By.XPATH, "/html//input[@id='GO_Btn']")
		elem.click()

		time.sleep(3)
		proceed = True

		try:
			driver.find_element(By.LINK_TEXT, LT)
			ws1['C'+str(l)].value = "OPEN"
			wb.save(filename = "PyAutoTest.xlsx")
	
		except NoSuchElementException as exception:
				print("Timecard for " + str(LT) + " is closed")
				ws1['C'+str(l)].value = "CLOSED"
				wb.save(filename = "PyAutoTest.xlsx")
				#email(emailUser, emailPass)
				proceed = False
		
    
		if proceed:
			elem = driver.find_element(By.LINK_TEXT, LT)
			elem.click()
	
			time.sleep(1)

			iframe = driver.find_element(By.ID, "main_frame")
			driver.switch_to.frame(iframe)

      
			for i in range(0,31):
		
				if driver.find_element(By.XPATH, "//table[@id='gdvTS']/tbody/tr[" + str(i+1) + "]/td[3]").find_element(By.TAG_NAME, "span").text == entryDate and driver.find_element(By.XPATH, "//table[@id='gdvTS']/tbody/tr[" + str(i+1) + "]/td[4]").text == "22[SUPPLEMENTAL]":
					break
			
				if len(driver.find_element(By.XPATH, "//table[@id='gdvTS']/tbody/tr[" + str(i+1) + "]/td[3]").find_element(By.TAG_NAME, "span").text) == 0:
          
					#PrepDate
					time.sleep(1)
					elem = driver.find_element(By.XPATH, "//table[@id='gdvTS']/tbody/tr[" + str(i+1) + "]/td[3]")
					elem.click()
					time.sleep(3)
	
					elem = Select(driver.find_element(By.XPATH, "//table[@id='gdvTS']/tbody/tr[" + str(i+1) + "]//select[@name='gdvTS_rw_" + str(i) + "_TPDATE_slc']"))
					elem.select_by_value("12/24/2022")#elem.select_by_index(15)
	

					#PrepType
					elem = driver.find_element(By.XPATH, "//table[@id='gdvTS']/tbody/tr[" + str(i+1) + "]/td[4]")
					elem.click()

					#EnterType
					elem = Select(driver.find_element(By.NAME, "gdvTS_rw_" + str(i) + "_NPAYCODE_slc"))
					elem.select_by_index(1)

					#PrepAmt
					elem = driver.find_element(By.XPATH, "//table[@id='gdvTS']/tbody/tr["+ str(i+1) + "]/td[11]")
					elem.click()

					#EnterType
					elem = driver.find_element(By.NAME, "gdvTS_rw_" + str(i) + "_YPAYAMT_txt")
					elem.send_keys(AT)

					#Save
					elem = driver.find_element(By.XPATH, "/html//input[@id='btn_Apply']")
					#elem.click()
			
					print("Commissions successfully entered for: " + str(LT))
			
					#time.sleep(1)
		
					break
		
		driver.get("https://online4.timeanywhere.com/TimeanywhereExt2/load-timeanywhere1?page=%2Fnovatime%2Fattendance_summary.aspx%3FQPZM%3D131b3a971cfa4e37814691ff41d8ea2d")

    
def closeWindow():
	driver.close()
