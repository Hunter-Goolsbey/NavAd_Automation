from NavAdScriptMaster import *

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


clientAccessSettings = {
			"allowDraws": True,
			"allowReturns": True,
			"allowInvoices": False,
			"allowShortShipment": True,
			"allowMiscTransactions": True,
			"minInvoiceAmount": 0,
			"maxInvoiceAmount": 0,
			"min_ShortShipmentAmount": 0,
			"max_ShortShipmentAmount": 0,
			"minMiscAmount": 0,
			"maxMiscAmount": 0,
			"clientUN": "test@email.com", #typically email... cannot think of another reason to use other data
			"clientPW": "FlatTire" + str(random.randint(1,17)),
			"clientName": "John Smith"
		}
		
terrChangeArr = [["003939","Bob Smith"]]

#inputs for zone import(1): 1 filepath (directs to wherever weekly zone imports are stored)
#--date suffix can be used to automate further as long as .xlsx doc name is prepared with same date as script run-date
ziDate = datetime.now().strftime("%Y%m%d")
zoneImport("/Users//Downloads/zones_import_by_date" + str(ziDate) + ".xlsx")

#inputs for correct M actuals (2): 1 start date, 1 end date
correctMActuals()

#inputs for client access(2): 1 accountID, 1 dictionary **See example object at top**
clientAccess("03022", clientAccessSettings)

for i in terrChangeArr:
	#inputs for change territory(1): 1 (2-axis array containing accountID and desired rep on account)
	changeTerritory(i)
  
#inputs for zone import: 1 filepath as string, directing to reference wkbk.  
#--wkbk is wherever you have stored commission and rep-name data for the given month (masterComm)
postSECommissions()

time.sleep(3)

closeWindow()
