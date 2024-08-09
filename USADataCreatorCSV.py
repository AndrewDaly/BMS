import csv
from random import randint
from datetime import date
import os
import time
import string
import random
from shutil import copy
import USAStrings
import pysftp
import pandas as pd

# Randomly generate a new global ID
lengthOfGlobalID = 9
globalID = ''.join(["{}".format(randint(0, 9)) for num in range(0, lengthOfGlobalID)])
print("Global ID: " + globalID)

# Create a location to save created files
directory = "USA Global ID " + globalID
path = os.path.join(os.getcwd(), directory)
os.mkdir(path)

# Customize RR email
RRemail = USAStrings.RRFirstName + USAStrings.RRLastName + str(globalID)[5:] + "@bms.com"
print("RR email: " + RRemail)

# Get timestamp for file naming
fdate = date.today().strftime('%m%d%Y')
timestamp = time.strftime("%H%M")
MMDDYYYYHHMM = str(fdate) + timestamp
actorProfiles = "Actor_Profile_Test_" + MMDDYYYYHHMM + ".csv"
REMSRoster = "REMS_Roster_Test_" + MMDDYYYYHHMM + ".csv"
siteMasterData = "SITE_MASTER_TEST_REMS_" + MMDDYYYYHHMM + ".csv"



# Create the actorProfiles csv file
topRow = ['Role', 'BP_ID', 'First_Name', 'Last_Name', 'Job_Title', 'Credentials', 'Primary_Phone', 'Fax', 'Email_Address', 'Global_ID', 'AccountName', 'Product', 'Product_Activation', 'AddressLine1', 'City', 'State', 'Country', 'AffilEntityID', 'AffilEntityName', 'Affil_AccountType', 'AffilEntityType','Affil_To_AE_Mgmt_Hosp']
bottomRow = ['Authorized Representative','', USAStrings.RRFirstName, USAStrings.RRLastName, 'Authorized Rep', USAStrings.RRCredentials, USAStrings.RRPhone, USAStrings.RRFax, RRemail, globalID, USAStrings.accountName, USAStrings.product, '','','','','','','','','','']
actorProfilesDF = pd.DataFrame(columns=topRow)
to_append = bottomRow
a_series = pd.Series(to_append, index = actorProfilesDF.columns)
actorProfilesDF = actorProfilesDF.append(a_series, ignore_index=True)
actorProfilesDF.to_csv(actorProfiles, sep = '|', index=False)
copy(actorProfiles, path)

# Create the REMSRoster csv file
topRow = ['First_Name','Last_Name','Email','Phone_Number','Contact_Type','Global_ID','Product_Activation']
MSLRow = [USAStrings.FMRFirstName, USAStrings.FMRLastName, USAStrings.FMRTitle + str(globalID)[5:] + "@bms.com", USAStrings.FMRPhone, USAStrings.FMRTitle,globalID,USAStrings.product]
CTCSMRow = [USAStrings.KAMFirstName, USAStrings.KAMLastName, USAStrings.KAMTitle + str(globalID)[5:] + "@bms.com", USAStrings.KAMPhone, USAStrings.KAMTitle,globalID,USAStrings.product]
LEMRow = [USAStrings.LEMFirstName, USAStrings.LEMLastName, USAStrings.LEMTitle + str(globalID)[5:] + "@bms.com", USAStrings.LEMPhone, USAStrings.LEMTitle,globalID,USAStrings.product]

# topRow = ['First_Name', 'Last_Name', 'Email', 'Phone_Number', 'Contact_Type', 'Global_ID', 'Product Activation']
# bottomRow = [USAStrings.FMRFirstName, USAStrings.FMRLastName, USAStrings.FMRTitle + str(globalID)[5:] + "@bms.com", USAStrings.FMRPhone, USAStrings.FMRTitle, globalID, USAStrings.product]
REMSRosterDF = pd.DataFrame(columns=topRow)
to_append = MSLRow
a_series = pd.Series(to_append, index = REMSRosterDF.columns)
REMSRosterDF = REMSRosterDF.append(a_series, ignore_index=True)
to_append = CTCSMRow
a_series = pd.Series(to_append, index = REMSRosterDF.columns)
REMSRosterDF = REMSRosterDF.append(a_series, ignore_index=True)
to_append = LEMRow
a_series = pd.Series(to_append, index = REMSRosterDF.columns)
REMSRosterDF = REMSRosterDF.append(a_series, ignore_index=True)

REMSRosterDF.to_csv(REMSRoster, sep = '|', index=False)
copy(REMSRoster, path)

# Create the actorProfile csv file
topRow = ['Global_ID','AccountName','AccountType','Product_Activation','DEA','AddressLine1',                        'City',             'State',               'ZIP',               'Country',          'AffilEntityID','AffilEntityName','Affil_AccountType',          'Affil_DEA','Affil_AddressLine1','Affil_City',         'Affil_State',       'Affil_ZIP','Affil_Country','AffilEntityType','Rems_Site_ID']
bottomRow = [globalID, USAStrings.accountName, 'Hospital', USAStrings.product, 'AA1231232', USAStrings.siteAddress, USAStrings.siteCity, USAStrings.siteState, USAStrings.siteZip, USAStrings.country, globalID,USAStrings.AffilEntityName,USAStrings.Affil_AccountType,'',USAStrings.Affil_AddressLine1,USAStrings.Affil_City,USAStrings.Affil_State,USAStrings.Affil_Zip,USAStrings.Affil_Country,'','']

siteMasterDataDF = pd.DataFrame(columns=topRow)
to_append = bottomRow
a_series = pd.Series(to_append, index = siteMasterDataDF.columns)
siteMasterDataDF = siteMasterDataDF.append(a_series, ignore_index=True)
siteMasterDataDF.to_csv(siteMasterData, sep = '|', index=False)
copy(siteMasterData, path)
print("Files generated")
