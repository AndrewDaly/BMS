import csv
from random import randint
from datetime import date
import os
import time
import string
import random
from shutil import copy
import ItalyStrings
import pysftp
import pandas as pd

# Randomly generate a new global ID
lengthOfGlobalID = 9
globalID = ''.join(["{}".format(randint(0, 9)) for num in range(0, lengthOfGlobalID)])
print("Global ID: " + globalID)

# Create a location to save created files
directory = "Italy Global ID " + globalID
path = os.path.join(os.getcwd(), directory)
os.mkdir(path)

# Customize RR email
RRemail = ItalyStrings.RRFirstName + ItalyStrings.RRLastName + str(globalID)[5:] + "@bms.com"
print("RR email: " + RRemail)

# Get timestamp for file naming
fdate = date.today().strftime('%m%d%Y')
timestamp = time.strftime("%H%M")
MMDDYYYYHHMM = str(fdate) + timestamp
actorProfiles = "ACTOR_PROFILES_REMS_Italy" + MMDDYYYYHHMM + ".csv"
REMSRoster = "REMS_ROSTER_Italy" + MMDDYYYYHHMM + ".csv"
siteMasterData = "SITE_MASTER_DATA_REMS_Italy" + MMDDYYYYHHMM + ".csv"


# Create the actorProfiles csv file
topRow = ['Role', 'BP_ID', 'First_Name', 'Last_Name', 'Job_Title', 'Credentials', 'Primary_Phone', 'Fax', 'Email_Address', 'Global_ID', 'AccountName', 'Product', 'Product_Activation', 'AddressLine1', 'City', 'State', 'Country', 'AffilEntityID', 'AffilEntityName', 'Affil_AccountType', 'AffilEntityType','Affil_To_AE_Mgmt_Hosp']
bottomRow = ['Authorized Representative','', ItalyStrings.RRFirstName, ItalyStrings.RRLastName, 'Authorized Rep', ItalyStrings.RRCredentials, ItalyStrings.RRPhone, ItalyStrings.RRFax, RRemail, globalID, ItalyStrings.accountName, ItalyStrings.product, '','','','','','','','','','']
actorProfilesDF = pd.DataFrame(columns=topRow)
to_append = bottomRow
a_series = pd.Series(to_append, index = actorProfilesDF.columns)
actorProfilesDF = actorProfilesDF.append(a_series, ignore_index=True)
actorProfilesDF.to_csv(actorProfiles, sep = '|', index=False)
copy(actorProfiles, path)

# Create the REMSRoster csv file
topRow = ['First_Name','Last_Name','Email','Phone_Number','Contact_Type','Global_ID','Product_Activation']
MSLRow = [ItalyStrings.FMRFirstName, ItalyStrings.FMRLastName, ItalyStrings.FMRTitle + str(globalID)[5:] + "@bms.com", ItalyStrings.FMRPhone, ItalyStrings.FMRTitle,globalID,ItalyStrings.product]
CTCSMRow = [ItalyStrings.KAMFirstName, ItalyStrings.KAMLastName, ItalyStrings.KAMTitle + str(globalID)[5:] + "@bms.com", ItalyStrings.KAMPhone, ItalyStrings.KAMTitle,globalID,ItalyStrings.product]
# topRow = ['First_Name', 'Last_Name', 'Email', 'Phone_Number', 'Contact_Type', 'Global_ID', 'Product Activation']
# bottomRow = [ItalyStrings.FMRFirstName, ItalyStrings.FMRLastName, ItalyStrings.FMRTitle + str(globalID)[5:] + "@bms.com", ItalyStrings.FMRPhone, ItalyStrings.FMRTitle, globalID, ItalyStrings.product]
REMSRosterDF = pd.DataFrame(columns=topRow)
to_append = MSLRow
a_series = pd.Series(to_append, index = REMSRosterDF.columns)
REMSRosterDF = REMSRosterDF.append(a_series, ignore_index=True)
to_append = CTCSMRow
a_series = pd.Series(to_append, index = REMSRosterDF.columns)
REMSRosterDF = REMSRosterDF.append(a_series, ignore_index=True)
REMSRosterDF.to_csv(REMSRoster, sep = '|', index=False)
copy(REMSRoster, path)

# Create the actorProfile csv file
topRow = ['Global_ID','AccountName','AccountType','Product_Activation','DEA','AddressLine1','City','State','ZIP','Country','AffilEntityID','AffilEntityName','Affil_AccountType','Affil_DEA','Affil_AddressLine1','Affil_City','Affil_State','Affil_ZIP','Affil_Country','AffilEntityType','Rems_Site_ID']
bottomRow = [globalID,ItalyStrings.accountName, 'Hospital', ItalyStrings.product, 'AA1231232', ItalyStrings.siteAddress, ItalyStrings.siteCity, ItalyStrings.siteProvince, ItalyStrings.siteZip, ItalyStrings.country,'','','','','','','','','','','']

siteMasterDataDF = pd.DataFrame(columns=topRow)
to_append = bottomRow
a_series = pd.Series(to_append, index = siteMasterDataDF.columns)
siteMasterDataDF = siteMasterDataDF.append(a_series, ignore_index=True)
siteMasterDataDF.to_csv(siteMasterData, sep = '|', index=False)
copy(siteMasterData, path)
print("Files generated")

