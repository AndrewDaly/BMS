import csv
from random import randint
from datetime import date
import os
import time
import string
import random
from shutil import copy
import GermanyStrings
import pysftp
import pandas as pd

# Randomly generate a new global ID
lengthOfGlobalID = 9
globalID = ''.join(["{}".format(randint(0, 9)) for num in range(0, lengthOfGlobalID)])
print("Global ID: " + globalID)

# Create a location to save created files
directory = "Germany Global ID " + globalID
path = os.path.join(os.getcwd(), directory)
os.mkdir(path)

# Customize RR email
RRemail =GermanyStrings.RRFirstName +GermanyStrings.RRLastName + str(globalID)[5:] + "@bms.com"
print("RR email: " + RRemail)

# Get timestamp for file naming
fdate = date.today().strftime('%m%d%Y')
timestamp = time.strftime("%H%M")
MMDDYYYYHHMM = str(fdate) + timestamp
actorProfiles = "ACTOR_PROFILES_REMS_Germany" + MMDDYYYYHHMM + ".csv"
REMSRoster = "REMS_ROSTER_Germany" + MMDDYYYYHHMM + ".csv"
siteMasterData = "SITE_MASTER_DATA_REMS_Germany" + MMDDYYYYHHMM + ".csv"


# Create the actorProfiles csv file
topRow = ['Role', 'BP_ID', 'First_Name', 'Last_Name', 'Job_Title', 'Credentials', 'Primary_Phone', 'Fax', 'Email_Address', 'Global_ID', 'AccountName', 'Product', 'Product_Activation', 'AddressLine1', 'City', 'State', 'Country', 'AffilEntityID', 'AffilEntityName', 'Affil_AccountType', 'AffilEntityType','Affil_To_AE_Mgmt_Hosp']
bottomRow = ['Authorized Representative','',GermanyStrings.RRFirstName,GermanyStrings.RRLastName, 'Authorized Rep',GermanyStrings.RRCredentials,GermanyStrings.RRPhone,GermanyStrings.RRFax, RRemail, globalID,GermanyStrings.accountName,GermanyStrings.product, '','','','','','','','','','']
actorProfilesDF = pd.DataFrame(columns=topRow)
to_append = bottomRow
a_series = pd.Series(to_append, index = actorProfilesDF.columns)
actorProfilesDF = actorProfilesDF.append(a_series, ignore_index=True)
actorProfilesDF.to_csv(actorProfiles, sep = '|', index=False)
copy(actorProfiles, path)

# Create the REMSRoster csv file
topRow = ['First_Name','Last_Name','Email','Phone_Number','Contact_Type','Global_ID','Product_Activation']
MSLRow = [GermanyStrings.FMRFirstName,GermanyStrings.FMRLastName,GermanyStrings.FMRTitle + str(globalID)[5:] + "@bms.com",GermanyStrings.FMRPhone,GermanyStrings.FMRTitle,globalID,GermanyStrings.product]
CTCSMRow = [GermanyStrings.KAMFirstName,GermanyStrings.KAMLastName,GermanyStrings.KAMTitle + str(globalID)[5:] + "@bms.com",GermanyStrings.KAMPhone,GermanyStrings.KAMTitle,globalID,GermanyStrings.product]
# topRow = ['First_Name', 'Last_Name', 'Email', 'Phone_Number', 'Contact_Type', 'Global_ID', 'Product Activation']
# bottomRow = [GermanyStrings.FMRFirstName,GermanyStrings.FMRLastName,GermanyStrings.FMRTitle + str(globalID)[5:] + "@bms.com",GermanyStrings.FMRPhone,GermanyStrings.FMRTitle, globalID,GermanyStrings.product]
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
bottomRow = [globalID,GermanyStrings.accountName, 'Hospital',GermanyStrings.product, 'AA1231232',GermanyStrings.siteAddress,GermanyStrings.siteCity,GermanyStrings.siteProvince,GermanyStrings.siteZip,GermanyStrings.country,'','','','','','','','','','','']

siteMasterDataDF = pd.DataFrame(columns=topRow)
to_append = bottomRow
a_series = pd.Series(to_append, index = siteMasterDataDF.columns)
siteMasterDataDF = siteMasterDataDF.append(a_series, ignore_index=True)
siteMasterDataDF.to_csv(siteMasterData, sep = '|', index=False)
copy(siteMasterData, path)
print("Files generated")

