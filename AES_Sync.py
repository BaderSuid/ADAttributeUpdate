from ms_active_directory import ADDomain
from pyad import aduser
from pyad import adcontainer
from pyad import addomain
import pandas as pd
import csv
import smtplib
from email.message import EmailMessage

#Global Variables that will be used throughout the script
domain = ADDomain('###Name OF Domain')
session = domain.create_session_as_user('###SAMaccount name folowed by @domainname', '###Password')
ADPExport = pd.read_csv('###directory location of APD export')
reportfields = ['Name', 'Reports To', 'Job Title', 'Number Of Direct Reports', 'Department', 'Location', 'PositionID']
reportfieldschange = ['Name', 'Reports To', 'Job Title', 'Department', 'Location', 'PositionID']
unresolved = []
unresolvedexport = '###Same Directory as ADP export\\unresolvedexport.csv'
changesdetected = '###Same Directory as ADP export\\changesdetected.csv'
indexnumber = 0
AESaccountsOU = 'OU=AES Accounts,DC=###DomainName,DC=com'

#counts the ammount of rows to define numbeofusers
with open('###directory location of APD export', 'r') as f:
    reader = csv.reader(f)
    numberofusers = sum(1 for rown in reader) - 1

#starts a while loop to iterate through every row in the ADP export. Increments by one at the end of iteration.
while indexnumber < numberofusers:
    change = False

    #creates an array of the attributes within the AES report to be compaired against later. This will be used to detect changes.
    exportattributes = []

#error handling to account for anomalies
    try:
        #finds the employee's positionID and queries Active directory for an object with a matching value for the specified attribute.
        directreportpositionID = ADPExport.loc[indexnumber, 'Position ID']
        directreportemployeeID = str(directreportpositionID).removeprefix('###Prefix 1').removeprefix('###Prefix 2').removeprefix('###Prefix 3').removeprefix('###Prefix 4')
        directreportADarray = session.find_users_by_attribute('employeeNumber', directreportemployeeID, ['cn', 'manager', 'description', 'department', 'location', 'physicalDeliveryOfficeName', 'employeeID', 'distinguishedName'])

        #uses the array that gets returned to define variables that can be used to identify that user.
        for user in directreportADarray:
            directreportDN = user.get('distinguishedName')
            currentmanager = str(user.get('manager'))
            directreportCN = user.get('cn')

        directreportAD = aduser.ADUser.from_dn(directreportDN)

        #checks if the user currently has a manager for later use.
        if pd.isnull(ADPExport.loc[indexnumber, 'Reports To']):
            hasmanager = False
        else:
            #uses the manager's name to find the instance of that person within the same report to get their PositionID. Uses this to find the manager in Active Directory.
            reportto = ADPExport.loc[indexnumber, 'Reports To']
            findmanager = ADPExport.loc[ADPExport['Name'] == reportto]
            managerrow = findmanager.index[0]
            newmanagerpositionID = ADPExport.loc[managerrow, 'Position ID']
            newmanageremployeeID = str(newmanagerpositionID).removeprefix('QDS').removeprefix('XUV').removeprefix('QZX').removeprefix('SKR')
            newmanagerADarray = session.find_users_by_attribute('employeeNumber', newmanageremployeeID, ['distinguishedname', 'cn'])

            for users in newmanagerADarray:
                newmanagerDN = user.get('distinguishedname')
                newmanagerCN = user.get('cn')

            exportattributes.append(str(newmanagerDN))

        #finds the new location, title, and department of the user and appends them to the exportattributes array.
        location = str(ADPExport.loc[indexnumber, 'Location'])

        if pd.isnull(location):
            office = 'NA'
        else:
            office = location

        exportattributes.append(str(location))

        description = str(ADPExport.loc[indexnumber, 'Job Title'])

        if pd.isnull(description):
            description = 'NA'

        exportattributes.append(description)

        department = str(ADPExport.loc[indexnumber, 'Department'])

        if pd.isnull(department):
            department = 'NA'

        exportattributes.append(department)

        #checks if the user doesn't exsist in Active Directory. Creates basic user account with no access if it doesn't.
        if directreportADarray == []:
            if directreportpositionID[:3] == 'RYW':
                pass
            else:
                if pd.isnull(ADPExport.loc[indexnumber, 'Preferred Or Chosen Last Name']):
                    directreportlastname = ADPExport.loc[indexnumber, 'Last Name']
                
            if pd.isnull(ADPExport.loc[indexnumber, 'Preferred Or Chosen First Name']):
                directreportfirstname = ADPExport.loc[indexnumber, 'First Name']
            else:
                directreportfirstname = ADPExport.loc[indexnumber, 'Preferred Or Chosen First Name']

            username = f'{directreportfirstname} {directreportlastname}'
            newuserDN = str(f'CN={username},OU=IT,DC=specialdevices,DC=com')
            usercreation = aduser.ADUser.create(name=username, container_object=adcontainer.ADContainer.from_dn(AESaccountsOU), password='5@PeFf$WV4iQZ8')
            usercreation.update_attribute('employeeNumber', directreportemployeeID)
            #starts the loop from the top on the same iteration to update the attributes of the newly created user.
            continue
        else:
            #creates an array for the current attributes of the user and appends value of current attributes
            currentattributes = []

            for user in directreportADarray:
                currentattributes.append(user.get('cn'))
                currentattributes.append(user.get('manager'))
                arraytostrdescription = ''.join(user.get('description'))
                currentattributes.append(arraytostrdescription)
                currentattributes.append(user.get('department'))
                currentattributes.append(user.get('physicalDeliveryOfficeName'))

            #creates a dictionary that will be used to audit changes detected. Keys are blank and will be defined as discrepancies between the current attributes and the export attribute are found.
            updatedattributesdic = {
                'directreport': '',
                'manager': '',
                'description': '',
                'department': '',
                'physicalDeliveryOfficeName': ''
            }

            #iterates through every value of the exportattributes array and checks if they are current. If a discrepancy is detected, that attribute defines the related value's key in the updatedattributes dictionary and is used to udpate the user in Active Directory.
            #if no change is detected the key stays blank to show no change has been made.
            for attribute in exportattributes:
                if attribute not in currentattributes:
                    change = True

                    updatedattributesdic['directreport'] = directreportCN
                    if attribute == newmanagerDN:
                        directreportAD.update_attribute('manager', newmanagerDN)
                        updatedattributesdic['manager'] = newmanagerCN
                    elif attribute == description:
                        directreportAD.update_attribute('description', attribute)
                        directreportAD.update_attribute('title', attribute)
                        updatedattributesdic['description'] = attribute
                    elif attribute == department:
                        directreportAD.update_attribute('department', attribute)
                        updatedattributesdic['department'] = attribute
                    elif attribute == office:
                        directreportAD.update_attribute('physicalDeliveryOfficeName', attribute)
                        updatedattributesdic['physicalDeliveryOfficeName'] = attribute

            #appends the user's old and new attributes to the changsdetected csv. 
            with open(changesdetected, 'a', newline='') as f:
                writerdic = csv.DictWriter(f, fieldnames=updatedattributesdic.keys())
                writer = csv.writer(f)
                if change == True:
                    writer.writerow(currentattributes)
                    writerdic.writerow(updatedattributesdic)
                    writer.writerow('')
    
    #if an error occurs that would halt the excution of this script, it skips that user's iteration and includes all the details of that row in the unresolved array. 
    except:
        notinad = ADPExport.loc[indexnumber, :]
        unresolved.append(notinad)

    #jumps to next user's row number
    indexnumber+=1

#uses the unresolved array to create the unresolved export that will be sent out as an email.
with open(unresolvedexport, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(reportfields)
    writer.writerows(unresolved)

#starts smtp connection with internal mail relay. Outlook or gmail can be used if no mail server exists.
server = smtplib.SMTP('###Internal Mail Server Relay', 25)

#creates the post execution report
msg = EmailMessage()
msg['From'] = 'AESAutomation'
msg['To'] = ['###Desired Recipiants']
msg['Subject'] = 'AES Automation'
msg.set_content('Execution of AES Automation successful. Reports of changes made and instances of failure included')

with open(unresolvedexport, 'rb') as f:
    unresolvedexportmail = f.read()

with open(changesdetected, 'rb') as f:
    changesdetectedmail = f.read()

#adds the changesdetected and unresolvedexport csvs as attachments to the email. 
msg.add_attachment(unresolvedexportmail, maintype='application', subtype='octet-stream', filename='UnresolvedExport.csv')
msg.add_attachment(changesdetectedmail, maintype='application', subtype='octet-stream', filename='ChangesDetected.csv')
server.send_message(msg)
server.close()

#erases the changesdetected csv for the next execution.
with open(changesdetected, 'w', newline='') as f:
    f.write('')
    writer = csv.writer(f)
    writer.writerow(reportfieldschange)