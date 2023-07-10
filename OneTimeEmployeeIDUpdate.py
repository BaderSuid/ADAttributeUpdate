from ms_active_directory import ADDomain
import csv
import smtplib
from email.message import EmailMessage
import pandas as pd

idexport = pd.read_csv("***Directory Of Report")
indexnumber = 0
numberofusers = len(idexport.loc[:,'Position ID'].values)
domain = ADDomain('***Domain Name')
session = domain.create_session_as_user('***SAMaccountname@DomainName', '***Password')
unresolved = []
unresolvedexport = "***Empty Csv called unresolvedexport"
toppart = ['Position ID', 'Legal Last Name', 'Legal First Name']

while indexnumber < numberofusers:
    firstname = idexport.loc[indexnumber, 'Legal First Name']
    lastname = idexport.loc[indexnumber, 'Legal Last Name']
    username = f'{firstname} {lastname}'
    positionID = str(idexport.loc[indexnumber, 'Position ID'])

    employeeID = str(positionID).removeprefix('QDS').removeprefix('XUV').removeprefix('QZX').removeprefix('SKR')

    try:
        location = positionID[:3]
        #did this to skip users in a specific location that isn't apart of our domain. Soon to close down. 
        if location == 'RYW':
            print('BDK User Skipped')
        else:
            userad = session.find_user_by_name(username)
            session.overwrite_attribute_for_user(userad, 'employeeNumber', employeeID)
            session.overwrite_attribute_for_user(userad, 'employeeID', positionID)

    except:
        unresolved.append(idexport.loc[indexnumber, :])

    indexnumber+=1


with open(unresolvedexport, 'w', newline='') as csvfile:
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(toppart)
    csvwriter.writerows(unresolved)


#Creates email with csv of all the people that were not found in AD
server = smtplib.SMTP('***Local Mail Relay', 25)

msg = EmailMessage()
msg['From'] = 'ADAttributeAutomation@***relayname.com'
msg['To'] = ['Desired Recipiants']
msg['Subject'] = 'AD Attribute Edit Unresolved Report'
msg.set_content('The following users did not have their attributes changed automatically. ')

with open(unresolvedexport, 'rb') as f:
    export = f.read()

msg.add_attachment(export, maintype='application', subtype='octet-stream', filename='UnresolvedExport.csv')
server.send_message(msg)
server.close