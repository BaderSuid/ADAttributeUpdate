from ms_active_directory import ADDomain
from pyad import aduser
from pyad import adcontainer
from pyad import addomain
from pyad import adgroup
import pandas as pd
import csv
import smtplib
from email.message import EmailMessage
import pysftp
import os
from datetime import date
import pysftp as sftp
import shutil

#Credentials for remote SFTP server
FTP_HOST = '######'
FTP_USER = '########'
FTP_PASS = '#########'

cnopts = sftp.CnOpts()
cnopts.hostkeys = None

with sftp.Connection(host=FTP_HOST, port=22, username=FTP_USER, password=FTP_PASS, cnopts=cnopts) as sftp:
    reportlist = sftp.listdir('/fromADP')
    if len(reportlist) > 1:
        sftp.remove(f'/fromADP/{reportlist[0]}')
        quit()
    else:
        sftp.get_r(remotedir='/fromADP/', localdir='###ReportDir')

domain = ADDomain('###Domain Name###')
session = domain.create_session_as_user('###SAMAccountname@Domainname', '###PSSWD')
ADPExport = pd.read_csv('###ReportDir')
reportfields = ['Employee Last Name, First Name', 'Last Name', 'First Name', 'Legal Middle Name', 'Worked in State Description', 'Associate ID', 'Position ID', 'Payroll Company Code', 'Job Title Code', 'Home Department Description', 'Location City', 'Location Description', 'Worker Category Description', 'Position Status', 'Termination Date', 'Reports To Name', 'Preferred Or Chosen First Name', 'Preferred Or Chosen Middle Name', 'Preferred Or Chosen Last Name', 'Job Title Description', 'Reports To Position ID', 'Reports To Associate ID']
reportfieldschange = ['Name', 'Reports To', 'Job Title', 'Location', 'PositionID']
unresolved = []
unresolvedexport = 'C:/Users/######/AES/unresolvedexport.csv'
changesdetected = 'C:/Users/######/AES/changesdetected.csv'
#CHANGE ARCHIVE DIR TO THE USER_ARCHIVE DIR IN ITDATA
archive_dir = '//####fileshare/######/User_Archive/'
AESaccountsOU = adcontainer.ADContainer.from_dn('OU=AES,DC=####domain,DC=com')

def getinfo_update(pipe_useramount):
    nonactive_indexstorage = []
    indexnumber = 0

    while indexnumber < pipe_useramount:
        print(f'{indexnumber} First Function')
        try:
            change = False
            status = str(ADPExport.loc[indexnumber, 'Position Status'])
            directreport_associateID = ADPExport.loc[indexnumber, 'Associate ID']
            directreport_positionID = ADPExport.loc[indexnumber, 'Position ID']
            company_code = ADPExport.loc[indexnumber, 'Payroll Company Code']
            directreportADarray = session.find_users_by_attribute('employeeID', directreport_associateID, ['cn', 'manager', 'description', 'department', 'location', 'physicalDeliveryOfficeName', 'employeeID', 'distinguishedName', 'userAccountControl', 'extensionAttribute1'])
            if status == 'Active':
                if directreportADarray == []:
                    if company_code[:3] == 'RYW':
                        pass
                    else:
                        if pd.isnull(ADPExport.loc[indexnumber, 'Preferred or Chosen Last Name']):
                            directreportlastname = ADPExport.loc[indexnumber, 'Last Name']
                        else:
                            directreportlastname = ADPExport.loc[indexnumber, 'Preferred or Chosen Last Name']                    
                        if pd.isnull(ADPExport.loc[indexnumber, 'Preferred or Chosen First Name']):
                            directreportfirstname = ADPExport.loc[indexnumber, 'First Name']
                        else:
                            directreportfirstname = ADPExport.loc[indexnumber, 'Preferred Or Chosen First Name']

                        username = f'{directreportfirstname} {directreportlastname}'
                        aduser.ADUser.create(username, AESaccountsOU, password='###Newpassword', optional_attributes={'employeeID': directreport_associateID, })
                    
                    #starts the loop from the top on the same iteration to update the attributes of the newly created user.
                        continue
                else:
                    for duser in directreportADarray:
                        directreport_userAccountControl = duser.get('userAccountControl')
                        directreportDN = duser.get('distinguishedName')
                        currentmanager = str(duser.get('manager'))
                        directreportCN = duser.get('cn')
                        directreport_extentionattribute = duser.get('extensionAttribute1')

                    print(directreportCN)

                    directreportAD = aduser.ADUser.from_dn(directreportDN)
                    directreportAD.update_attribute('employeeNumber', directreport_positionID)

                    if pd.isnull(ADPExport.loc[indexnumber, 'Reports To Associate ID']):
                        hasmanager = False
                    else:
                        #uses the manager's name to find the instance of that person within the same report to get their PositionID. Uses this to find the manager in Active Directory.
                        newmanager_assocaiteID = ADPExport.loc[indexnumber, 'Associate ID']
                        newmanagerADarray = session.find_users_by_attribute('employeeID', newmanager_assocaiteID, ['distinguishedname', 'cn'])
                        
                        if newmanagerADarray != []:
                            hasmanager = True

                    for muser in newmanagerADarray:
                        newmanagerDN = muser.get('distinguishedname')
                        newmanagerCN = muser.get('cn')
                    print(newmanagerDN)

                    location = str(ADPExport.loc[indexnumber, 'Location Description'])

                    if pd.isnull(location):
                        office = 'NA'
                    else:
                        office = location

                    description = str(ADPExport.loc[indexnumber, 'Job Title Code'])

                    if pd.isnull(description):
                        description = 'NA'

                    department = str(ADPExport.loc[indexnumber, 'Home Department Description'])

                    if pd.isnull(department):
                        department = 'NA'

            #creates an array of the attributes within the AES report to be compaired against later. This will be used to detect changes.
                    exportattributes = []
                    exportattributes.append(str(newmanagerDN))
                    exportattributes.append(str(location))
                    exportattributes.append(description)
                    exportattributes.append(department)
                
                    #creates an array for the current attributes of the user and appends value of current attributes
                    currentattributes = []

                    for user in directreportADarray:
                        currentattributes.append(user.get('cn'))
                        currentattributes.append(user.get('manager'))
                        try:
                            arraytostrdescription = ''.join(user.get('description'))
                        except TypeError:
                            arraytostrdescription = None
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
                    
                    if directreport_userAccountControl == 514 and directreport_extentionattribute == 'Leave':
                        directreportAD.update_attribute('userAccountControl', 512)
                        directreportAD.clear_attribute_attribute('extensionAttribute1')
                        print('returning leave user enabled')
                        
            else:
                nonactive_indexstorage.append(indexnumber)
                    
        except:
            #FUNCTION BELOW NOT WORKING
            print('broke')
            with open(unresolvedexport, 'a', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(reportfields)
                writer.writerow(ADPExport.loc[indexnumber, :]) 
        indexnumber+=1

    return nonactive_indexstorage

def getinfo_leave_term(rownum):
    duser_groupCN = []
    status = str(ADPExport.loc[rownum, 'Position Status'])
    directreport_associateID = ADPExport.loc[rownum, 'Associate ID']
    directreport_positionID = ADPExport.loc[rownum, 'Position ID']
    directreportADarray = session.find_users_by_attribute('employeeID', directreport_associateID, ['cn', 'department', 'description', 'distinguishedName', 'employeeNumber', 'homeDirectory', 'manager', 'memberOf', 'objectGUID', 'objectSid', 'physicalDeliveryOfficeName', 'proxyAddresses', 'whenCreated'])

    if directreportADarray == []:
        pass

    else:
        try:
            for term_user in directreportADarray:
                duser_dn = term_user.get('distinguishedName')
                duser_employeeNumber = term_user.get('employeeNumber')
            directreportAD = aduser.ADUser.from_dn(duser_dn)

            if status == 'Leave':
                directreportAD.disable()
                directreportAD.update_attribute('extentionAttribute1', 'Leave')
                print('leave user disabled')
            else:
                if directreport_positionID == duser_employeeNumber:
                    for duser in directreportADarray:
                        duser_cn = duser.get('cn')
                        duser_department = duser.get('department')
                        duser_description = duser.get('description')
                        duser_homedir = duser.get('homeDirectory')
                        duser_manager = duser.get('manager')
                        duser_groups = duser.get('memberOf')
                        duser_GUID = duser.get('objectGUID')
                        duser_SID = duser.get('objectSid')
                        duser_location = duser.get('physicalDeliveryOfficeName')
                        duser_proxies = duser.get('proxyAddresses')
                        duser_creation_date = duser.get('whenCreated')

                    for group in duser_groups:
                        duser_groupobject = adgroup.ADGroup(group)
                        duser_groupCN.append(duser_groupobject.get_attribute('cn'))

                    if not os.path.exists(f'{archive_dir}{duser_cn}'):
                        os.makedirs(f'{archive_dir}{duser_cn}')

                    with open(f'{archive_dir}{duser_cn}/AD Properties Archive', 'w') as f:
                        f.write(duser_cn)
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Department: {str(duser_department)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Description: {str(duser_description)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Home Directory: {str(duser_homedir)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Manager: {str(duser_manager)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'ObjectGUID: {str(duser_GUID)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'ObjectSID: {str(duser_SID)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Location: {str(duser_location)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Proxy Addresses: {str(duser_proxies)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Creation Date: {str(duser_creation_date)}')
                        f.write('\n')
                        f.write('\n')
                        f.write(f'Groups {str(duser_groupCN)}')

                    directreportAD.disable()
                    memberof_groups = directreportAD.get_attribute('memberOf')
                    
                    for group in memberof_groups:
                        group_object = adgroup.ADGroup(group)
                        directreportAD.remove_from_group(group_object)
                    
                    directreportAD.clear_attribute('manager')
                    directreportAD.clear_attribute('telephoneNumber')
                    directreportAD.update_attribute('description', str(date.today()))
        except:
            print('term broke')

if 

    with open('C:/Users/######/AES/AES Report.csv', 'r') as f:
        reader = csv.reader(f)
        numberofusers = sum(1 for rown in reader) - 1

    notactive = getinfo_update(numberofusers)

    for indicies in notactive:
        print(indicies)
        getinfo_leave_term(indicies)

            

    #erases the changesdetected csv for the next execution.
    with open(changesdetected, 'w', newline='') as f:
        f.write('')
        writer = csv.writer(f)
        writer.writerow(reportfieldschange)
