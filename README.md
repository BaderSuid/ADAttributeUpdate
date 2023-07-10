# ADAttributeUpdate
Uses Automated Export Servives (AES) reports generated by ADP to update the attribute of users and audit their changes.
Can be customized to match the format of any report. Just change the column name to match that of your report.
In a previous verison of this, I used the user's CN to query AD and find the user's object but there were too many naming discrepencies between HR's records and AD. To counter this, I asked HR to run a report of all users with their emloyeeIDs. I used this report to find as many users as I could by matching CNs and defining their "employeeID" and "employeeNumber" attributes. 

***we have 3 leading letters on the positionID that shows location, so I just put the raw positionID in employeeID and removed the first three characters to define th employeeNumber. 

I made the accounts that couldn't be found part of a post execution export and delt with them manually. Once this is done, a standard should be set to define these attributes when onboarding to ensure that they get picked up by this automation. You can also have the script create an account for all intances of empty arrays. (When a user can't be found by the attribute value specified, [] is returned.) I'll include the scirpt I made for the one time sync. 
