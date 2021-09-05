from groupy.client import Client
import os.path
import xlrd
from xlwt import *

token = "VH1pehGboi1Pe4CDGUGeqF8YKdQBnEK58ID3VvKL"

wb = xlrd.open_workbook(os.path.join(r'C:\Users\jmkre\OneDrive\Desktop\free-time-coding\purgebot', 'registerfinal.xlsx'))
namesheet = wb.sheet_by_name('Responses')
# 02, first name    03 preferred name    04 last name

#getting list of registered members with as many variations of names as possible from excel file
registeredMembersNameList = []
count = 1
while(count < namesheet.nrows - 1):
    firstname = namesheet.cell_value(count, 2).lower().strip() #get first name
    registeredMembersNameList.append(firstname) #add first name check
    preferredname = namesheet.cell_value(count, 3).lower().strip() #get preferred name
    lastname = namesheet.cell_value(count, 4).lower().strip() #get last name
    registeredMembersNameList.append(lastname) #add last name
    if(len(lastname) > 0):
        lastnameInitial = namesheet.cell_value(count, 4)[0].lower()
        registeredMembersNameList.append(firstname + " " + lastnameInitial) #get first name + last name initial
    registeredMembersNameList.append(firstname + " " + lastname)
    count += 1


#Groupy api client connection, fetch list of group members from group object
client = Client.from_token(token)
runclubgroup = client.groups.get(id=69513817)
rcGroupMembers = runclubgroup.members

#creating new excel sheet to match names with member ids and if they're registered or not
wb = Workbook()
idlist = []
bigsheet = wb.add_sheet('blah', cell_overwrite_ok=True)
bigsheet.write(0,0,'First')
bigsheet.write(0,1,'Last')
bigsheet.write(0,2,'MemberId')
bigsheet.write(0,3,'First Last')
bigsheet.write(0,4,'Kick (1 | 0)')
groupmeNameList = [m.nickname.lower().strip() for m in rcGroupMembers]
count = 1
for m in rcGroupMembers:
    name = m.nickname.lower().strip()
    namesplit = name.split()
    first = namesplit[0] #write the first name in the excel sheet
    bigsheet.write(count,0,first)

    #if they have a last name implement checks to whitelist members
    if(len(namesplit) > 1):
        last = namesplit[1]

        #if their last name matches a registered member names
        if (last in registeredMembersNameList):
            bigsheet.write(count, 4, '1')
            idlist.append(m.id) #whitelist member id

        #if their first + last initial is in registered member names
        if ((first + " " + last[0]) in registeredMembersNameList):
            bigsheet.write(count, 4, '1')
            idlist.append(m.id) #whitelist member id

        bigsheet.write(count,1,last) #write to last name column in excel sheet
        firstlast = first + " " + last

        #if their full combined name is in registered member names
        if(firstlast in registeredMembersNameList):
            bigsheet.write(count, 4, '1')
            idlist.append(m.id) #first + last name is in whitelist
        bigsheet.write(count, 3, firstlast)

    #if they don't have a full name just write their first to the combined column
    else:
        bigsheet.write(count, 3, first)
        bigsheet.write(count, 1, '') #empty lastname

    bigsheet.write(count, 2, m.id) #write member id
    count+=1
wb.save('bookie.xls')

idlist = list(dict.fromkeys(idlist)) #final list of all whitelisted/registered member ids

#remove all members from group who are not on whitelist
for m in rcGroupMembers:
    if m.id not in idlist:
         # m.remove()
