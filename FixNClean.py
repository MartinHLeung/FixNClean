import xlrd
from os.path import join, dirname, abspath
fname = join(dirname(abspath("Volunteer info survey (Responses)")), "Volunteer info survey (Responses).xlsx")

class group:
    def __init__(self, members, timeslot, allergies):
        self.members = members
        self.timeslot = timeslot
        self.allergies = allergies
        self.client = None
    def addmember(self, newmember, newallergies):
        self.members.append(newmember)
        for i in newallergies:
            if i not in self.allergies:
                self.allergies.append(i)


volunteerExcel = xlrd.open_workbook(fname)

sheet = volunteerExcel.sheet_by_index(0)

groups = {}

volunteerInformation = {}
groupCounter = 0
dates = ['Saturday September 30th - 9:00-12:00', 'Saturday September 30th - 1:00-4:00', 'Sunday October 1st - 9:00-12:00', 'Sunday October 1st - 1:00-4:00', "Only my first option works :("]
for i in range(1,sheet.nrows):
    for j in range(5):
        volunteerInformation[str(sheet.cell(i, 1+j*4))[7:-2]] = [str(sheet.cell(i, 2+j*4))[7:-1], str(sheet.cell(i, 3+j*4))[7:-1], dates.index(str(sheet.cell(i, 20))[7:-1]), dates.index(str(sheet.cell(i, 21))[7:-1]), [l for l in str(sheet.cell(i, 22))[7:-1].split(",")]]
        if sheet.cell(i, 4+j*4) == "No":
            break
        elif j>0:
            groups[str(groupCounter)].addmember(str(sheet.cell(i, 1+j*4))[7:-2], volunteerInformation[str(sheet.cell(i, 1+j*4))[7:-2]][-1])
        else:
            groupCounter += 1
            groups[str(groupCounter)] = group([str(sheet.cell(i, 1+j*4))[7:-2]], volunteerInformation[str(sheet.cell(i, 1+j*4))[7:-2]][2:3], volunteerInformation[str(sheet.cell(i, 1+j*4))[7:-2]][-1])
