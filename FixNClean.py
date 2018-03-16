import xlrd
from os.path import join, dirname, abspath

fname = join(dirname(abspath("Volunteer info survey (Responses)")), "Volunteer info survey (Responses).xlsx")


class group:
    def __init__(self, members, timeslot, allergies, groupnumber):
        self.members = members
        self.timeslot = timeslot
        self.allergies = allergies
        self.client = None
        self.groupnumber = groupnumber

    def addmember(self, newmember, newallergies):
        self.members.append(newmember)
        for i in newallergies:
            if i not in self.allergies:
                self.allergies.append(i)

    def merge(self, othergroup):
        for i in othergroup.members:
            self.addmember(i, i.allergies)


volunteerExcel = xlrd.open_workbook(fname)

sheet = volunteerExcel.sheet_by_index(0)

groups = {}

volunteerInformation = {}
groupCounter = 0
dates = ['Saturday September 30th - 9:00-12:00', 'Saturday September 30th - 1:00-4:00',
         'Sunday October 1st - 9:00-12:00', 'Sunday October 1st - 1:00-4:00', "Only my first option works :("]

signuporder = []
for i in range(1, sheet.nrows):
    for j in range(5):
        sn = str(sheet.cell(i, 1 + j * 4))[7:-2]
        volunteerInformation[sn] = [str(sheet.cell(i, 2 + j * 4))[6:-1], str(sheet.cell(i, 3 + j * 4))[6:-1],
                                    dates.index(str(sheet.cell(i, 20))[6:-1]),
                                    dates.index(str(sheet.cell(i, 21))[6:-1]),
                                    [l for l in str(sheet.cell(i, 22))[6:-1].split(",")], -1]

        signuporder.append(str(sheet.cell(i, 1 + j * 4))[7:-2])
        if str(sheet.cell(i, 4 + j * 4))[6:-1] == "No":
            break
        elif j > 0:
            groups[str(groupCounter)].addmember(sn, volunteerInformation[sn][-1])
            volunteerInformation[sn][-1] = groupCounter
        else:
            groupCounter += 1
            groups[str(groupCounter)] = group([sn], volunteerInformation[sn][2:4], volunteerInformation[sn][-1],
                                              groupCounter)
            volunteerInformation[sn][-1] = groupCounter
