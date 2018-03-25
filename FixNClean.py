import xlrd
from os.path import join, dirname, abspath
fname = join(dirname(abspath("Volunteer info survey (Responses)")), "Volunteer info survey (Responses).xlsx")
cname = join(dirname(abspath("Community member information (Responses)")), "Community member information (Responses).xlsx")

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
        if not ((volunteerInformation[newmember][2:4] == self.timeslot) or (volunteerInformation[newmember][2:4] == self.timeslot[::-1])):
            newtimeslot = []
            for i in volunteerInformation[newmember][2:4]:
                if i in self.timeslot:
                    newtimeslot.append(i)
            newtimeslot.append(4)
            self.timeslot = newtimeslot

    def merge(self, othergroup):
        for i in othergroup.members:
            self.addmember(i, i.allergies)

def redistribute(times, start, target):
    for i in times[start][::-1]:
        if volunteerInformation[i][3] == target:
            if volunteerInformation[i][-1] != -1:
                indices = [times[start].index(l) for l in groups[volunteerInformation[i][-1]].members]
                spot = [place(target, l) for l in groups[volunteerInformation[i][-1]].members]
                for j in indices:
                    times[target].insert(spot[j], times[start].pop(j))
                return times
            else:
                spot = place(target, i)
                times[target].insert(spot, times[start].pop(times[start].index(i)))
                return times

    return False

def place(lis, element):
    for i in range(len(lis)):
        if signuporder.index(lis[i])>signuporder.index(element):
            return i
    return len(lis)

volunteerExcel = xlrd.open_workbook(fname)

sheet = volunteerExcel.sheet_by_index(0)
csheet = xlrd.open_workbook(cname).sheet_by_index(0)
groups = {}

volunteerInformation = {}
communityInformation = {}
groupCounter = 0
dates = ['Saturday September 30th - 9:00-12:00', 'Saturday September 30th - 1:00-4:00',
         'Sunday October 1st - 9:00-12:00', 'Sunday October 1st - 1:00-4:00', "Only my first option works :("]



signuporder = []
volunteers = [[], [], [], []]
members = [0, 0, 0, 0]
for i in range(1,sheet.nrows):
    for j in range(5):
        sn = str(sheet.cell(i, 1+j*4))[7:-2]
        volunteerInformation[sn] = [str(sheet.cell(i, 2+j*4))[6:-1],
                                    str(sheet.cell(i, 3+j*4))[6:-1],
                                    dates.index(str(sheet.cell(i, 20))[6:-1]),
                                    dates.index(str(sheet.cell(i, 21))[6:-1]),
                                    [l for l in str(sheet.cell(i, 22))[6:-1].split(",")],False, -1]

        signuporder.append(sn)
        if str(sheet.cell(i, 4+j*4))[6:-1] == "No":
            break
        elif j>0:
            groups[str(groupCounter)].addmember(sn, volunteerInformation[sn][-3])
            volunteerInformation[sn][-1] = str(groupCounter)
        else:
            groupCounter += 1
            groups[str(groupCounter)] = group([sn], volunteerInformation[sn][2:4], volunteerInformation[sn][-3], groupCounter)
            volunteerInformation[sn][-1] = str(groupCounter)

for i in range(1, csheet.nrows):
    communityInformation[str(csheet.cell(i, 1))[6:-1]] = [str(csheet.cell(i, 2))[7:-2],
                                                          str(csheet.cell(i, 3))[6:-1],
                                                          str(csheet.cell(i, 4))[6:-1],
                                                          dates.index(str(csheet.cell(i, 5))[6:-1]),
                                                          str(csheet.cell(i, 6))[6:-1],
                                                          [l for l in str(csheet.cell(i, 7))[6:-1].split(",")]]
    members[communityInformation[str(csheet.cell(i, 1))[6:-1]][3]]+=1




averageGroupSize=len(signuporder)/len(communityInformation.keys())



tolerance = 0.25

organized = False
for i in range(4):
    w = 0
    while len(volunteers[i])/members[i]<averageGroupSize and len(volunteers[i])/members[i]<5:
        if w>len(signuporder):
            break
        if volunteerInformation[signuporder[w]][2] == i or volunteerInformation[signuporder[w]][3] == i:
            if volunteerInformation[signuporder[w]][-1] != -1:
                for k in groups[str(volunteerInformation[signuporder[w]][-1])].members:
                    w+=1
                    volunteers[i].append(k)
            else:
                volunteers[i].append(signuporder[w])
                w+=1
averageByTime = [len(volunteers[i])/members[i] for i in range(4)]

while not organized:
    if all(averageGroupSize-tolerance<=averageByTime[i]<=averageGroupSize+tolerance for i in range(4)):
        organized = True
    else:
        count = [len(i) for i in volunteers]
        volunteers = redistribute(volunteers, count.index(max(count)), count.index(min(count)))

for i in range(4):
    w=0
    while w<len(volunteers[i]):
        if volunteerInformation[volunteers[i][w]][-1] != -1:
            for x in groups[volunteerInformation[volunteers[i][w]][-1]].members:
                volunteerInformation[x][-2] = True
            size = int(averageGroupSize)-len(groups[volunteerInformation[volunteers[i][w]][-1]].members)
            while size>0 and len(groups[volunteerInformation[volunteers[i][w]][-1]].members)<int(averageGroupSize):
                for x in range(len(volunteers[i])):
                    if volunteerInformation[volunteers[i][x]][-1] != -1:
                        similar = volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4] or volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4][::-1]
                        if size == len(groups[volunteerInformation[volunteers[i][x]][-1]].members) and similar:
                            for v in groups[volunteerInformation[volunteers[i][x]][-1]].members:
                                volunteerInformation[v][-2] = True
                            groups[volunteerInformation[volunteers[i][w]][-1]].merge(groups[volunteerInformation[volunteers[i][x]][-1]])

                    else:
                        similar = volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4] or volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4][::-1]
                        if size == 1 and similar:
                            volunteerInformation[volunteers[i][x]][-2] = True
                            groups[volunteerInformation[volunteers[i][w]][-1]].addmember(volunteerInformation[volunteers[i][x]])
                size-=1
        else:
            groupCounter += 1
            groups[str(groupCounter)] = group([sn], volunteerInformation[sn][2:4], volunteerInformation[sn][-1],
                                              groupCounter)
            volunteerInformation[sn][-1] = str(groupCounter)
            volunteerInformation[volunteers[i][w]][-2] = True
            size = int(averageGroupSize) - len(groups[volunteerInformation[volunteers[i][w]][-1]].members)
            while size > 0 and len(groups[volunteerInformation[volunteers[i][w]][-1]].members) < int(averageGroupSize):
                for x in range(len(volunteers[i])):
                    similar = volunteerInformation[volunteers[i][w]][2:4] == volunteerInformation[volunteers[i][x]][2:4] or volunteerInformation[volunteers[i][w]][2:4] == volunteerInformation[volunteers[i][x]][2:4][::-1]

                    if volunteerInformation[volunteers[i][x]][-1] != -1:
                        if size == len(groups[volunteerInformation[volunteers[i][x]][-1]].members) and similar:
                            for v in groups[volunteerInformation[volunteers[i][x]][-1]].members:
                                volunteerInformation[v][-2] = True
                            groups[volunteerInformation[volunteers[i][w]][-1]].merge(
                                groups[volunteerInformation[volunteers[i][x]][-1]])

                    else:
                        if size == 1 and similar:
                            volunteerInformation[volunteers[i][x]][-2] = True
                            groups[volunteerInformation[volunteers[i][w]][-1]].addmember(
                                volunteerInformation[volunteers[i][x]])
                size -= 1
        while volunteerInformation[volunteers[i][w]] and w<len(volunteers[i]):
            w += 1
    leftovers = []
    for f in range(len(volunteers[i])):
        if not volunteerInformation[volunteers[i][f]][-2]:
            leftovers.append(volunteers[i][f])
    for l in leftovers:
        if  volunteerInformation[l][-1] != -1:
            size = 5-len(groups[volunteerInformation[l][-1]].members)
            for g in groups.keys():
                if len(groups[g].members) <= size:
                    groups[g].merge(groups[volunteerInformation[l][-1]])
        else:
            size = 4
            for g in groups.keys():
                if len(groups[g].members) <= size:
                    groups[g].merge(groups[volunteerInformation[l][-1]])
