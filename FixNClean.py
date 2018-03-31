import xlrd, xlwt, time
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
        self.switched = False
    def addmember(self, newmember, newallergies):
        self.members.append(newmember)
        if self.allergies == -1:
            self.allergies = ["None"]
        for i in newallergies:
            if not(i in self.allergies):
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
            self.addmember(i, othergroup.allergies)


def redistribute(times, start, target):
    for i in reversed(times[start]):
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
         'Sunday October 1st - 9:00-12:00', 'Sunday October 1st - 1:00-4:00', "Only my first choice works for me :("]



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
                                    [l for l in str(sheet.cell(i, 22))[6:-1].split(", ")],False, -1]

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
                                                          [l for l in str(csheet.cell(i, 7))[6:-1].split(", ")]]
    members[communityInformation[str(csheet.cell(i, 1))[6:-1]][3]]+=1




averageGroupSize=len(signuporder)/len(communityInformation.keys())

tolerance = 0.25

organized = False
for i in range(4):
    w = 0
    while len(volunteers[i])/members[i]<averageGroupSize-tolerance and len(volunteers[i])/members[i]<5:
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
        else:
            w+=1
averageByTime = [len(volunteers[i])/members[i] for i in range(4)]
leftovers = []
leave = 0
while not organized and leave<1000:
    if all(averageGroupSize-tolerance<=averageByTime[i]<=averageGroupSize+tolerance for i in range(4)):
        organized = True
    else:
        count = [len(i) for i in volunteers]
        yeet = redistribute(volunteers, count.index(max(count)), count.index(min(count)))
        if yeet:
            volunteers = yeet
        leave +=1


if averageGroupSize>5:
    averageGroupSize = 5


for i in range(4):
    w=0
    while w<=len(volunteers[i])-2:
        if volunteerInformation[volunteers[i][w]][-1] != -1:
            for x in groups[volunteerInformation[volunteers[i][w]][-1]].members:
                volunteerInformation[x][-2] = True
            size = int(averageGroupSize)-len(groups[volunteerInformation[volunteers[i][w]][-1]].members)
            while size>0 and len(groups[volunteerInformation[volunteers[i][w]][-1]].members)<int(averageGroupSize):
                for x in range(len(volunteers[i])):
                    if not volunteerInformation[volunteers[i][x]][-2]:
                        if volunteerInformation[volunteers[i][x]][-1] != -1:
                            similar = volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4] or volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4][::-1]
                            if size == len(groups[volunteerInformation[volunteers[i][x]][-1]].members) and similar:
                                for v in groups[volunteerInformation[volunteers[i][x]][-1]].members:
                                    volunteerInformation[v][-2] = True
                                size -= len(groups[volunteerInformation[volunteers[i][x]][-1]].members)
                                groups[volunteerInformation[volunteers[i][w]][-1]].merge(groups[volunteerInformation[volunteers[i][x]][-1]])
                                #groups.pop(volunteerInformation[volunteers[i][w]][-1])

                        else:
                            similar = volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4] or volunteerInformation[volunteers[i][w]][2:4]== volunteerInformation[volunteers[i][x]][2:4][::-1]
                            if size == 1 and similar:
                                volunteerInformation[volunteers[i][x]][-2] = True
                                groups[volunteerInformation[volunteers[i][w]][-1]].addmember(volunteers[i][x], volunteerInformation[volunteers[i][x]][4])
                                size -= 1
                size-=1
        else:#aaaaaaaaaaaaaaaaaaa
            groupCounter += 1
            sn = volunteers[i][w]
            groups[str(groupCounter)] = group([sn], volunteerInformation[sn][2:4], volunteerInformation[sn][-1],
                                              groupCounter)
            volunteerInformation[sn][-1] = str(groupCounter)
            volunteerInformation[volunteers[i][w]][-2] = True
            size = int(averageGroupSize) - 1
            while size > 0 and len(groups[volunteerInformation[volunteers[i][w]][-1]].members) < int(averageGroupSize):
                for x in range(len(volunteers[i])):
                    if not volunteerInformation[volunteers[i][x]][-2]:
                        similar = volunteerInformation[volunteers[i][w]][2:4] == volunteerInformation[volunteers[i][x]][2:4] or volunteerInformation[volunteers[i][w]][2:4] == volunteerInformation[volunteers[i][x]][2:4][::-1] and volunteerInformation[volunteers[i][x]][4] != 4
                        if volunteerInformation[volunteers[i][x]][-1] != -1:
                            if size == len(groups[volunteerInformation[volunteers[i][x]][-1]].members) and similar:
                                for v in groups[volunteerInformation[volunteers[i][x]][-1]].members:
                                    volunteerInformation[v][-2] = True
                                size -= len(groups[volunteerInformation[volunteers[i][x]][-1]])
                                groups[volunteerInformation[volunteers[i][w]][-1]].merge(
                                    groups[volunteerInformation[volunteers[i][x]][-1]])
                                break

                                #groups.pop(volunteerInformation[volunteers[i][w]][-1])

                        else:
                            if size == 1 and similar:
                                volunteerInformation[volunteers[i][x]][-2] = True
                                size -= 1
                                groups[volunteerInformation[volunteers[i][w]][-1]].addmember(volunteers[i][x], volunteerInformation[volunteers[i][x]][4])
                                break
                size -= 1
        while volunteerInformation[volunteers[i][w]][-2] and w<=len(volunteers[i])-2:
            w += 1

    for f in range(len(volunteers[i])):
        if not volunteerInformation[volunteers[i][f]][-2]:
            leftovers.append(volunteers[i][f])
    for l in leftovers:
        if volunteerInformation[l][-1] != -1:
            size = 5 - len(groups[volunteerInformation[l][-1]].members)
            for g in groups.keys():
                if len(groups[g].members) < size and (
                        volunteerInformation[l][2] in groups[g].timeslot or volunteerInformation[l][3] in groups[
                    g].timeslot):
                    for v in groups[volunteerInformation[l][-1]].members:
                        volunteerInformation[v][-2] = True
                    size -= len(groups[volunteerInformation[l][-1]].members)
                    groups[g].merge(groups[volunteerInformation[l][-1]])
                    break
        else:
            size = 4
            for g in groups.keys():
                if len(groups[g].members) < size and (
                        volunteerInformation[l][2] in groups[g].timeslot or volunteerInformation[l][3] in groups[
                    g].timeslot):
                    groups[g].addmember(l, volunteerInformation[l][4])
                    volunteerInformation[l][-2] = True
                    size -= 4



for i in communityInformation.keys():
    for g in groups.keys():
        #print(communityInformation[i][4])
        if communityInformation[i][3] in groups[g].timeslot and groups[g].client == None:
            groups[g].client = i
            break
for g in groups.keys():
    if groups[g].client:

        time = communityInformation[groups[g].client][3]
        size = 5 - len(groups[g].members)
        while len(groups[g].members) < int(averageGroupSize) and size != 0:
            xyz=0
            while xyz <= len(signuporder)-1:
                if size > 1:
                    if volunteerInformation[signuporder[xyz]][-1] != -1:
                        if (not volunteerInformation[signuporder[xyz]][-2]) and len(groups[volunteerInformation[signuporder[xyz]][-1]].members) <= size:
                            for v in groups[volunteerInformation[signuporder[xyz]][-1]].members:
                                volunteerInformation[v][-2] = True
                            groups[g].merge(groups[volunteerInformation[signuporder[xyz]][-1]])
                            size -= len(groups[volunteerInformation[signuporder[xyz]][-1]].members)
                            xyz += len(groups[volunteerInformation[signuporder[xyz]][-1]].members)
                elif size == 1:
                    if (not volunteerInformation[signuporder[xyz]][-2]) and volunteerInformation[signuporder[xyz]][-1] == -1:
                        groups[g].addmember(signuporder[xyz], volunteerInformation[signuporder[xyz]][4])
                        volunteerInformation[signuporder[xyz]][-2] = True
                        xyz+=1
                        size -= 1
                xyz+=1
            size-=1




book1 = xlwt.Workbook()
sheet1 = book1.add_sheet("Sheet 1", cell_overwrite_ok=True)
count = [0,0,0,0]
for i in range(4):
    sheet1.write(0, i * 6, dates[i])


for g in groups.keys():
    if groups[g].client:
        down = communityInformation[groups[g].client][3]
        sheet1.write(1+count[down]*6,6*down, groups[g].client)
        sheet1.write(2 + count[down] * 6, 6 * down, communityInformation[groups[g].client][1])

        sheet1.write(3 + count[down] * 6, 6 * down, communityInformation[groups[g].client][2])

        for i in range(len(groups[g].members)):
            sheet1.write(i+1+count[down]*6, 6*down+1, volunteerInformation[groups[g].members[i]][0])
            sheet1.write(i + 1 + count[down] * 6, 6 * down + 2, groups[g].members[i])
            sheet1.write(i+1+count[down]*6, 6*down+3, volunteerInformation[groups[g].members[i]][1])

        count[(communityInformation[groups[g].client][3])] += 1



file = open("emails for first schedule.txt", "w+")

for g in groups.keys():
    if groups[g].client:
        for i in groups[g].members:
            file.write("Hey " + volunteerInformation[i][0] + ",\n Thanks for signing up for this year's Fix 'N' Clean!\n\nYou've been assigned to " + groups[g].client + "'s house, during the " + dates[communityInformation[groups[g].cient][3]] + " timeslot. The people in your group are:\n")
            for l in groups[g].members: file.write(volunteerInformation[l][0] + "\n")
            file.write("\n Please try and be there at the beginning of the time slot, a link to some directions to your community member's house can be found below\n")
            file.write(google maps link probably + "\n\n")
            file.write("Thanks again for participating,\nFix 'N' Clean Co-ordinators\n\n\n\n")


for g in groups.keys():
    if len(groups[g].timeslot) == 2 and groups[g].client:
        if groups[g].timeslot[1] != 4:
            for l in groups.keys():
                if len(groups[l].timeslot) == 2 and groups[l].client and groups[l].timeslot[1] != 4 and l != g and not(groups[g].switched or groups[l].switched):
                    if groups[g].timeslot == groups[l].timeslot or groups[g].timeslot == groups[l].timeslot[::-1]:
                        temp = groups[g].client
                        groups[g].client = groups[l].client
                        groups[l].client = temp
                        groups[l].switched = True
                        groups[g].switched = True
sheet2 = book1.add_sheet("Sheet 1", cell_overwrite_ok=True)
count = [0,0,0,0]
for i in range(4):
    sheet2.write(0, i * 6, dates[i])


for g in groups.keys():
    if groups[g].client:
        down = communityInformation[groups[g].client][3]
        sheet2.write(1+count[down]*6,6*down, groups[g].client)
        sheet2.write(2 + count[down] * 6, 6 * down, communityInformation[groups[g].client][1])

        sheet2.write(3 + count[down] * 6, 6 * down, communityInformation[groups[g].client][2])

        for i in range(len(groups[g].members)):
            sheet2.write(i+1+count[down]*6, 6*down+1, volunteerInformation[groups[g].members[i]][0])
            sheet2.write(i + 1 + count[down] * 6, 6 * down + 2, groups[g].members[i])
            sheet2.write(i+1+count[down]*6, 6*down+3, volunteerInformation[groups[g].members[i]][1])

        count[(communityInformation[groups[g].client][3])] += 1



book1.save("Schedule.xls")