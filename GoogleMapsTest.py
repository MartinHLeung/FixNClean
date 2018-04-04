import time
import datetime
from datetime import date
import googlemaps

import xlrd, xlwt
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


def month_string_to_number(string):
    m = {
        'January': 1,
        'February': 2,
        'March': 3,
        'April':4,
         'May':5,
         'June':6,
         'July':7,
         'August':8,
         'sep':9,
         'oct':10,
         'nov':11,
         'dec':12
        }
    s = string.strip()[:3].lower()

    try:
        out = m[s]
        return out
    except:
        raise ValueError('Not a month')

def converttime(int):
    if int<6:
        int = int+12
    return int


linkarray = []
api_key = 'AIzaSyDF02j1j4horv-sjTz7-Akp3EHGGANIMyA'
gmaps = googlemaps.Client(key=api_key)


for i in communityInformation.keys():

    geocode_result = gmaps.geocode('45 Union St W, Kingston, ON')
    address = communityInformation[i][4]
    address = address.replace(" ","+")
    time = dates[communityInformation[i][3]]
    now = datetime.datetime.now()
    arr = time.split(' ')
    day = arr[2][:-2]
    month = arr[1]
    year = now.year
    hour = arr[4][:1]
    hour = converttime(int(hour))
    epocht = int(datetime.datetime(year,month_string_to_number(month),int(day),hour,0).timestamp())
    dateinfo = datetime.datetime(year,month_string_to_number(month),int(day),hour,0)
    directions_result = gmaps.directions("45 Union St W, Kingston, ON",
                                         communityInformation[i][4],
                                         mode="transit",
                                         departure_time=dateinfo)

    googledata = "/data=!4m6!4m5!2m3!6e1!7e3!8j" + str(epocht) + "!3e3"
    htmladdress = 'https://www.google.com/maps/dir/45+Union+St+W,+Kingston,+ON/' + address + googledata
    duration = directions_result[0]['legs'][0]['duration']['text']
    minutes = duration.split(' ')
    if minutes[1] == "hours":
        linkarray.append("take private transit")
        continue
    elif minutes[1] == "hour":
        linkarray.append("take private transit")
        continue
    elif int(minutes[0]) > 30:
        linkarray.append("take private transit")
        continue
    else : linkarray.append(htmladdress)

