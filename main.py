# openpyxl for excel worksheet opening
import openpyxl as oxl
# pyexcel_ods for excel and calc open standard opening
from pyexcel_ods import get_data as calc
# csv for creating comma separated value files
import csv
# os.listdir gets files in a directory
from os import listdir
from os.path import isfile, join
from os import remove

# Constants
PATH = "/home/user/Documents/TestFeedback"
CLASSRANGE = [1,12]
LEADERRANGE = [16,19]
level5 = []
level6 = []



# Classes for storing values (can scroll through)
class person:
    # glorified dictonary
    def __init__(self, name, rel, team, crt, prod):
        self.name = name
        self.rel = rel
        self.team = team
        self.crt = crt
        self.prod = prod

    # formatting print command
    def __str__(self):
        return f"{self.name},{self.rel},{self.team},{self.crt},{self.prod}"

    def isName(self, name):
        if name == self.name:
            return True
        return False

    def getName(self):
        return self.name

    def getCSV(self):
        return f"{self.rel},{self.team},{self.crt},{self.prod}"

class leader:
    def __init__(self, name, mot, fair, cpt, cmt):
        self.name = name
        self.mot = mot
        self.fair = fair
        self.cpt = cpt
        self.cmt = cmt

    def isName(self, name):
        if name == self.name:
            return True
        return False

    def getName(self):
        return self.name

    def getCSV(self):
        return f"{self.mot},{self.fair},{self.cpt},{self.cmt}"



def main():
    # grabs file names from a path
    files = get_files(PATH)
    amountOfFiles = len(files)
    # sets up value stores

    # for each feedback form
    for file in files:
        # store the file type i.e. xls, xlsx, csv, or ods
        extension = file.split(".")[1]

        # if its a standard excel format
        if extension == "xls" or extension == "xlsx":
            convert_xsl_dict(file)

        # if its an open standard
        elif extension == "ods":
            convert_ods_dict(file)


        else:
            throwErr("format", file)

    # for person in level5:
    #     print(person)
    # for leader in level6:
    #     print(leader)
    # should equal 12 and 4 respectively (class amounts)
    print(f"classmates accounted : {len(level5)/amountOfFiles}\nleaders accounted    : {len(level6)/amountOfFiles}")


def convert_ods_dict(file):
    ods = calc(f"{PATH}/{file}")
    # for each team member
    for i in range(CLASSRANGE[0], CLASSRANGE[1] + 1):
        # grab their row
        meCheck = False
        p = ods["Sheet1"][i]
        me = p[9]
        # make sure their score is correctly inputted or throw an error
        if p[6] == 0:
            # create object and add it to stores
            e = person(p[0], p[1], p[2], p[3], p[4])
            level5.append(e)
        elif me == 1 and meCheck is False:
            meCheck = True
        else:
            level5.append({"name":p[0], "error":"wrong values"})
            throwErr("total", f"{file}:{p[0]}")

    # works the same as above just using a different object for ease later
    for i in range(LEADERRANGE[0], LEADERRANGE[1] + 1):
        meCheck = False
        l = ods["Sheet1"][i]
        me = l[9]
        if l[6] == 0:
            e = leader(l[0], l[1], l[2], l[3], l[4])
            level6.append(e)
        elif me == 1 and meCheck is False:
            meCheck = True
        else:

            level6.append({"name": l[0], "error": "wrong values"})
            throwErr("total", f"{file}:{l[0]}")

def convert_xsl_dict(file):
    xsl = oxl.load_workbook(PATH + "/" + file)
    sht = xsl.active
    for row in sht.iter_rows(min_row=CLASSRANGE[0]+1, max_row=CLASSRANGE[1]+1, min_col=0, max_col=10):
        meCheck = False
        me = row[9].value
        n = row[0].value
        r = row[1].value
        t = row[2].value
        c = row[3].value
        p = row[4].value

        if r + t + c + p == 100:
            e = person(n, r, t, c, p)
            level5.append(e)
        elif me == 1 and meCheck is False:
            meCheck = True
        else:
            if n != "Members":
                level5.append({"name": n, "error": "wrong values"})
            throwErr("total", f"{file}:{n}")

    for row in sht.iter_rows(min_row=LEADERRANGE[0]+1, max_row=LEADERRANGE[1]+1, min_col=0, max_col=10):
        meCheck = False
        me = row[9].value
        n = row[0].value
        r = row[1].value
        t = row[2].value
        c = row[3].value
        p = row[4].value

        if r + t + c + p == 100:
            e = leader(n, r, t, c, p)
            level6.append(e)
        elif me == 1 and meCheck is False:
            meCheck = True
        else:

            level6.append({"name": n, "error": "wrong values"})
            throwErr("total", f"{file}:{n}")



# returns a list of file names from a path
def get_files(directory):
    files = [f for f in listdir(directory) if isfile(join(directory, f))]
    return files

def throwErr(err, loc):
    if err == "total":
        print(f"ERROR: {loc} has not been totalled correctly")
    if err == "format":
        print(f"ERROR: {loc} is not formatted correctly, only allows xslx, xsl, and ods")






if __name__ == '__main__':
    main()

