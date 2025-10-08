import pandas as pd
from datetime import datetime
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
storage_path = os.path.join(script_dir, "storage.txt")

os.chdir(script_dir); os.path.abspath(script_dir); os.chdir("..")
spreadsheet_path = os.path.join(os.path.abspath(os.curdir), "A-Level Spreadsheet.xlsx")

def matchToTables(v, correlatingTable, doNotClear: bool | None = None, returnVariables: bool | None = None):
    if len(v) == 0: v = " "
    if not doNotClear:
        os.system("cls")
    matches = {
        "Any": {},
        "Linear": {},
        "Totals": {}
    }
    lowerv = v.lower()
    for b in correlatingTable:
        matches["Any"][b] = 0
        matches["Linear"][b] = 0
        strval = b.lower()
        for i in range(len(lowerv)):
            if len(strval)-1 >= i and lowerv[i] == strval[i]:
                matches["Linear"][b] += 1
        for i in range(len(strval)):
            for n in range(len(lowerv)):
                if lowerv[n] == strval[i]:
                    matches["Any"][b] += 1
        matches["Totals"][b] = (matches["Linear"][b] + (matches["Any"][b]/len(lowerv)))*(3/2)
    max, var = 0, ""
    for m in matches["Totals"]:
        value = matches["Totals"][m]
        if value > max:
            max = value
            var = m
    if not returnVariables:
        return var
    else:
        return var, matches      
def beautifulPrint(tableToPrint, limit = 9999999):
    iterations = 0
    for v in tableToPrint:
        iterations += 1
        print(" {iterations}. {v}".format(v=v, iterations=iterations))
        if iterations > limit-1: break
def getLines(f):
    return f.readline().strip(), f.readline().strip(), f.readline().strip(), int(f.readline().strip()), f.readline().strip(), int(f.readline().strip())
def orderNumberArray(array):
    newarray = []
    for i in range(len(array)):
        max, val, iter, siter = 0, "", 1, 0
        for a in array:
            b = array[a]
            if b > max:
                max = b
                val = a
                siter = iter
            iter += 1
        if max != 0:
            del array[val]
            newarray.append(val)
    if len(newarray) > 0:
        return newarray
    else:
        return ["No values matched"]
def clamp(n, min_value, max_value):
    return max(min_value, min(n, max_value))


datatables = {
    "Types":["Physical","Digital"],
    "Subject":["Mathematics","Further Mathematics","Physics"],
    "Decision":["Yes","No"],
    "Alphabet":["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"],
    "InitDecisions":["Spreadsheet","Name generator"],
    "sResponse":["Searching for a topic","Identify key areas of improvement"]
}

while True:
    response = matchToTables(input("Available services: Spreadsheet Management | Name Generator\n\nRequested service?: "), datatables["InitDecisions"])
    if response == "Spreadsheet":
        sheet1 = pd.read_excel(spreadsheet_path, sheet_name=0, header=1, index_col=1)
        sresponse = matchToTables(input("Search for a topic, or identify key areas of improvement?: "), datatables["sResponse"])
        if sresponse == "Searching for a topic":
            while True:
                searchResults, additionalResults = matchToTables(input("Value to find?: "), sheet1["Topic"], True, True)
                print("Search results: "); beautifulPrint(orderNumberArray(additionalResults["Totals"]), 3); input()
                if matchToTables(input("Return to menu?: "), datatables["Decision"]) == "Yes":
                    break
                else:
                    continue
        elif sresponse == "Identify key areas of improvement":
            while True:
                subjectsOfReviewInterest = {}
                for b in sheet1["Topic"]:
                    row = int(pd.Index(sheet1["Topic"]).get_loc(b))
                    v = sheet1.iloc[row]["Date Reviewed"]
                    split = str(v).split("-")
                    calculation = datetime.now().timestamp()-datetime(int(split[0]), int(split[1]), int(split[2].split(" ")[0])).timestamp() 
                    if calculation >= 1209600:
                        subjectsOfReviewInterest[b] = [calculation/1209600, row]
                for d in subjectsOfReviewInterest:
                    data = subjectsOfReviewInterest[d]
                    if sheet1.iloc[data[1]]["Strength"]:
                        data[0] /= ((int(sheet1.iloc[data[1]]["Strength"])-(data[0]))/data[0])
                if len(subjectsOfReviewInterest) > 0:
                    reformat = {}
                    for a in subjectsOfReviewInterest:
                        data = subjectsOfReviewInterest[a]
                        if data[0] > 0.5:
                            reformat[a] = data[0]
                    print("Areas identified which could be improved: ")
                    beautifulPrint(reformat)
                    input()
                else:
                    print("No areas of improvement.\n")
                if matchToTables(input("Return to menu?: "), datatables["Decision"]) == "Yes":
                    break
                else:
                    continue
    elif response == "Name generator":
        while True:
            typeofFile = matchToTables(input("Type of file: "), datatables["Types"])
            result = ""
            if typeofFile:
                if typeofFile == "Physical":
                    currentindex = 1
                    line1, line2, line3, line4, line5, line6 = "", "","",0,"",0
                    with open(storage_path) as f:
                        line1, line2, line3, line4, line5, line6 = getLines(f)
                        line2v = datatables["Alphabet"].index(line2)+1
                        if matchToTables(input("Create a new group, decision: "),datatables["Decision"]) == "Yes":
                            if line4 >= 20:
                                line4 = 1
                                line3v = datatables["Alphabet"].index(line3)+1
                                if line3v >= 26:
                                    line3 = "A"
                                    line2v += 1
                                else:
                                    line3 = datatables["Alphabet"][line3v]
                                    line2 = datatables["Alphabet"][line2v-1]
                            else:
                                line4 += 1
                        x = datetime.now().strftime("%x")
                        if line5 != x:
                            line5 = x
                            line6 = 1
                        else:
                            line6 += 1
                        result = result + line1+line2+line3+str(line4)
                        currentindex = line6
                    result = result + typeofFile[0].upper()
                else:
                    result = result + typeofFile[0].upper()
                    with open(storage_path) as f:
                        line1, line2, line3, line4, line5, line6 = getLines(f)
                        x = datetime.now().strftime("%x")
                        if line5 != x:
                            line5 = x
                            line6 = 1
                        else:
                            line6 += 1
                        currentindex = line6
                with open(storage_path, "w") as f:
                    f.writelines([line1, "\n"+line2, "\n"+line3, "\n"+str(line4), "\n"+str(line5), "\n"+str(line6)])
                subject = matchToTables(input("What subject: "),datatables["Subject"])
                if subject == "Further Mathematics":
                    result = result + "FM"
                else:
                    result = result + subject[0]
                x = datetime.now().strftime("%A")
                if x == "Thursday" or x == "Sunday":
                    result = result + x[0] + x[1]
                else:
                    result = result + x[0]
                result = result + datetime.now().strftime("%m")
                result = result + str(int(datetime.now().strftime("%d"))+currentindex)
                if typeofFile == "Physical":
                    result = result + str(input("Ring Binder Code: "))
                print(result.upper())
                input()
            else:
                print("Unable to resolve a type of file")
            if matchToTables(input("Return to menu?: "), datatables["Decision"]) == "Yes":
                break
            else:
                continue