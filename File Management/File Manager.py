import pandas as pd
from datetime import datetime
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
storage_path = os.path.join(script_dir, "storage.txt")

os.path.abspath(script_dir); os.chdir("..")
spreadsheet_path = os.path.join(os.path.abspath(os.curdir), "A-Level Spreadsheet.xlsx")

def matchToTables(v, correlatingTable, doNotClear = None):
    if not doNotClear:
        os.system("cls")
    charactersMatched, results, = 0, None
    lowerv = v.lower()
    for b in correlatingTable:
        strval = b.lower()
        localmatched = 0
        for i in range(len(lowerv)):
            if strval[i] and lowerv[i] and strval[i] == lowerv[i]:
                localmatched += 1
            else:
                if localmatched > charactersMatched:
                    charactersMatched = localmatched
                    results = b
                else:
                    continue
            if localmatched == len(lowerv):
                charactersMatched = localmatched
                results = b
    return results
def beautifulPrint(tableToPrint):
    print("{")
    for v in tableToPrint:
        print(" {v}".format(v=v))
    print("}")
def getLines(f):
    return f.readline().strip(), f.readline().strip(), f.readline().strip(), int(f.readline().strip()), f.readline().strip(), int(f.readline().strip())

datatables = {
    "Types":["Physical","Digital"],
    "Subject":["Mathematics","Further Mathematics","Physics"],
    "Alphabet":["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","Z","Y","Z"],
    "Decision":["Yes","No"],
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
                searchResults = matchToTables(input("Value to find?: "), sheet1["Topic"], True)
                responseToValue = matchToTables(input("Is {val} the value that you wanted to find?: ".format(val=searchResults)), datatables["Decision"])
                if responseToValue == "Yes":
                    print("Search result: {val}".format(val=searchResults)); input()
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
                    if sheet1.iloc[data[1]]["Re-review"]:
                        data[0] *= sheet1.iloc[data[1]]["Re-review"]
                if len(subjectsOfReviewInterest) > 0:
                    print("Areas of improvement: ")
                    beautifulPrint(subjectsOfReviewInterest)
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
