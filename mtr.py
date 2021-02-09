import pandas as pd

df = pd.read_excel(r'C:\Users\jberg\OneDrive - A-T Controls, Inc\pythonMTR\mtrinput.xlsx')

size = []
materiallist = []
material = []
partno = []
component = []
endstyle = []
endoptions = ["BW", "DA", "F1", "F3", "F6", "L1", "L3", "LUG", "TH", "SA", "SF", "SO", "SW", "WAFER"]

# adding pdf file extension to filename column
df["FileName(exact).pdf"] = df["FileName(exact).pdf"] + ".pdf"

# extacting the partnumber and component from the format Mars provides
for line in df["A-TPartNo."]:
    partno.append(line[line.find('(') + 1:line.find(')')])
    component.append(line[line.find(' ') + 1:line.find('(')])

component = [i.upper() for i in component]
partno = [i.upper() for i in partno]


# fixing formatting errors in the component column
for i, line in enumerate(component):
    if component[i] == 'BLIND END':
        component[i] = 'BLIND'
    if component[i] == 'TSM TOP':
        component[i] = 'TOP/TSM'
    if component[i] == 'TSM DOWN':
        component[i] = 'DOWN/TSM'


# extracting size from the part number field
size = df["A-TPartNo."].str.split(' ').str.get(0)

# extracting the heat number to be enumerated in next step
for line in df["HeatNo."]:
    materiallist.append(line)
materiallist = [str(i) for i in materiallist]
materiallist = [i.upper() for i in materiallist]

# determining the material from the heat number
for number, line in enumerate(materiallist):
    if materiallist[number][-1] == "W":
        material.append("WCB")
    elif materiallist[number][-1] == "S":
        material.append("CF8M")
    elif materiallist[number][-1] == "L":
        material.append("CF3M")
    else:
        material.append("")


# writing the endstyle column from option set up in variable endoptions
for i, line in enumerate(partno):
    endstyle.append("")
    for end in endoptions:
        if end in line:
            endstyle[i] = end


# fixing formatting errors in the size column
for i, line in enumerate(size):
    if size[i] == '11/2"':
        size[i] = '1-1/2"'
    elif size[i] == '11/4"':
        size[i] = '1-1/4"'
    elif size[i] == '21/2"':
        size[i] = '2-1/2"'


# saving values back to df
df["Size"] = size
df["A-TPartNo."] = partno
df["Component"] = component
df["EndStyle"] = endstyle
df["Material"] = material

# output to csv file
df.to_csv(r'C:\Users\jberg\OneDrive - A-T Controls, Inc\pythonMTR\output.csv', index=False)


df2 = pd.read_csv(r'C:\Users\jberg\OneDrive - A-T Controls, Inc\pythonMTR\items.csv', usecols=['Name'], encoding='ISO-8859-1')

missing = []

items = []

# Function takes in 2 lists and compares each element from the first provided list to to second list.
def non_match_elements(list_a, list_b):
    non_match = []
    for i in list_a:
        if i not in list_b:
            non_match.append(i)
    return non_match


for i in df2['Name']:
    items.append(i)

missing = non_match_elements(partno, items)

if len(missing) > 0:
    df3 = pd.DataFrame (missing, columns=['Missing Part Number'])

    df3.to_csv(r'C:\Users\jberg\OneDrive - A-T Controls, Inc\pythonMTR\missing.csv', index=False)