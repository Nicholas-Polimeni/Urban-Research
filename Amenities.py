import pandas as pd
import numpy
import os, psutil
import xlsxwriter

gc.enable()

# Goes through the "amentities" column and creates an array of distinct strings,
# representing every type of amentity in the data. The labels in this array
# will be used as the column headers. 
def makeColHeaders(prop):
    try:
        prop = prop[0].split(", ")
        for amen in prop:
            if amen not in colHeaders:
                colHeaders.append(amen)
    except:
        pass

# Goes through the "amentities" column for each individual property listing,
# looking to see if there are any amentities. For each amentity the property has,
# a 1 is added in the binary array being made, under the column labeled for that
# amenity.
def fillBinary(propAmenities):
    binaryResult = []
    try:
        prop = propAmenities[0].split(", ")
    except:
        pass
    
    for header in colHeaders:
        binaryResult.append(0)
        try:
            for amen in prop:
                if amen.lower() == header.lower():
                    binaryResult[len(binaryResult)-1] = 1
        except:
            pass
    return binaryResult

process = psutil.Process(os.getpid())
print(process.memory_info().rss)
print("Started... Please Wait...")
amenitiesDataFrame = pd.read_excel('ALL DATA!!.xlsx', usecols=["Amenities"])
print("Made DataFrame")
amenitiesArray = amenitiesDataFrame.to_numpy()
colHeaders = []

print("makeColHeaders")
for prop in amenitiesArray:
    makeColHeaders(prop)

binaryArray = []

print("binaryArray")
for prop in amenitiesArray:
    binaryArray.append(fillBinary(prop))

print("Final Array")
finalDataFrame = pd.DataFrame(binaryArray, columns=colHeaders)
finalDataFrame.index.name = "id"

print("Writing")
headers = ["id"] + finalDataFrame.columns.tolist()
binaryVals = finalDataFrame.values.tolist()

for i in range(len(binaryVals)):
    binaryVals[i] = [i + 1] + binaryVals[i]

writeVals = [headers] + binaryVals

# Writes data using xlsxwriter. I've tried ExcelWriter, pycelerate, xlsxwriter
# with constant_memory set to true and set to false. This is where the problem
# with memory occurs.
workbook = xlsxwriter.Workbook('Binary Amenities.xlsx', {'constant_memory': True})
worksheet = workbook.add_worksheet()

for rowNum, rowData in enumerate(writeVals):
    for colNum, colData in enumerate(rowData):
        worksheet.write(rowNum, colNum, colData)
        
workbook.close()

print("Done!")