"""
Reads entries from a spreadsheet, picks the appropriate zip file for each
entry an attempts to extract files containing the name in the spreadsheet
before renaming the file to maintain the spreadsheet's order
"""

import xlrd
import zipfile
from os import rename, remove

pathFolder = "..\\"
# Open spreadsheet file
workbook = xlrd.open_workbook('zzzChallengeSubmissions.xls')
gameSheet = workbook.sheet_by_name('Challenge Submissions')
i=1
failList=[]
# # Iterate through every game in the spreadsheet
for x in range(1, (gameSheet.nrows)):
    # Identify which zip contains the game
    packName = gameSheet.cell(x,3).value.replace('/', '')+'.zip'
    with zipfile.ZipFile(packName) as romPack:
        # Extract all files in the cell that contain the name of the game in the cell
        targetGame = [j for j in romPack.namelist() if gameSheet.cell(x,2).value.lower().replace(":"," -") in j.lower()]
        if len(targetGame) == 0:
            failList.append(i, gameSheet.cell(x,2).value)
        else:
            extractPath = romPack.extract(targetGame[0], pathFolder)
            try:
                rename(extractPath, (pathFolder + "{:03d} ".format(i) + targetGame[0]))
            except FileExistsError:
                print(j + "already in folder")
                remove(extractPath)

    i+=1
if len(failList) > 0:
    print("{} games not found:".format(len(failList)))
    for k in failList:
        print(k)
