import xlwings as xw
import csv
import math
import tkinter as tk
from tkinter import filedialog
from pathlib import Path    
   
def CheckSupport(wb):
    btchbeam_sheet = wb.sheets["BTCHBEAM"]
    btchsupp_sheet = wb.sheets["BTCHSUPP"]

    beam_dict = {}
    beam_vals = btchbeam_sheet.range("A2:D{}".format(btchbeam_sheet.api.UsedRange.Rows.Count+1)).value
    beam_dict = {str(row[0]) + str(row[1]) + str(row[2]): row[3] for row in beam_vals}

    supp_vals = btchsupp_sheet.range("A2:E{}".format(btchsupp_sheet.api.UsedRange.Rows.Count+1)).value

    output_vals = []
    for row_index in range(len(supp_vals)):
        beam_key = str(supp_vals[row_index][0]) + str(supp_vals[row_index][1]) + str(supp_vals[row_index][2])
        if beam_key in beam_dict:
            output_vals.append((beam_dict[beam_key], supp_vals[row_index][3], supp_vals[row_index + 1][3]))
        else:
            output_vals.append(("","",""))

    btchsupp_sheet.range("M2").value = output_vals

def ImportFile(wb):
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select a folder")
    if not folder_path:
        print("No folder selected!")
        exit()
    folder_path=Path(folder_path)

    file_names = ["BEAMDATA.TXT", "BTCHBEAM.TXT", "BTCHSUPP.TXT", "COLMDATA.TXT", "COLMBEAM.TXT"]
    sheet_names = ["BEAMDATA", "BTCHBEAM", "BTCHSUPP", "COLMDATA", "COLMBEAM"]

    for file_name, sheet_name in zip(file_names, sheet_names):
        file_path = folder_path / file_name
        if file_path.is_file():
            sheet = wb.sheets[sheet_name]
            sheet.clear()
            with file_path.open('r') as f:
                reader = csv.reader(f, delimiter="\t")
                data = [[cell.replace(' ', '" "') for cell in row] for row in reader]
                sheet.range("A2").value = data

def Length(wb):
    worksheet_names = ["BEAMDATA", "BTCHSUPP", "REVITDATA", "COLMDATA", "COLMBEAM", "SADS"]
    worksheets = {}
    for name in worksheet_names:
        worksheets[name] = wb.sheets[name]

    beamDict = {}
    # Add data from worksheet "BEAMDATA"
    beamVals = worksheets["BEAMDATA"].range("A2:G{}".format(worksheets["BEAMDATA"].api.UsedRange.Rows.Count+1)).value
    beamDict = {str(beamVals[i][0]): [beamVals[i][5], beamVals[i][6]] for i in range(len(beamVals)) if str(beamVals[i][0]) not in beamDict} 
    # Add data from worksheet "BTCHSUPP"
    beamVals = worksheets["BTCHSUPP"].range("M2:O{}".format(worksheets["BTCHSUPP"].api.UsedRange.Rows.Count+1)).value
    beamDict = {beamname: beamValues + [beamVals[i][1], beamVals[i][2]] for i in range(len(beamVals)) for beamname, beamValues in beamDict.items() if beamname == str(beamVals[i][0])}
    # Add data from worksheet "REVIT"
    beamVals = worksheets["REVITDATA"].range("C3:G{}".format(worksheets["REVITDATA"].api.UsedRange.Rows.Count+1)).value
    beamVals = [['C'+str(beamVal[0])] + beamVal[1:] if beamVal[4] == 'Yes' else beamVal for beamVal in beamVals]
    beamDict = {beamname: beamValues + [beamVals[i][1], beamVals[i][2], beamVals[i][3]] for i in range(len(beamVals)) for beamname, beamValues in beamDict.items() if beamname == str(beamVals[i][0])}

    colmDict = {}
    # Add data from worksheet "COLMDATA"
    colmVals = worksheets["COLMDATA"].range("A2:E{}".format(worksheets["COLMDATA"].api.UsedRange.Rows.Count+1)).value
    colmDict = {str(colmVals[i][0]) + str(colmVals[i][1]): [colmVals[i][3], colmVals[i][4]] for i in range(len(colmVals)) if str(colmVals[i][0]) + str(colmVals[i][1]) not in colmDict} 
    # Add data from worksheet "COLMBEAM"
    colmVals = worksheets["COLMBEAM"].range("A2:F{}".format(worksheets["COLMBEAM"].api.UsedRange.Rows.Count+1)).value
    for i in range(len(colmVals)):
        if str(colmVals[i][3]) in beamDict:
            beamValues = beamDict[colmVals[i][3]]
            if colmVals[i][5]<45 or 135<colmVals[i][5]<225 or colmVals[i][5]>315:
                width = colmDict[str(colmVals[i][0])+str(colmVals[i][1])][0]
            else:
                width = colmDict[str(colmVals[i][0])+str(colmVals[i][1])][1]
            beamDict[colmVals[i][3]] = beamValues + [width]

    worksheets["SADS"].range("3:{}".format(worksheets["SADS"].api.UsedRange.Rows.Count+1)).clear()
    sheet = worksheets["SADS"]
    for beamname, beamValues in beamDict.items():
        supp = [0,0]
        k = 0
        for i in range(1,3):
            if beamValues[i+1] is None:
                supp[i-1] = 0
            else:
                if str(beamValues[i+1]) not in beamDict:
                    if beamValues[i+1][0]=='C':
                        k += 1
                        supp[i-1] = beamValues[k+6]
                    else:
                        supp[i-1] = 400
                else:
                    supp[i-1] = beamDict[beamValues[i+1]][0]
        beamDict[beamname] = beamValues + supp

    data_to_write = []
    for index, (beamname, beamValues) in enumerate(beamDict.items(), start=3):
        data_row = [
            beamname,
            '=FLOOR((C{0}+D{0}/2+E{0}/2)/1000,0.05)'.format(index),
            beamValues[4],
            beamValues[len(beamValues)-2],
            beamValues[len(beamValues)-1],
            beamValues[0],
            beamValues[5],
            beamValues[1],
            beamValues[6],
            '=IF(AND(F{0}=G{0},H{0}=I{0}),"OK","FALSE")'.format(index)
        ]
        data_to_write.append(data_row)
    num_rows = len(data_to_write)
    num_columns = len(data_to_write[0])
    data_range = sheet.range(3, 1).expand("table").resize(num_rows, num_columns)
    data_range.value = data_to_write
    data_range.api.Borders.LineStyle = -4142
    data_range.api.Borders.ColorIndex = 0
    data_range.api.Borders.Weight = 2
    sheet.range("B3:B{}".format(num_rows+2)).api.Font.Bold = True

def main():
    wb=xw.books.active
    ImportFile(wb)
    CheckSupport(wb)
    Length(wb)

if __name__=="__main__":
    main() 