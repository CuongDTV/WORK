import xlwings as xw
import numpy as np
import csv
import math
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path    

def Nbrebar(width):
    widths = [150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350, 1400, 1450, 1500, 1550, 1600, 1650, 1700, 1750, 1800, 1850, 1900, 1950, 2000, 2050, 2100, 2150]
    Nbrebars = [2, 2, 2, 3, 3, 4, 4, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14, 15, 15, 16, 16, 17, 17, 18, 18, 19, 19]
    return Nbrebars[widths.index(width)]

def CalAsProvide(AsRebar):
    AsProvide = 0
    for i in range(0,4):
        AsProvide += AsRebar[2*i] * math.pi * AsRebar[2*i+1]**2/4
    return AsProvide

def ReduceRebar(AsRebar):
    Rebar_list=[20, 25, 32, 40]
    for i in range(3,-1,-1):
        if AsRebar[i] > 20:
            AsRebar[i] = Rebar_list[Rebar_list.index(AsRebar[i])-1]
            return AsRebar

def IncreaseRebar(AsRebar,limit):
    Rebar_list=[0, 20, 25, 32, 40]     
    for i in range(0,4,1):
        if AsRebar[i] < limit:
            AsRebar[i] = Rebar_list[Rebar_list.index(AsRebar[i])+1]
            return AsRebar

def Legs(width):
    thresholds = [400, 700, 900, 1100, 1400, 1600, 1800, 2000]
    leg_counts = [2, 4, 6, 8, 10, 12, 14, 16, 18]

    for i in range(len(thresholds)):
        if width < thresholds[i]:
            return leg_counts[i]

    return leg_counts[8]

def DiaAndSpacing(asreq,legs):
    dias = [10, 12, 16, 20]
    s = 0 
    i = -1
    while s < 100:
        i += 1
        s = int((min(250, (legs * dias[i] ** 2 * math.pi) / (4 * asreq * 1.1) * 1000, dias[i] * 20) // 50) * 50)    
    return [dias[i], s]

def Spacing(s, location, suppcond, rebardia):

    if location in (1, 3) and suppcond in (1, 3, 4):
        spacing = min(s, 150, rebardia * 8)
    else:
        spacing = min(s, rebardia * 12)

    spacing = int((spacing // 50) * 50)
    return spacing

def BackUpRebar(wb):
    worksheet_names = ["REBAR SADS DATA", "BACKUP REBAR"]
    worksheets = {name: wb.sheets[name] for name in worksheet_names}

    beamVals = worksheets["REBAR SADS DATA"].range("A3:F{}".format(worksheets["REBAR SADS DATA"].api.UsedRange.Rows.Count+1)).value
    rebarVals = worksheets["REBAR SADS DATA"].range("Y3:BT{}".format(worksheets["REBAR SADS DATA"].api.UsedRange.Rows.Count+1)).value
    beamVals = [row for row in beamVals if not all(cell is None for cell in row)]
    rebarVals = [row for row in rebarVals if not all(cell is None for cell in row)]

    sheet = worksheets["BACKUP REBAR"]
    sheet.range("3:{}".format(sheet.api.UsedRange.Rows.Count+1)).clear()
    sheet.range("A3").value = beamVals
    sheet.range("G3").value = rebarVals
    num_rows = len(beamVals)
    num_columns = 54
    data_range = sheet.range(3, 1).expand("table").resize(num_rows, num_columns)

    sheet.range("C3:C{}".format(num_rows+2)).api.Font.Bold = True
    sheet.range("C3:C{}".format(num_rows+2)).api.Font.Color = xw.constants.RgbColor.rgbRed
    data_range.api.Borders.LineStyle = -4142
    data_range.api.Borders.ColorIndex = 0
    data_range.api.Borders.Weight = 2
    sheet.range("G3:BB{}".format(num_rows+2)).api.Font.Bold = True
    sheet.range("G3:BB{}".format(num_rows+2)).api.Interior.Pattern = -4142
    sheet.range("G3:BB{}".format(num_rows+2)).api.Interior.Color = int("B7B8E6", 16)

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
            output_vals.append((beam_dict[beam_key], supp_vals[row_index][4], supp_vals[row_index + 1][4]))
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

    file_names = ["BEAMBARS.TXT", "BARINFO.TXT", "BEAMDATA.TXT", "BTCHDATA.TXT", "BTCHBEAM.TXT", "BTCHSUPP.TXT", "STIRINFO.TXT"]
    sheet_names = ["BEAMBARS", "BARINFO", "BEAMDATA", "BTCHDATA", "BTCHBEAM", "BTCHSUPP", "STIRINFO"]

    for file_name, sheet_name in zip(file_names, sheet_names):
        file_path = folder_path / file_name
        if file_path.is_file():
            sheet = wb.sheets[sheet_name]
            sheet.clear()
            with file_path.open('r') as f:
                reader = csv.reader(f, delimiter="\t")
                data = [[cell.replace(' ', '" "') for cell in row] for row in reader]
                sheet.range("A2").value = data

def RebarSadsData(wb,allowance):
    worksheet_names = ["BTCHBEAM", "REBAR SADS DATA", "BEAMDATA", "BTCHSUPP", "BEAMBARS", "BARINFO", "BACKUP REBAR"]
    worksheets = {name: wb.sheets[name] for name in worksheet_names}
    
    global beamDict
    beamDict = {}
    # Add data from worksheet "BTCHBEAM"
    beamVals = worksheets["BTCHBEAM"].range("A2:D{}".format(worksheets["BTCHBEAM"].api.UsedRange.Rows.Count+1)).value
    beamDict = {str(beamVals[i][3]): [beamVals[i][0]] for i in range(len(beamVals)) if str(beamVals[i][3]) not in beamDict}
    # Add data from worksheet "BEAMDATA"
    beamVals = worksheets["BEAMDATA"].range("A2:G{}".format(worksheets["BEAMDATA"].api.UsedRange.Rows.Count+1)).value
    beamDict = {beamname: beamValues + [beamVals[i][5], beamVals[i][6], beamVals[i][2]] for i in range(len(beamVals)) for beamname, beamValues in beamDict.items() if beamname == str(beamVals[i][0])}
    # Add data from worksheet "BTCHSUPP"
    beamVals = worksheets["BTCHSUPP"].range("M2:O{}".format(worksheets["BTCHSUPP"].api.UsedRange.Rows.Count+1)).value
    beamDict = {beamname: beamValues + [beamVals[i][1], beamVals[i][2]] for i in range(len(beamVals)) for beamname, beamValues in beamDict.items() if beamname == str(beamVals[i][0])}
    # Add data from worksheet "BEAMBARS"
    beamVals = worksheets["BEAMBARS"].range("A2:W{}".format(worksheets["BEAMBARS"].api.UsedRange.Rows.Count+1)).value
    beamDict = {beamname: beamValues + [
        max(beamVals[i][11], beamVals[i][20])*allowance,
        max(beamVals[i][12], beamVals[i][21])*allowance,
        max(beamVals[i][13], beamVals[i][22])*allowance,
        max(beamVals[i][14], beamVals[i][17])*allowance,
        max(beamVals[i][15], beamVals[i][18])*allowance,
        max(beamVals[i][16], beamVals[i][19])*allowance
    ] for i in range(len(beamVals)) for beamname, beamValues in beamDict.items() if beamname == str(beamVals[i][0])}
    # Add data from worksheet "BARINFO"
    beamVals = worksheets["BARINFO"].range("A2:F{}".format(worksheets["BARINFO"].api.UsedRange.Rows.Count+1)).value
    rebar={}
    for beamname in list(beamDict.keys()):
        rebar[beamname] = [0]*48
    for i in range(len(beamVals)):
        if beamVals[i][0] in rebar:
            rebar[beamVals[i][0]][int((beamVals[i][3]-1)*8+(beamVals[i][2]-1)*2)]=beamVals[i][5]
            rebar[beamVals[i][0]][int((beamVals[i][3]-1)*8+(beamVals[i][2]-1)*2+1)]=beamVals[i][4]
    beamDict = {beamname: beamValues + rebar[beamname] for beamname, beamValues in beamDict.items()}

    # Fix rebar
    for index, beamname in enumerate(list(beamDict.keys()), start=1):
        fixrebar = beamDict[beamname]
        #Fix number rebar
        for i in range(1,25):
            fixrebar[10+2*i]=Nbrebar(fixrebar[1]) if fixrebar[11+2*i] != 0 else fixrebar[10+2*i]
        #Check left rebar
        if fixrebar[4] in (1,3):
            #Check top left rebar
            while CalAsProvide(fixrebar[12:20])/fixrebar[1]/fixrebar[2]*100 > 2.5:
                reduce=ReduceRebar([fixrebar[13],fixrebar[15],fixrebar[17],fixrebar[19]])
                fixrebar[13]=reduce[0]
                fixrebar[15]=reduce[1]
                fixrebar[17]=reduce[2]
                fixrebar[19]=reduce[3]
            #Check bottom left rebar
            while CalAsProvide(fixrebar[12:20])>2*CalAsProvide(fixrebar[36:44]):
                increase=IncreaseRebar([fixrebar[37],fixrebar[39],fixrebar[41],fixrebar[43]],fixrebar[13])
                fixrebar[37]=increase[0]
                fixrebar[39]=increase[1]
                fixrebar[38]=fixrebar[36] if fixrebar[39]!=0 else fixrebar[38]
                fixrebar[41]=increase[2]
                fixrebar[40]=fixrebar[36] if fixrebar[41]!=0 else fixrebar[40]
                fixrebar[43]=increase[3]
                fixrebar[42]=fixrebar[36] if fixrebar[43]!=0 else fixrebar[42]
        #Check right rebar
        if fixrebar[5] in (1,3):
            #Check top right rebar
            while CalAsProvide(fixrebar[28:36])/fixrebar[1]/fixrebar[2]*100 > 2.5:
                reduce=ReduceRebar([fixrebar[29],fixrebar[31],fixrebar[33],fixrebar[35]])
                fixrebar[29]=reduce[0]
                fixrebar[31]=reduce[1]
                fixrebar[33]=reduce[2]
                fixrebar[35]=reduce[3]      
            #Check bottom right rebar
            while CalAsProvide(fixrebar[28:36])>2*CalAsProvide(fixrebar[52:60]):
                increase=IncreaseRebar([fixrebar[53],fixrebar[55],fixrebar[57],fixrebar[59]],fixrebar[29])
                fixrebar[53]=increase[0]
                fixrebar[55]=increase[1]
                fixrebar[54]=fixrebar[52] if fixrebar[55]!=0 else fixrebar[54]
                fixrebar[57]=increase[2]
                fixrebar[56]=fixrebar[52] if fixrebar[57]!=0 else fixrebar[56]
                fixrebar[59]=increase[3]  
                fixrebar[58]=fixrebar[52] if fixrebar[59]!=0 else fixrebar[58]
        #Fix top middle rebar
        fixrebar[21]=max(min(fixrebar[13],fixrebar[29]),fixrebar[21])
        fixrebar[23]=max(fixrebar[15],fixrebar[31],fixrebar[23]) if fixrebar[23] !=0 else fixrebar[23]
        fixrebar[25]=max(fixrebar[17],fixrebar[33],fixrebar[25]) if fixrebar[25] !=0 else fixrebar[25]
        fixrebar[27]=max(fixrebar[19],fixrebar[35],fixrebar[27]) if fixrebar[27] !=0 else fixrebar[27]
        #Fix bottom middle rebar
        fixrebar[45]=max(min(fixrebar[37],fixrebar[53]),fixrebar[45])
        #Fix layer 1 rebar
        fixrebar[13] = fixrebar[21] = fixrebar[29] = max(fixrebar[13], fixrebar[21], fixrebar[29])
        fixrebar[37] = fixrebar[45] = fixrebar[53] = max(fixrebar[37], fixrebar[45], fixrebar[53])
    worksheets["REBAR SADS DATA"].range("3:{}".format(worksheets["REBAR SADS DATA"].api.UsedRange.Rows.Count+1)).clear()
    sheet = worksheets["REBAR SADS DATA"]
    data_to_write = []

    for index, (beamname, beamValues) in enumerate(beamDict.items(), start=1):
        data_row = [
            index,
            beamValues[0],
            beamname,
            *beamValues[1:9],
            '=IF(C{0}="", "", Y{0}*PI()*Z{0}^2/4+AA{0}*PI()*AB{0}^2/4+AC{0}*PI()*AD{0}^2/4+AE{0}*PI()*AF{0}^2/4)'.format(index+2),
            '=IF(C{0}="", "", AG{0}*PI()*AH{0}^2/4+AI{0}*PI()*AJ{0}^2/4+AK{0}*PI()*AL{0}^2/4+AM{0}*PI()*AN{0}^2/4)'.format(index+2),
            '=IF(C{0}="", "", AO{0}*PI()*AP{0}^2/4+AQ{0}*PI()*AR{0}^2/4+AS{0}*PI()*AT{0}^2/4+AU{0}*PI()*AV{0}^2/4)'.format(index+2),
            *beamValues[9:12],
            '=IF(C{0}="", "", AW{0}*PI()*AX{0}^2/4+AY{0}*PI()*AZ{0}^2/4+BA{0}*PI()*BB{0}^2/4+BC{0}*PI()*BD{0}^2/4)'.format(index+2),
            '=IF(C{0}="", "", BE{0}*PI()*BF{0}^2/4+BG{0}*PI()*BH{0}^2/4+BI{0}*PI()*BJ{0}^2/4+BK{0}*PI()*BL{0}^2/4)'.format(index+2),
            '=IF(C{0}="", "", BM{0}*PI()*BN{0}^2/4+BO{0}*PI()*BP{0}^2/4+BQ{0}*PI()*BR{0}^2/4+BS{0}*PI()*BT{0}^2/4)'.format(index+2),
            '=IF(C{0}="", "", IF(AND(L{0}>=I{0},M{0}>=J{0},N{0}>=K{0}),"OK","CHECK REBAR"))'.format(index+2),
            '=IF(C{0}="", "", IF(AND(R{0}>=O{0},S{0}>=P{0},T{0}>=Q{0}),"OK","CHECK REBAR"))'.format(index+2),
            '=IF(C{0}="", "", IF(OR(G{0}=1,G{0}=3),IF(AND(L{0}/D{0}/E{0}*100<=2.5,L{0}<=2*R{0}),"OK","FAILED"),""))'.format(index+2),
            '=IF(C{0}="", "", IF(OR(H{0}=1,H{0}=3),IF(AND(N{0}/D{0}/E{0}*100<=2.5,N{0}<=2*T{0}),"OK","FAILED"),""))'.format(index+2),
            *beamValues[12:60],
        ]
        data_to_write.append(data_row)

    # Bulk write the data to Excel
    num_rows = len(data_to_write)
    num_columns = len(data_to_write[0])
    data_range = sheet.range(3, 1).expand("table").resize(num_rows, num_columns)
    data_range.value = data_to_write
    sheet.range("C3:C{}".format(num_rows+2)).api.Font.Bold = True
    sheet.range("C3:C{}".format(num_rows+2)).api.Font.Color = xw.constants.RgbColor.rgbRed
    data_range.api.Borders.LineStyle = -4142
    data_range.api.Borders.ColorIndex = 0
    data_range.api.Borders.Weight = 2
    sheet.range("Y3:BT{}".format(num_rows+2)).api.Font.Bold = True
    sheet.range("Y3:BT{}".format(num_rows+2)).api.Interior.Pattern = -4142
    sheet.range("Y3:BT{}".format(num_rows+2)).api.Interior.Color = int("B7B8E6", 16)

def ImportBackup(wb):
    worksheet_names = ["REBAR SADS DATA", "BACKUP REBAR"]
    worksheets = {name: wb.sheets[name] for name in worksheet_names}

    backupVals = worksheets["BACKUP REBAR"].range("C3:BB{}".format(worksheets["BACKUP REBAR"].api.UsedRange.Rows.Count + 1)).value
    backupVals = [row for row in backupVals if not all(cell is None for cell in row)]
    
    beamVals = worksheets["REBAR SADS DATA"].range("C3:F{}".format(worksheets["REBAR SADS DATA"].api.UsedRange.Rows.Count + 1)).value
    beamVals = [row for row in beamVals if not all(cell is None for cell in row)]

    beamVals_np = np.array(beamVals)
    backupVals_np = np.array(backupVals)

    beamVals_id = beamVals_np[:, 0:4]
    backupVals_id = backupVals_np[:, 0:4]

    matching_indices = np.where(np.all(beamVals_id[:, None] == backupVals_id, axis=2))
    if len(matching_indices[0]) == 0:
        return
    update_data_list = []
    begin = matching_indices[0][0]
    end = matching_indices[0][0]-1
    for index in matching_indices[0]:
        if end == index - 1:
            end = index
        else:
            update_data_array = np.array(update_data_list)
            worksheets["REBAR SADS DATA"].range("Y{0}:BT{1}".format(begin+3,end+3)).value = update_data_array
            worksheets["REBAR SADS DATA"].range("Y{0}:BT{1}".format(begin+3,end+3)).api.Interior.Color = int("E2B48D", 16)
            begin = index
            end = index
            update_data_list = []
        update_data = backupVals_np[index, 4:52]
        update_data_list.append(update_data)
        if end == matching_indices[0][-1]:
            update_data_array = np.array(update_data_list)
            worksheets["REBAR SADS DATA"].range("Y{0}:BT{1}".format(begin+3,end+3)).value = update_data_array
            worksheets["REBAR SADS DATA"].range("Y{0}:BT{1}".format(begin+3,end+3)).api.Interior.Color = int("E2B48D", 16)

def StirrupProvide(wb):
    worksheet_names = ["STIRINFO", "STIRRUP PROVIDE"]
    worksheets = {name: wb.sheets[name] for name in worksheet_names}  

    global beamDict

    # Add data from worksheet "STIRINFO"
    beamVals = worksheets["STIRINFO"].range("A2:H{}".format(worksheets["STIRINFO"].api.UsedRange.Rows.Count+1)).value
    beamDict = {beamname: beamValues + [
        beamVals[i][7] if beamVals[i][7] != 0 else beamVals[i+1][7],
        beamVals[i+1][7],
        beamVals[i+2][7] if beamVals[i+2][7] != 0 else beamVals[i+1][7],
        DiaAndSpacing(max(beamVals[i][7],beamVals[i+1][7],beamVals[i+2][7]),Legs(beamValues[1]))[0]
    ] for i in range(0, len(beamVals), 3) for beamname, beamValues in beamDict.items() if beamname == str(beamVals[i][0])}
    
    worksheets["STIRRUP PROVIDE"].range("3:{}".format(worksheets["STIRRUP PROVIDE"].api.UsedRange.Rows.Count+1)).clear()
    sheet = worksheets["STIRRUP PROVIDE"]
    data_to_write = []
    for index, (beamname, beamValues) in enumerate(beamDict.items(), start=0):
        legs_value = Legs(beamValues[1])
        for i in range(1, 4):
            supp = beamValues[4] if i == 1 else beamValues[5] if i == 3 else 0
            RebarDia = min(beamValues[13+(i-1)*8],beamValues[37+(i-1)*8])
            s = int((min(250, (legs_value * beamValues[63] ** 2 * math.pi) / (4 * beamValues[59+i] * 1.1) * 1000, beamValues[63] * 20) // 50) * 50)
            data_row = [
                beamname,
                beamValues[1],
                beamValues[3],
                supp,
                RebarDia,
                i,
                beamValues[63],
                legs_value,
                Spacing(s, i, 4, RebarDia) if (beamValues[4]==4) or (beamValues[5]==4) else Spacing(s, i, supp, RebarDia),
                0,
                beamValues[59+i],
                '=IF(J{0}<100,"FAILED","OK")'.format(index*3+i+2)
            ]
            data_to_write.append(data_row)
        if beamValues[3] <= 2:
            min_value = min(data_to_write[index*3][8], data_to_write[index*3+1][8], data_to_write[index*3+2][8])
            data_to_write[index*3][9] = min_value
            data_to_write[index*3+1][9] = min_value
            data_to_write[index*3+2][9] = min_value
        else:
            data_to_write[index*3][9] = data_to_write[index*3][8]
            data_to_write[index*3+1][9] = data_to_write[index*3+1][8]
            data_to_write[index*3+2][9] = data_to_write[index*3+2][8]         

    # Bulk write the data to Excel
    num_rows = len(data_to_write)
    num_columns = len(data_to_write[0])
    data_range = sheet.range(3, 1).expand("table").resize(num_rows, num_columns)
    data_range.value = data_to_write
    data_range.api.Borders.LineStyle = -4142
    data_range.api.Borders.ColorIndex = 0
    data_range.api.Borders.Weight = 2
    sheet.range("G3:G{}".format(num_rows+2)).api.Font.Bold = True
    sheet.range("G3:G{}".format(num_rows+2)).api.Interior.Pattern = -4142
    sheet.range("G3:G{}".format(num_rows+2)).api.Interior.Color = int("00FFFF", 16)
       
def main():
    wb=xw.books.active
    result = messagebox.askquestion("Backup Confirmation", "Do you want to backup the REBAR SADS DATA sheet before importing the new data?", icon='question')
    if result == 'yes':
        BackUpRebar(wb)
    ImportFile(wb)
    CheckSupport(wb)
    allowance=1.05
    RebarSadsData(wb,allowance)
    ImportBackup(wb)
    StirrupProvide(wb)

if __name__=="__main__":
    main() 