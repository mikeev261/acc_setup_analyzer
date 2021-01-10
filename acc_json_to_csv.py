import sys
import json
import csv
import xlwt
import pandas as pd
from datetime import datetime

import code


print("Hello world")

file_setup_json_0 = sys.argv[1]
file_setup_json_1 = sys.argv[2]

print("First file: " + str(file_setup_json_0))

#STEP 0: Import JSON 

#data = json.load(file_setup_json_0)

df_setup_0 = pd.read_json(file_setup_json_0)

#print(df_setup_0['tyres'])
#file_setup_json_0.close()
#file_setup_json_1.close()

#STEP 1: Define XLS Styles
style0 = xlwt.easyxf('font: name Arial, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

style_header = xlwt.easyxf('font: name Arial, color-index blue, bold on; align: vert centre, horiz centre', 
    num_format_str='#,##0.00')


style_super_header = xlwt.XFStyle()
#Basic black background
black_pattern = xlwt.Pattern()
black_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
black_pattern.pattern_fore_colour = xlwt.Style.colour_map['blue']
style_super_header.pattern = black_pattern




#STEP 2: Create XLS Workbook
wb = xlwt.Workbook()

#--Sheets
ws_tyres = wb.add_sheet('Tyres')
ws_electronics = wb.add_sheet('Electronics')
ws_strategy = wb.add_sheet('Strategy')
ws_mechanical = wb.add_sheet('Mechanical Balance')
ws_aero = wb.add_sheet('Aerodynamics')

#Step 3: Populate these sheets
#--3a: TYRES
col_tyres_setups = ws_tyres.col(0)
col_tyres_setups.width = 300*20
car_name = "Car: " + str(df_setup_0['carName'][0])
#----Car Name
ws_tyres.write(0,0,car_name, style0)
#----Headers
ws_tyres.write(2,0,"Setup Name", style_header)
ws_tyres.write(2,1,"Compound", style_header)

#Take the cells to merge as r1, r2, c1, c2, and accept an optional style parameter.
ws_tyres.write_merge(1, 1, 2, 5, '----Pressures (PSI)----', style_super_header)

ws_tyres.write(2,2,"FL", style_header)
ws_tyres.write(2,3,"FR", style_header)
ws_tyres.write(2,4,"RL", style_header)
ws_tyres.write(2,5,"RL", style_header)

#----Data
ws_tyres.write(3,0, file_setup_json_0)
ws_tyres.write(4,0, file_setup_json_1)


ws_electronics.write(1, 0, datetime.now(), style1)
ws_strategy.write(2, 0, 1)
ws_mechanical.write(2, 1, 1)
ws_aero.write(2, 2, xlwt.Formula("A3+B3"))

wb.save('example.xls')


print(car_name)




#Enable this to keep interactive terminal open
#code.interact(banner="Start", local=locals(), exitmsg="End")
