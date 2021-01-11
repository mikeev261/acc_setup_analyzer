import sys
import json
import csv
import xlwt
import pandas as pd
from datetime import datetime

import code


print("Hello world")

#GLOBAL CONSTANTS
ROW_OFFSET = 3

#file_setup_json_0 = sys.argv[1]
#file_setup_json_1 = sys.argv[2]

setup_names = []

n = 1 #Instantiate our argv number
more_args = True #Instantiate bool for more args still existing (2 or more)
df_setup = {} #Instantiate dict of dataframes for setups
while more_args:
    try:
        df_setup[n-1] = pd.read_json(sys.argv[n]) #Ingest JSON into a new dataframe
        temp_string = str(sys.argv[n]) #Grab the name of the file from the arg
        temp_string = temp_string.split('/')[-1] #Cut off any file folder prefixes
        temp_string = temp_string.split('.')[0] #Remove the .json suffix
        setup_names.append(temp_string) #Add this processed name to the list
        print("Setup " + str(n-1) + " = " + setup_names[n-1])
        n = n+1 #Increment the argv counter

    except Exception as e:
        #print(e)
        more_args = False
        print("No more args!")



#print("First file: " + str()


#STEP 1: Define XLS Styles
style0 = xlwt.easyxf('font: name Arial, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

style_header = xlwt.easyxf('font: name Arial, color-index black, bold on;  align: vert centre, horiz centre', 
    num_format_str='#,##0.00')

style_header_blue = xlwt.easyxf('font: name Arial, color-index blue, bold on;  align: vert centre, horiz centre', 
    num_format_str='#,##0.00')

##################
#Full XFStyle()

style_super_header = xlwt.XFStyle()

#Basic black background
black_pattern = xlwt.Pattern()
black_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
black_pattern.pattern_fore_colour = xlwt.Style.colour_map['blue']

#Alignment 
align = xlwt.Alignment()
align.horz = align.HORZ_CENTER
align.vert = align.VERT_CENTER


style_super_header.font.colour_index = xlwt.Style.colour_map['white']
style_super_header.font.bold = True
style_super_header.pattern = black_pattern
style_super_header.alignment = align




### Creating Workbook and Tabs ####################################

wb = xlwt.Workbook()

#--Sheets
ws_tyres = wb.add_sheet('Tyres')
ws_electronics = wb.add_sheet('Electronics')
ws_strategy = wb.add_sheet('Strategy')
ws_mechanical = wb.add_sheet('Mechanical Balance')
ws_aero = wb.add_sheet('Aerodynamics')

#--Set up first page w/ car name at the top
col_tyres_setups = ws_tyres.col(0) #Initializing the first column
col_tyres_setups.width = 300*20 #Column width of the first column (setups names)
car_name = "Car: " + str(df_setup[0]['carName'][0]) #Forming the car label at the top
print(car_name) #Debug
#--Car Name
ws_tyres.write(0,0,car_name, style0) #Write the car name






### Writing Headers ####################################
ws_tyres.write(2,0,"Setup Name", style_header)
ws_tyres.write(2,1,"Comp.", style_header)

#--Take the cells to merge as r1, r2, c1, c2, and accept an optional style parameter.
ws_tyres.write_merge(1, 1, 2, 5, '----Pressures (PSI)----', style_super_header)

ws_tyres.write(2,2,"FL", style_header_blue)
ws_tyres.write(2,3,"FR", style_header_blue)
ws_tyres.write(2,4,"RL", style_header_blue)
ws_tyres.write(2,5,"RL", style_header_blue)


####################################

#STEP 0: Import JSON 

#data = json.load(file_setup_json_0)

#df_setup_0 = pd.read_json(file_setup_json_0)

#print(df_setup_0['tyres'])
#file_setup_json_0.close()
#file_setup_json_1.close()

def setup_loop(row, df, col_offset, value_offset, value_divisor):
    incr_offset = 0
    for x in df:
        incr = x/value_divisor
        value = value_offset + incr
        col = col_offset+incr_offset
        ws_tyres.write(row, col, value) #Write Tyre Pressure (LR)
        incr_offset+=1

def process_setup(setup_num):
    #TYRES
    row = setup_num + ROW_OFFSET
    #--Headers

    if(df_setup[setup_num]['basicSetup']['tyres']['tyreCompound']):
        tyreCompound = "Wet"
    else: 
        tyreCompound = "Dry"

    #--Data
    ws_tyres.write(row,0, setup_names[setup_num]) #Write Setup Name
    ws_tyres.write(row,1, tyreCompound) #Write Tyre Compound

    first_offset = 2
    #Pressures
    setup_loop(row, df_setup[setup_num]['basicSetup']['tyres']['tyrePressure'], first_offset, 20.3, 10)
    #Camber
    setup_loop(row, df_setup[setup_num]['basicSetup']['alignment']['camber'], first_offset+4, -4, 10)
    #Toe
    setup_loop(row, df_setup[setup_num]['basicSetup']['alignment']['toe'], first_offset+4+4, -.40, 100)
    #CasterLF
    ws_tyres.write(row, first_offset+4+4+4, df_setup[setup_num]['basicSetup']['alignment']['casterLF'])
    #CasterRF
    ws_tyres.write(row, first_offset+4+4+4+1, df_setup[setup_num]['basicSetup']['alignment']['casterRF'])
    #SteeringRatio
    ws_tyres.write(row, first_offset+4+4+4+1+1, df_setup[setup_num]['basicSetup']['alignment']['steerRatio'])



#Step 3: Populate these sheets
for x in df_setup:
    process_setup(x)

ws_electronics.write(1, 0, datetime.now(), style1)
ws_strategy.write(2, 0, 1)
ws_mechanical.write(2, 1, 1)
ws_aero.write(2, 2, xlwt.Formula("A3+B3"))

wb.save('example.xls')






#Enable this to keep interactive terminal open
#code.interact(banner="Start", local=locals(), exitmsg="End")


