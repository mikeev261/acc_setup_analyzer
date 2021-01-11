import sys
import json
import csv
import xlwt
import pandas as pd
from datetime import datetime

import code


print("ACC Setup Analyzer")

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

style_2_dec = xlwt.easyxf(num_format_str='0.00')
style_1_dec = xlwt.easyxf(num_format_str='0.0')

style_header = xlwt.easyxf('font: name Arial, color-index black, bold on;  align: vert centre, horiz centre', 
    num_format_str='#,##0.00')

style_header_blue = xlwt.easyxf('font: name Arial, color-index blue, bold on;  align: vert centre, horiz centre', 
    num_format_str='#,##0.00')

style_header_red = xlwt.easyxf('font: name Arial, color-index red, bold on;  align: vert centre, horiz centre', 
    num_format_str='#,##0.00')

style_header_green = xlwt.easyxf('font: name Arial, color-index green, bold on;  align: vert centre, horiz centre', 
    num_format_str='#,##0.00')

style_header_orange = xlwt.easyxf('font: name Arial, color-index orange, bold on;  align: vert centre, horiz centre', 
    num_format_str='#,##0.00')
##################
#Full XFStyle()



#Basic black background
blue_pattern = xlwt.Pattern()
blue_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
blue_pattern.pattern_fore_colour = xlwt.Style.colour_map['blue']

red_pattern = xlwt.Pattern()
red_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
red_pattern.pattern_fore_colour = xlwt.Style.colour_map['red']

green_pattern = xlwt.Pattern()
green_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
green_pattern.pattern_fore_colour = xlwt.Style.colour_map['green']

orange_pattern = xlwt.Pattern()
orange_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
orange_pattern.pattern_fore_colour = xlwt.Style.colour_map['orange']


#Alignment 
align = xlwt.Alignment()
align.horz = align.HORZ_CENTER
align.vert = align.VERT_CENTER



style_super_header_blue = xlwt.XFStyle()
style_super_header_blue.font.colour_index = xlwt.Style.colour_map['white']
style_super_header_blue.font.bold = True
style_super_header_blue.pattern = blue_pattern
style_super_header_blue.alignment = align

style_super_header_red = xlwt.XFStyle()
style_super_header_red.font.colour_index = xlwt.Style.colour_map['white']
style_super_header_red.font.bold = True
style_super_header_red.pattern = red_pattern
style_super_header_red.alignment = align

style_super_header_green = xlwt.XFStyle()
style_super_header_green.font.colour_index = xlwt.Style.colour_map['white']
style_super_header_green.font.bold = True
style_super_header_green.pattern = green_pattern
style_super_header_green.alignment = align

style_super_header_orange = xlwt.XFStyle()
style_super_header_orange.font.colour_index = xlwt.Style.colour_map['white']
style_super_header_orange.font.bold = True
style_super_header_orange.pattern = orange_pattern
style_super_header_orange.alignment = align


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
def write_headers(row, size, begin, end, label, style0, style1):
    ws_tyres.write_merge(row-1, row-1, begin, end, label, style0)
    ws_tyres.write(row,begin,"FL", style1)
    ws_tyres.write(row,begin+1,"FR", style1)
    if(size>2):
        ws_tyres.write(row,begin+2,"RL", style1)
        ws_tyres.write(row,begin+3,"RL", style1)


ws_tyres.write(2,0,"Setup Name", style_header)
ws_tyres.write(2,1,"Comp.", style_header)

###---PRESSURES
#--Take the cells to merge as r1, r2, c1, c2, and accept an optional style parameter.
write_headers(2, 4, 2, 5, 'Pressures (PSI)', style_super_header_blue, style_header_blue)

###---CAMBER
write_headers(2, 4, 6, 9, 'Camber', style_super_header_red, style_header_red)

###---TOE
write_headers(2, 4, 10, 13, 'Toe', style_super_header_green, style_header_green)

###---CASTER
ws_tyres.write_merge(1, 1, 14, 15, 'Caster', style_super_header_orange)
ws_tyres.write(2,14,"LF", style_header_orange)
ws_tyres.write(2,15,"RF", style_header_orange)

###---STEERING RATIO
ws_tyres.write(2,16,"Steer Ratio", style_header)
\
####################################

#STEP 0: Import JSON 

#data = json.load(file_setup_json_0)

#df_setup_0 = pd.read_json(file_setup_json_0)

#print(df_setup_0['tyres'])
#file_setup_json_0.close()
#file_setup_json_1.close()

def setup_loop(row, df, col_offset, value_offset, value_divisor, style):
    incr_offset = 0
    for x in df:
        incr = x/value_divisor
        value = value_offset + incr
        col = col_offset+incr_offset
        ws_tyres.write(row, col, value, style) #Write Tyre Pressure (LR)
        incr_offset+=1

def process_setup(setup_num):
    #TYRES
    row = setup_num + ROW_OFFSET
    #--Headers

    print(df_setup[setup_num]['basicSetup']['tyres']['tyreCompound'])
    if(df_setup[setup_num]['basicSetup']['tyres']['tyreCompound']):
        tyreCompound = "Wet"
    else: 
        tyreCompound = "Dry"

    #--Data
    ws_tyres.write(row,0, setup_names[setup_num]) #Write Setup Name
    ws_tyres.write(row,1, tyreCompound) #Write Tyre Compound

    first_offset = 2
    #Pressures
    setup_loop(row, df_setup[setup_num]['basicSetup']['tyres']['tyrePressure'], first_offset, 20.3, 10, style_1_dec)
    #Camber
    setup_loop(row, df_setup[setup_num]['basicSetup']['alignment']['camber'], first_offset+4, -4, 10, style_1_dec)
    #Toe
    setup_loop(row, df_setup[setup_num]['basicSetup']['alignment']['toe'], first_offset+4+4, -.40, 100, style_2_dec)
    #CasterLF
    ws_tyres.write(row, first_offset+4+4+4, df_setup[setup_num]['basicSetup']['alignment']['casterLF'], style_1_dec)
    #CasterRF
    ws_tyres.write(row, first_offset+4+4+4+1, df_setup[setup_num]['basicSetup']['alignment']['casterRF'], style_1_dec)
    #SteeringRatio
    ws_tyres.write(row, first_offset+4+4+4+1+1, df_setup[setup_num]['basicSetup']['alignment']['steerRatio'])



#Step 3: Populate these sheets
for x in df_setup:
    process_setup(x)

#ws_electronics.write(1, 0, datetime.now(), style1)
ws_strategy.write(2, 0, 1)
ws_mechanical.write(2, 1, 1)
ws_aero.write(2, 2, xlwt.Formula("A3+B3"))

wb.save('example.xls')






#Enable this to keep interactive terminal open
#code.interact(banner="Start", local=locals(), exitmsg="End")


