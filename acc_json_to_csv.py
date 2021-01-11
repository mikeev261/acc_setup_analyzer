import sys
import json
import csv
import xlwt
import pandas as pd
from datetime import datetime

import code


print("Hello world")

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

####################################

#STEP 0: Import JSON 

#data = json.load(file_setup_json_0)

#df_setup_0 = pd.read_json(file_setup_json_0)

#print(df_setup_0['tyres'])
#file_setup_json_0.close()
#file_setup_json_1.close()








def process_setup():
    #TYRES

    #--Headers
    ws_tyres.write(2,0,"Setup Name", style_header)
    ws_tyres.write(2,1,"Comp.", style_header)

    #Take the cells to merge as r1, r2, c1, c2, and accept an optional style parameter.
    ws_tyres.write_merge(1, 1, 2, 5, '----Pressures (PSI)----', style_super_header)

    ws_tyres.write(2,2,"FL", style_header_blue)
    ws_tyres.write(2,3,"FR", style_header_blue)
    ws_tyres.write(2,4,"RL", style_header_blue)
    ws_tyres.write(2,5,"RL", style_header_blue)

    #--Data
    #ws_tyres.write(3,0, setup_names[0])
    #ws_tyres.write(4,0, setup_names[1])






#STEP 2: Create XLS Workbook
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

#Step 3: Populate these sheets
process_setup()

ws_electronics.write(1, 0, datetime.now(), style1)
ws_strategy.write(2, 0, 1)
ws_mechanical.write(2, 1, 1)
ws_aero.write(2, 2, xlwt.Formula("A3+B3"))

wb.save('example.xls')






#Enable this to keep interactive terminal open
#code.interact(banner="Start", local=locals(), exitmsg="End")


