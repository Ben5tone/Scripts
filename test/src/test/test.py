'''
Created on 08/02/2018

@author: Julio Cesar Chavez Flores
         Python Software Engineer
         FORD CSAP
'''

from os.path import join, dirname, abspath
import xlrd
import xml.etree.ElementTree as ET

xml_bushings = []
exl_bushings = []

#Assign excel workbook path
fname = join(dirname(dirname(abspath(__file__))),'test_data', 'Extract_bushing_rates_tool.xlsm')
#Assign Inline .xml file
xname = join(dirname(dirname(abspath(__file__))),'test_data', 'INLINEC.xml')

#Open the workbook
x1_workbook = xlrd.open_workbook(fname)

#Obtain spreadsheet names of the current workbook
sheet_names = x1_workbook.sheet_names()
print('Sheet Names',sheet_names)

tree = ET.parse(xname)
root = tree.getroot()

x1_sheet = x1_workbook.sheet_by_index(0)

#Iterate over inline.xml file to obtain all the bushing names 
for child in root.iter("NVHC_PROPERTY"):
    for child2 in child.iter("NAME"):
        xml_bushings.insert(0, child2.text)
print(xml_bushings)


num_cols = x1_sheet.ncols  # Number of columns
for row_idx in range(0, x1_sheet.nrows):    # Iterate through rows
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = x1_sheet.cell(row_idx, col_idx).value  # Get cell object by row, col 
        if cell_obj == "Bushing":
            for col_idx in range(col_idx, num_cols):    # Iterate through rows
                for row_idx in range(row_idx, x1_sheet.nrows): # Iterate through columns
                    if x1_sheet.cell(row_idx, col_idx).value != "":
                        cell_obj2 = x1_sheet.cell(row_idx, col_idx).value
                        for i in range(xml_bushings.__len__()):
                            if cell_obj2 == xml_bushings[i]:
                                a = row_idx 
                                #print(cell_obj2, xml_bushings[i], i )
                                for col_idx2 in range(col_idx+1, 10):    # Iterate through rows
                                    for row_idx in range(col_idx+1, 6):  # Iterate through columns
                                        cell_obj3 = x1_sheet.cell(a-1, col_idx2).value
                                        print(cell_obj3,cell_obj2, xml_bushings[i], a)
                                        exl_bushings.insert(0, str(i) + "_" + str(cell_obj3))
                                        break                
                    
print(exl_bushings)


                                    
                        
    