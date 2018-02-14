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

oname = join(dirname(dirname(abspath(__file__))),'test_data', 'output2.xml')

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

cont = 0
bushings_dict = {}
num_cols = x1_sheet.ncols  # Number of columns
for row_idx in range(0, x1_sheet.nrows):    # Iterate through rows
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = x1_sheet.cell(row_idx, col_idx).value  # Get cell object by row, col 
        if cell_obj == "Bushing":
            for col_idx in range(col_idx, num_cols):    # Iterate through rows
                for row_idx in range(row_idx, x1_sheet.nrows): # Iterate through columns
                    #if x1_sheet.cell(row_idx, col_idx).value != "":
                    cell_obj2 = x1_sheet.cell(row_idx, col_idx).value
                    for i in range(xml_bushings.__len__()):
                        if cell_obj2 == xml_bushings[i]:
                            a = row_idx 
                            #print(cell_obj2, xml_bushings[i], i )
                            for col_idx2 in range(col_idx+1, 10):    # Iterate through rows
                                for row_idx in range(col_idx+1, 6):  # Iterate through columns
                                    cell_obj3 = x1_sheet.cell(a, col_idx2).value
                                    #print(cell_obj3,cell_obj2, xml_bushings[i], a + 1)
                                    exl_bushings.insert(0, cell_obj3)
                                    cont = cont + 1
                                    if cont == 6:
                                        cont = 0
                                        bushings_dict[str(cell_obj2)] = exl_bushings
                                        exl_bushings = [] 
                                        break
                                    break

print(bushings_dict)

parameters_list = []
for child3 in root.iter("NVHC_PROPERTY"):
    for child4 in child3.iter("NAME"):
        for key,value in bushings_dict.items():
            if child4.text == key:
                
                print(key)
                #Accesign to each element of the list that is inside the found key
                r1 = bushings_dict[child4.text][0]
                r2 = bushings_dict[child4.text][1] 
                r3 = bushings_dict[child4.text][2] 
                r4 = bushings_dict[child4.text][3] 
                r5 = bushings_dict[child4.text][4] 
                r6 = bushings_dict[child4.text][5]
                  
                for child4 in child3.iter("PARAMETERS"):
                    for child5 in child4:
                        if child5.tag == "B_VALUES":
                            
                            child5.set("b1", str(r6))
                            child5.set("b2", str(r5))
                            child5.set("b3", str(r4))
                            child5.set("b4", str(r3))
                            child5.set("b5", str(r2))
                            child5.set("b6", str(r1))
                            
                            tree.write(oname)  
                            print(child5.attrib)
