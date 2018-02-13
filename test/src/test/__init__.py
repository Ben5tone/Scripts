from os.path import join, dirname, abspath
import xlrd
from xlrd.sheet import ctype_text
import xml.etree.ElementTree as ET


#Iterations over columns
"""print('(Column #) type:value')
for idx, cell_obj in enumerate(row):
    cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
    print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))"""
cell_value = []   
#Iterations over columns and cells
def GetCellValue(): 
    num_cols = x1_sheet.ncols  # Number of columns
    for row_idx in range(0, x1_sheet.nrows):    # Iterate through rows
        print ('-'*40)
        print ('Row: %s' % row_idx)   # Print row number
        for col_idx in range(0, num_cols):  # Iterate through columns
            cell_obj = x1_sheet.cell(row_idx, col_idx).value  # Get cell object by row, col
            if cell_obj == '':
                print ('Column: %s Cell: %s' % (col_idx, cell_obj))
            elif cell_obj == 'ge1' or 'ge2' or 'ge3' or 'ge4' or 'ge5' or 'ge6':
                cell_obj = x1_sheet.cell(row_idx, col_idx).value
                print ('Column: %s Cell: %s' % (col_idx, cell_obj))
                cell_value.insert(0, cell_obj)
            elif cell_obj == 'b1' or 'b2' or 'b3' or 'b5' or 'b6':
                cell_obj = x1_sheet.cell(row_idx, col_idx).value
                print ('Column: %s Cell: %s' % (col_idx, cell_obj))
                cell_value.insert(0, cell_obj)
            elif cell_obj == 'k1' or 'k2' or 'k3' or 'k4' or 'k5' or 'k6':
                cell_obj = x1_sheet.cell(row_idx, col_idx).value
                print ('Column: %s Cell: %s' % (col_idx, cell_obj))
                cell_value.insert(0, cell_obj)
            else:
                print(None)
            


#Remove the GE_VALUES
def RemoveAttribs():
    a.pop("ge1",None)
    a.pop("ge2",None)
    a.pop("ge3",None)
    a.pop("ge4",None)
    a.pop("ge5",None)
    a.pop("ge6",None)
#Asignacion del path donde se encuentra el workbook
fname = join(dirname(dirname(abspath(__file__))),'test_data', 'Book1.xlsx')
xname = join(dirname(dirname(abspath(__file__))),'test_data', 'INLINEC.xml')

#Abrir el workbook
x1_workbook = xlrd.open_workbook(fname)

#Obtener nombres de las sheets que se encuentras en el workbook
sheet_names = x1_workbook.sheet_names()
print('Sheet Names',sheet_names)

x1_sheet = x1_workbook.sheet_by_index(0)
row = x1_sheet.row(0)

tree = ET.parse(xname)
root = tree.getroot()
a = root.find("./NVHC_PROPERTY//GE_VALUES").attrib
b = root.find("./NVHC_PROPERTY//GE_VALUES")

#RemoveAttribs()
#a.set('t1','update')
#b = a.set('a','1')
#tree.write('output.xml')
GetCellValue()
print(root.tag) 
print(a)
print(cell_value)




