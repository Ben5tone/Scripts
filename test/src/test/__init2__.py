from os.path import join, dirname, abspath
import xlrd
import xml.etree.ElementTree as ET

cell_value = {"ge1": 0,"ge2": 0,"ge3":0,"ge4":0,"ge5":0,"ge6":0,
              "b1": 0,"b2": 0,"b3": 0,"b4": 0,"b5":0,"b6":0,
              "k1":0,"k2":0,"k3":0,"k4":0,"k5":0,"k6":0}
   
#Iterations over columns and cells
def GetCellValue(): 
    num_cols = x1_sheet.ncols  # Number of columns
    for row_idx in range(0, x1_sheet.nrows):    # Iterate through rows
        for col_idx in range(0, num_cols):  # Iterate through columns
            cell_obj = x1_sheet.cell(row_idx, col_idx).value  # Get cell object by row, col
            if cell_obj == 'ge1':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["ge1"] = cell_obj2
            elif cell_obj == 'ge2':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["ge2"] = cell_obj2
            elif cell_obj == 'ge3':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["ge3"] = cell_obj2
            elif cell_obj == 'ge4':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["ge4"] = cell_obj2
            elif cell_obj == 'ge5':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["ge5"] = cell_obj2
            elif cell_obj == 'ge6':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["ge6"] = cell_obj2
            elif cell_obj == 'b1':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["b1"] = cell_obj2
            elif cell_obj == 'b2':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["b2"] = cell_obj2
            elif cell_obj == 'b3':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["b3"] = cell_obj2
            elif cell_obj == 'b4':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["b4"] = cell_obj2
            elif cell_obj == 'b5':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["b5"] = cell_obj2
            elif cell_obj == 'b6':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["b6"] = cell_obj2
            elif cell_obj == 'k1':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["k1"] = cell_obj2
            elif cell_obj == 'k2':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["k2"] = cell_obj2
            elif cell_obj == 'k3':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["k3"] = cell_obj2
            elif cell_obj == 'k4':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["k4"] = cell_obj2
            elif cell_obj == 'k5':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["k5"] = cell_obj2
            elif cell_obj == 'k6':
                cell_obj2 = x1_sheet.cell(row_idx + 1, col_idx).value
                print('%s' % (cell_obj2))
                cell_value["k6"] = cell_obj2
          
          
#Remove the GE_VALUES
def RemoveAttribs():
    a.pop("ge1",None)
    a.pop("ge2",None)
    a.pop("ge3",None)
    a.pop("ge4",None)
    a.pop("ge5",None)
    a.pop("ge6",None)
    b.pop("b1",None)
    b.pop("b2",None)
    b.pop("b3",None)
    b.pop("b4",None)
    b.pop("b5",None)
    b.pop("b6",None)
    c.pop("k1",None)
    c.pop("k2",None)
    c.pop("k3",None)
    c.pop("k4",None)
    c.pop("k5",None)
    c.pop("k6",None)

def WriteAttribs():
    
    a1.set("ge1", str(cell_value.get("ge1","none")))
    a1.set("ge2", str(cell_value.get("ge2","none")))
    a1.set("ge3", str(cell_value.get("ge3","none")))
    a1.set("ge4", str(cell_value.get("ge4","none")))
    a1.set("ge5", str(cell_value.get("ge5","none")))
    a1.set("ge6", str(cell_value.get("ge6","none")))
    b1.set("b1", str(cell_value.get("b1","none")))
    b1.set("b2", str(cell_value.get("b2","none")))
    b1.set("b3", str(cell_value.get("b3","none")))
    b1.set("b4", str(cell_value.get("b4","none")))
    b1.set("b5", str(cell_value.get("b5","none")))
    b1.set("b6", str(cell_value.get("b6","none")))
    c1.set("k1", str(cell_value.get("k1","none")))
    c1.set("k2", str(cell_value.get("k2","none")))
    c1.set("k3", str(cell_value.get("k3","none")))
    c1.set("k4", str(cell_value.get("k4","none")))
    c1.set("k5", str(cell_value.get("k5","none")))
    c1.set("k6", str(cell_value.get("k6","none")))

    
#Asignacion del path donde se encuentra el workbook
fname = join(dirname(dirname(abspath(__file__))),'test_data', 'Book2.xlsx')
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
a1 = root.find("./NVHC_PROPERTY//GE_VALUES")
b = root.find("./NVHC_PROPERTY//B_VALUES").attrib
b1 = root.find("./NVHC_PROPERTY//B_VALUES")
c = root.find("./NVHC_PROPERTY//K_VALUES").attrib
c1 = root.find("./NVHC_PROPERTY//K_VALUES")

control = False
"""for child in root.iter("NVHC_PROPERTY"):
    for child2 in child:
        for child3 in child2:
            for child4 in child3.iter("B_VALUES"):
                    child3.attrib.pop("b1", None)
                    child3.attrib.pop("b2", None)
                    child3.attrib.pop("b3", None)
                    child3.attrib.pop("b4", None)
                    child3.attrib.pop("b5", None)
                    child3.attrib.pop("b6", None)"""



        
"""for child2 in child:
            for child3 in child2:
        if child3.tag == "B_VALUES" and control != True:
                    control = True
                    b3 = root.find("./NVHC_PROPERTY//B_VALUES").attrib
                    b3.pop("b1", None)
                    b3.pop("b2", None)
                    b3.pop("b3", None)
                    b3.pop("b4", None)
                    b3.pop("b5", None)
                    b3.pop("b6", None)
        elif child3.tag == "K_VALUES"
        """
for child in root.iter("NVHC_PROPERTY"):
    for child2 in child.iter("NAME"):
        print(child2.attrib)
                          
GetCellValue()
#RemoveAttribs()
WriteAttribs()
tree.write('output.xml')
print(root.tag) 
print(a)
print(b)
print(c)
print(cell_value)




