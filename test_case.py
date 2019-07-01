import openpyxl
import sys
import xml.etree.cElementTree as ET
import datetime

excel_document = openpyxl.load_workbook('Adm.xlsx')

sheets = excel_document.get_sheet_names()

#print sheets

sheet = excel_document.get_sheet_by_name(u'Testf\xe4lle')

#root = ET.Element("xml")
#doc = ET.SubElement(root, "testsuite")


#multiple_cells = sheet['A2':'M252']
# for row in multiple_cells:
#    for cell in row:
#        print cell.value,
#        ET.SubElement(doc, "field1", name="blah").text = cell.value
#    print '\n'


root = ET.Element("testsuite")

i = 0
for rowOfCellObjects in sheet.rows:
    doc = ET.SubElement(root, "testcase internalid='"+str(i)+"'")
    j = 0

    for cellObj in rowOfCellObjects:
        print(cellObj.coordinate, cellObj.value)
        print type(cellObj.value)

        datatype = type(cellObj.value)
        if(datatype is unicode):
            ET.SubElement(doc, cellObj.coordinate).text = cellObj.value
        elif(datatype is long):
            print int(cellObj.value)
            ET.SubElement(doc, cellObj.coordinate).text = str(
                int(cellObj.value))
        elif isinstance(cellObj.value, datetime.date):
            #print cellObj.value.strftime('%m/%d/%Y')
            # pass
            ET.SubElement(doc, cellObj.coordinate).text = str(
                cellObj.value.strftime('%m/%d/%Y'))

        else:
            print "pass"
            pass

        j = j+1

    i = i+1
    print('--- END OF ROW ---')


tree = ET.ElementTree(root)
tree.write("filename.xml")
