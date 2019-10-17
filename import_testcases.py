import xml.etree.ElementTree as ET
import pandas as pd
import sys, getopt

def subelement_with_text(parent, tag, attrib={}, text=''):
    """
    Implements the recurrent task of creating a new child node with a given text (and attributes)
    """
    sub = ET.SubElement(parent, tag, attrib)
    sub.text = str(text)
    return sub

# path of excel file to convert is passed via the --path argument (python xxx.py --path path/to/file.xlsx)
opts = dict(getopt.getopt(sys.argv[1:], None, ['path='])[0])

# argument is mandatory
if '--path' not in opts:
    raise AssertionError('--path must be specified')

# import file with given path
infile = pd.read_excel(opts['--path'], sheet_name=None)

# excel file may contain multiple sheets. Loop over them and create an xml per sheet
for (sheetname, sheet) in infile.items():
    # if sheet is empty, skip
    if 0 in sheet.shape:
        continue
        
    # try this sheet and skip it if conversion fails
    try:
        # root node is <testcases>
        root = ET.Element('testcases')
        # drop empty lines that were imported for some reason
        sheet = sheet.loc[sheet.notna().any(axis=1),].copy()
        # lowercase column names and strip whitespace
        sheet.columns = map(lambda x: str(x).lower().strip(), sheet.columns)
        # tried to deal with different names for the same columns on different sheets here 
        sheet.rename({'name 1': 'name', 
                      'test name': 'name',
                      'test description': 'description', 
                      'test steps':'test step description', 
                      'expected result':'test step expected result'
                     }, axis='columns',inplace=True)
        # this fills in the name of the last test case in empty cells (often left out when there are multiple steps)
        sheet['name'].fillna(method='ffill', inplace=True)
        # loop over unique test case names
        for tc_name in sheet['name'].unique():
            # only select rows with the given test case name
            rows = sheet.loc[sheet['name']==tc_name,:].copy()
            rows.sort_values(by='test step #', inplace=True)
            # fill empty fields with empty strings (avoids errors later on)
            rows.fillna('', inplace=True)
            # get rid of line breaks in test case name
            firstrow=rows.iloc[0,:].copy()
            tcname = " ".join(firstrow['name'].split())
            # create test case node
            testcase = ET.SubElement(root, 'testcase', {'name': tcname})
            # create summary node and fill it with description text
            subelement_with_text(testcase,'summary', text=rows['description'].str.cat())
            keywords = ET.SubElement(testcase, 'keywords')
            
            ## The second file did not have keyword or Author columns, so I commented the following out
            #for kw in row['keywords'].split(","):
            #    ET.SubElement(keywords, 'keyword', {'name': kw.strip()})
            #custom_fields = ET.SubElement(testcase, 'custom_fields')
            #for fieldname,value in {'Autor1': row['bearbeiter 1'], 'Autor2': row['bearbeiter 2']}.items():
            #    custom_field = ET.SubElement(custom_fields, 'custom_field')
            #    subelement_with_text(custom_field, 'name', text=fieldname)
            #    subelement_with_text(custom_field, 'value', text=value)
            
            # create steps node
            steps = ET.SubElement(testcase, 'steps')
            # resetting the index to integer 0:n-1, using that as step number
            rows.reset_index(inplace=True)
            # loop over rows for this test case, each one containing a step
            for index, row in rows.iterrows():
                # create step node
                step = ET.SubElement(steps, 'step')
                # create 3 child nodes: actions, expected result and step number
                subelement_with_text(step, 'actions', text=row['test step description'])
                subelement_with_text(step, 'expectedresult', text=row['test step expected result'])
                subelement_with_text(step, 'step_number', text=index+1)
            
            # finally, write out the contents of the XML tree to sheetname.xml, with xml declaration and all
            ET.ElementTree(root).write(sheetname+".xml", encoding="UTF-8", xml_declaration=True, short_empty_elements=False)
            
    # if shit goes south, say so
    except Exception as err:
        print("Sheet " + sheetname + " failed")
        continue
