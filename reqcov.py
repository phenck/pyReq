#specific to extracting information from word documents
import os
import zipfile 
#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom

document = zipfile.ZipFile('CC10515_CON_SPEC_ICD_COPRA.docx')
for item in document.namelist():
    print(item)
uglyXml = xml.dom.minidom.parseString(document.read('word/document.xml')).toprettyxml(indent='  ')

text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)    
prettyXml = text_re.sub('>\g<1></', uglyXml)

print(uglyXml)