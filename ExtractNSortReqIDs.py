# -*- coding: utf-8 -*-
"""
Created on Feb 21 2019

@author: benjamin.roustan
"""

# for command line arguments
import sys, getopt
#specific to extracting information from word documents
import os
import zipfile
#other tools useful in extracting the information from our document
import re
#to navigate  our xml:
import xml.etree.ElementTree as ET


def main(scriptn, argv):
   in_file = ""
   out_file = ""
   try:
      opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
   except getopt.GetoptError:
      print ('{} -i <inputfile> -o <outputfile>'.format(scriptn))
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print ('{} -i <inputfile> -o <outputfile>'.format(scriptn))
         sys.exit()
      elif opt in ("-i", "--ifile"):
         in_file = arg
         if out_file == "":
            out_file = os.path.splitext(in_file)[0]+".outspecs.txt"
      elif opt in ("-o", "--ofile"):
         out_file = arg
   if in_file == "":
      print ('{} -i <inputfile> -o <outputfile>'.format(scriptn))
      sys.exit(2)
   print ('Input file is "{}"'.format(in_file))
   parseFile(in_file, out_file)
   print ('Output file is "{}"'.format(out_file))





def parseFile(input, output):
    WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    TBL = WORD_NAMESPACE + 'tbl'
    TEXT = WORD_NAMESPACE + 't'
    doc = zipfile.ZipFile(input)
    xx= doc.read('word/document.xml')
    doc.close()
    tree = ET.XML(xx)

    allID = []
    for x in tree.getiterator(TBL):
        texts = [node.text for node in x.getiterator(TEXT) if node.text]
        txt = ''.join(texts)
        if txt.find("EAUS_SW_")==0:
            #print ("{} => {}".format(x,txt))
            allID.append(txt)
            
    #print (allID)
    dicID={}
    for k in allID:
        id = k[0:17]
        curkey = int(re.sub("[-_]","",id[0:17][-9:]))
        if curkey in dicID:
            dicID[curkey].append({"ID": id, "text":k[17:]})
        else:
            dicID[curkey] = [{"ID": id, "text":k[17:]}]
    #print (dicID)

    outs = ""
    nbspecs = 1
    p="toutouyoutou"
    for k in sorted(dicID.keys()):
        id  = dicID[k][0]['ID']
        if (id[0:13]!=p):
            outs+= "------\n"
            p=id[0:13]
        
        nbreqs = len(dicID[k])
        if nbreqs>1:
            id = "/!\\ "+id
        for i in range(nbreqs):
            txt = dicID[k][i]['text']
            try:
              outs += "{:03} - {} -- {}\n".format(nbspecs, id, txt)
            except:
              outs +="{:03} - {} -- ERROR ENCODE\n".format(nbspecs, id)
            nbspecs +=1
    outs += "-- END OF SPECS --"
    f = open(output, 'w')
    f.write(outs)
    f.close()


if __name__ == "__main__":
   main(sys.argv[0],sys.argv[1:])