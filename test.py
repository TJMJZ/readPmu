import xlrd
from lxml import etree
import re
import os

_digits = re.compile('\d')
def contains_digits(d):
    return bool(_digits.search(d))
def formatItem(d):
    return d.replace(".","").replace(" ","")

def isSummary(rowItem):
    itemID = rowItem[0]
    itemVal = rowItem[2]
    if not formatItem(itemID).isdigit():
        # print(itemID)
        pass
    if "/" in itemID:
        # print(itemID)
        return -1
    elif itemVal == "":
        return 1
    else:
        return 0

def calNodeSum(currCrusor):
    subNodes = currCrusor.xpath('.//LEAF')

    temptotal = 0.0
    for tempNode in subNodes:
        # temptotal = tempNode.get("THISPRD")+temptotal
        if tempNode.find("THISPRD").text == "":
            tempnum = 0.0
        else:
            tempnum = float(tempNode.find("THISPRD").text)
            # if tempnum >0 :
            #     print(tempnum)
        
        temptotal = tempnum+temptotal
    
    # currCrusor.set("SUMTOTAL",str(rowItem[0]))
    currCrusor.set("SUMTHISP",str(temptotal))
    # currCrusor.set("SUMTPREV",str(rowItem[0]))
    return



def isCChead(rowEntry):
    format_val = formatItem(rowEntry[0]).lower()
    #print(format_val)
    if not (format_val.isdigit() or format_val == ""):
        if "costcenter" in format_val.lower():
            return True
    return False

def getCCheadID(rowEntry):
    rowItem = rowEntry[0]
    # print(rowItem)
    ccIDString = ""
    validC = ['1','2','3','4','5','6','7','8','9','0','.']
    for c in rowItem:
        if c in validC:
            # print(c)
            ccIDString = ccIDString + c
    # print(ccIDString)
    return ccIDString


def categorizeSubCC(iptStr):   
    nameStr = iptStr.lower()
    TN = "tunnel"
    BG = "bridge"
    RD = "road"
    if TN in nameStr:
        return "TN"
    elif BG in nameStr:
        return "BG"
    elif RD in nameStr:
        return "RD"
    else:
        print("subcc categorize error:"&iptStr)

def splitTnStr(spltStr,iptLR):

    if iptLR>0:
        LRstr = "RK"
    elif iptLR<0:
        LRstr = "LK"
    else:
        pass
    splitStrList = spltStr.split(LRstr)
    if len(splitStrList)<3:
        print("-------------split string error: "+spltStr)
        splitStrList = spltStr.split("K")

    pile1str = splitStrList[1].replace("L","").replace("R","")
    pile2str = splitStrList[2].replace("L","").replace("R","").replace(")","")

    try:
        pile1 = float(pile1str)
        pile2 = float(pile2str)
    except Exception as e:
        print(splitStrList)
        pile1str = pile1str.replace(",",".")
        pile2str = pile2str.replace(",",".")
        try:
            pile1 = float(pile1str)
            pile2 = float(pile2str)
        except Exception as e:
            pile1str = pile1str.replace(".","",1)
            pile2str = pile2str.replace(".","",1)
            pile1 = float(pile1str)
            pile2 = float(pile2str)




    return [pile1,pile2]

def bridgeLRjudge(iptStr0,iptStr1):
    iptStr0 = iptStr0.replace(" ","")
    iptStr1 = str(iptStr1).lower()
    if iptStr0 == "":
        if "bridge" in iptStr1:
            if ("left" in iptStr1):
                return 1
            elif("right" in iptStr1):
                return 2
    return -1


def parseTNmile(iptStr):
    if len(iptStr)<10:
        # print(iptStr)
        return ["NOT YET",-1,-1]
    iptStr = iptStr.replace("-","").replace("+","").replace(" ","").replace("—","").replace("~","")
    print(iptStr)
    #tnPattern = "\d*[mM](.*)[nNsS]-*[lLrR](.*)"
    #tnPatternRe = re.compile(tnPattern)
    if "SR" in iptStr:
        portal = "SR"
        pile = splitTnStr(iptStr,1)

    elif "NR" in iptStr:
        portal = "NR"
        pile = splitTnStr(iptStr,1)

    elif "SL" in iptStr:
        portal = "SL"
        pile = splitTnStr(iptStr,-1)

    elif "NL" in iptStr:
        portal = "NL"
        pile = splitTnStr(iptStr,-1)

    else:
        print("TUNNEL STRING INVALID ERROR")

    pile0 = pile[0]
    pile1 = pile[1]

    return [portal,pile0,pile1]

    # if len(test) > 10:
    #     print(test)

def parseBGpiers(iptStr):
    if len(iptStr)<5:
        return [iptStr,-1,-1]
    spltStr = iptStr.replace(" ","").replace("m","").replace(")","")
    spltStrList = spltStr.split("stage")
    pierNo = spltStrList[0]
    pierStage = spltStrList[1].split("(")[0]
    stageHeight = float(spltStrList[1].split("(")[1])
    return [pierNo,pierStage,stageHeight]

def parseRDmile(iptStr):

    spltStr = iptStr.replace("-","").replace("+","").replace(" ","").replace("—","").replace("~","").replace(",","")
    print(iptStr)

    lkStr = "LK"
    rkStr = "RK"

    if (lkStr in spltStr) and (rkStr in spltStr):
        splitLkList = spltStr.split(lkStr)
        splitRkList = splitLkList[2].split(rkStr)
        try:
            lk1 = float(splitLkList[1])
            lk2 = float(splitRkList[0])
            rk1 = float(splitRkList[1])
            rk2 = float(splitRkList[2].replace(")",""))
        except Exception as e:
            print(spltStr)
            lk1 = -1
            lk2 = -1
            rk1 = -1
            rk2 = -1
    elif (lkStr in spltStr) and (rkStr not in spltStr):
        splitLkList = spltStr.split(lkStr)
        lk1 = float(splitLkList[1])
        lk2 = float(splitLkList[2].replace(")",""))
        rk1 = -1
        rk2 = -1
    elif (lkStr not in spltStr) and (rkStr in spltStr):
        splitRkList = spltStr.split(rkStr)
        rk1 = float(splitRkList[1])
        rk2 = float(splitRkList[2].replace(")",""))
        lk1 = -1
        lk2 = -1
    else:
        lk1 = -1
        lk2 = -1
        rk1 = -1
        rk2 = -1
    return [[lk1,lk2],[rk1,rk2]]



wb = xlrd.open_workbook("test2.xlsx")

# create tree
IPC015ROOT = etree.Element('IPC015')
IPC015ROOT.set('DATE', '20171231')
IPC015ROOT.set('LEVEL', '1')
IPC015TREE = etree.ElementTree(IPC015ROOT)

# print corner cases
# for sheetNumber in range(1,15):
#     if sheetNumber < 10:
#         sh = wb.sheet_by_name("COST CENTER "+"0"+str(sheetNumber))
#     else:
#         sh = wb.sheet_by_name("COST CENTER "+str(sheetNumber))
#     for row in range(10, sh.nrows):
#         val = sh.row_values(row)[0]
#         format_val = formatItem(val)
#         if not (format_val.isdigit() or format_val == ""):
#             if not "costcenter" in format_val.lower():
#                 print(val)


for sheetNumber in range(1,15):
    print(sheetNumber)
    tempCC = etree.SubElement(IPC015ROOT,"CostCenter")
    tempCC.set("ID",str(sheetNumber))
    tempCC.set("LEVEL",'2')

    if sheetNumber < 10:
        sh = wb.sheet_by_name("COST CENTER "+"0"+str(sheetNumber))
    else:
        sh = wb.sheet_by_name("COST CENTER "+str(sheetNumber))

    for row in range(9, sh.nrows):
        rowItem = sh.row_values(row)
        #print(rowItem[0])
        # print(isCChead(rowItem))
        if isCChead(rowItem):

            ccID = getCCheadID(rowItem)
            subCC = etree.SubElement(tempCC,"SubCostCenter")
            # subCC.text = rowItem[1]
            subCC.set("LEVEL","3")
            subCC.set("ID",ccID)
            subCC.set("TEXT",rowItem[1])
            currCrusor = subCC
            if "bridge" in rowItem[1].lower():
                resetLbridge = 0
                resetRbridge = 0
        elif bridgeLRjudge(rowItem[0],rowItem[1])>0:
            if :
                pass
            lrBridge = etree.SubElement(subCC,"NODE")
            lrBridge.set("ID",subCC.get("ID"))
            lrBridge.set("TEXT",str(rowItem[1].split("/")[0]))
            currCrusor = lrBridge

        else:
            while True:

                if rowItem[0] == "":
                    break
                
                elif (currCrusor.get("ID") in rowItem[0]):
                    # create new entry
                    if isSummary(rowItem):
                        newentry = etree.SubElement(currCrusor,"NODE")
                        newentry.set("ID",str(rowItem[0]))
                        newentry.set("TEXT",str(rowItem[1].split("/")[0]))
                        currCrusor = newentry
                    
                    # is leaf
                    # leaf has no children
                    else:
                        newentry = etree.SubElement(currCrusor,"LEAF")
                        newentry.set("ID",str(rowItem[0]))
                        newentry.set("TEXT",str(rowItem[1]))

                        etree.SubElement(newentry,"TOTALAMT").text = str(rowItem[2])
                        etree.SubElement(newentry,"THISPRD").text = str(rowItem[6])
                        etree.SubElement(newentry,"PREVCM").text = str(rowItem[5])
                        if rowItem[6] == rowItem[2]:
                            etree.SubElement(newentry,"ISFULL").text = "1"
                        elif rowItem[5]+rowItem[6] == rowItem[2]:
                            etree.SubElement(newentry,"ISFULL").text = "2"
                        else:
                            etree.SubElement(newentry,"ISFULL").text = "0"
                        
                    
                    break
                    # print(currCrusor)
                else:
                    currCrusor = currCrusor.getparent()









IPC015TREE.write('test3.xml', pretty_print=True, xml_declaration=True, encoding='utf-8')


regexpNS = "http://exslt.org/regular-expressions"
find = IPC015TREE.xpath("//SubCostCenter[re:test(@TEXT, '.*tunnel.*','i')]",namespaces={"re": "http://exslt.org/regular-expressions"})


for temp in find:
    # print("+++++++++++++++++++++++++++++++++++")
    # print(temp.attrib["TEXT"])
    pass

exp = "//NODE[starts-with(@TEXT,'Excavation')]"
exp2 = "//NODE[re:test(@TEXT,'.*secondary.*','i')]"
exp3 = "//SubCostCenter[re:test(@TEXT,'.*road.*','i')]"
exp4 = "//SubCostCenter[re:test(@TEXT,'.*bridge.*','i')]"
exp5 = "//SubCostCenter[re:test(@TEXT,'.*bridge.*','i')]//"


for element in IPC015TREE.xpath(exp4,namespaces={"re": "http://exslt.org/regular-expressions"}):
    print("----------------------------------")
    print(element.attrib["TEXT"])
    print("----------------------------------")

    for child in element.iterchildren():
        # print(child.attrib["TEXT"])

        isLRchild = bridgeLRjudge("",child.attrib["TEXT"])
        if isLRchild == 1:
            
        if "Piers" in child.attrib["TEXT"]:
            for childchild in child.iterchildren():
                try:
                    piers = parseBGpiers(childchild.attrib["TEXT"])
                    

                    
                except Exception as e:
                    # print(childchild.attrib["TEXT"])
                    print(e)




'''

# findTunnel = IPC015TREE.xpath(exp)
findTunnel = IPC015TREE.xpath(exp2,namespaces={"re": "http://exslt.org/regular-expressions"})

for temp in findTunnel:
    # print("------------------------------------------------")
    # print(temp.getparent().attrib["TEXT"])
    for child in temp.iterchildren():
        templist = parseTNmile(child.attrib["TEXT"])
        # print(templist[1])
    #print(temp.attrib["ID"])
    


roadDict = {}

for element in IPC015TREE.xpath(exp3,namespaces={"re": "http://exslt.org/regular-expressions"}):
    print("----------------------------------")
    print(element.attrib["TEXT"])
    print("----------------------------------")
    
    print(parseRDmile(element.attrib["TEXT"]))


    for child in element.iterchildren():
        childKey = child.attrib["TEXT"].split('/')[0]
        if childKey in roadDict:
            roadDict[childKey] = roadDict[childKey]+1
        else:
            roadDict[childKey] = 1




print(roadDict)

for element in IPC015TREE.xpath('//IPC015'):
    calNodeSum(element)

for element in IPC015TREE.xpath('//CostCenter'):
    calNodeSum(element)

for element in IPC015TREE.xpath('//SubCostCenter'):
    calNodeSum(element)

for element in IPC015TREE.xpath('//NODE'):
    calNodeSum(element)

for bad in IPC015TREE.xpath("//SubCostCenter[@SUMTHISP<100]"):
    bad.getparent().remove(bad)

for bad in IPC015TREE.xpath("//NODE[@SUMTHISP<100]"):
    bad.getparent().remove(bad)

for bad in IPC015TREE.xpath("//LEAF[THISPRD<100]"):
    bad.getparent().remove(bad)




    #element = etree.SubElement(name, 'node')
    #element.set('id', str(val[0]))
    #element.set('x', str(val[1]))
    #element.set('y', str(val[2]))




# IPC015TREE.write('test2.xml', pretty_print=True, xml_declaration=True, encoding='utf-8')







# ThisPeriodNode = IPC015TREE.xpath('//SubCostCenter[@SUMTHISP>100]]')

# ThisPeriodRoot = etree.Element('ThisPeriod')

# ThisPeriodTree = etree.ElementTree(ThisPeriodRoot)

# for treeNode in ThisPeriodNode:
#     ThisPeriodRoot.append(treeNode)

# # ThisPeriodNode2 = ThisPeriodTree.xpath("//NODE[@SUMTHISP>100]")


# ThisPeriodTree.write('this.xml', pretty_print=True, xml_declaration=True, encoding='utf-8')

# sys.input()
'''