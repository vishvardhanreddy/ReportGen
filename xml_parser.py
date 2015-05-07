import xml.etree.ElementTree as ET
import csv
import tablib
from collections import OrderedDict
from collections import defaultdict


def setAnpHeader():
    headerlist = ['descr','RUpos','RUsize','id','rackNum','rackPos','type','pos','occupancy','eqtype']
    return headerlist

def setParametersHeader():
    headerlist = ['descr','shelf_position','slot_position','port_position','ppm_position','unit_id','name','value']
    return headerlist

def setPpmsHeader():
    headerlist = ['descr','ppmshelf','ppmslot','ppmnum','ppmtype','ppmport','ppmporttype']
    return headerlist

def setPatchcordsHeader():
    headerlist = ['descr','from_unit','to_unit','from_shelf','to_shelf','from_slot','to_slot','from_port','to_port']
    return headerlist

def setSidesHeader():
    headerlist = ['descr','side_name','shelf_in','slot_in','port_in','shelf_out','slot_out','port_out']
    return headerlist

def setAnsHeader():
    headerlist = ['descr','shelf_position','slot_position','port_position']
    return headerlist

def writetoCSV(dictlist, tagtype):
    filename = ('2.Process//nodesetup-'+tagtype+'.csv')
    
    headerlist = []
    if tagtype == 'anp':
        headerlist = setAnpHeader()
    elif tagtype == 'parameters':
        headerlist = setParametersHeader()
    elif tagtype == 'ppms':
        headerlist = setPpmsHeader()
    elif tagtype == 'patchcords':
        headerlist = setPatchcordsHeader()
    elif tagtype == 'sides':
        headerlist = setSidesHeader()
    elif tagtype == 'ans':
        headerlist = setAnsHeader()
        
    header = dictlist[0].keys()
    for i in dictlist:
        for j in i.keys():
            if j not in header:
                header.append(j)
    for i in header:
        if i not in headerlist:
            headerlist.append(i)
            
    with open(filename,'wb') as out:
        writer = csv.DictWriter(out,headerlist,extrasaction='raise', dialect = 'excel')
        writer.writeheader()
        writer.writerows(dictlist)
    
    
def anpCSV(root):
    
    dictlist = []
    for innode in root:
        anps = innode.findall(".//anp")
        for anp in anps:
            for shelf in anp:
                for slot in shelf:
                    elementdict = {}
                    for element in slot:
                        elementdict.update({element.tag:element.text})
                    anpdict = {}
                    anpdict.update(elementdict)
                    anpdict.update(slot.attrib)
                    anpdict.update(shelf.attrib)
                    anpdict.update(innode.attrib)
                    dictlist.append(anpdict)
    writetoCSV(dictlist,'anp')
                    
def tagtoCSV(root, tagtype):
    parameterlist = []
    for innode in root:
        nodedata = innode.findall(".//"+tagtype)
        for parameters in nodedata:
            for parameter in parameters:
                paramdict = {}
                for data in parameter:
                    paramdict.update({data.tag:data.text})
                    if data.attrib:
                        for i in data.attrib.keys():
                           data.attrib[data.tag+'_'+i] = data.attrib.pop(i)
                        paramdict.update(data.attrib)
                paramdict.update(innode.attrib)
                paramdict.update(parameters.attrib)
                paramdict.update(parameter.attrib)
                parameterlist.append(paramdict) 
    writetoCSV(parameterlist, tagtype)    
    return None
    
def main():
        tree = ET.parse('1.Collector//neupdate.xml')
        root = tree.getroot()
        anpCSV(root)
        tagtoCSV(root,'parameters')
        tagtoCSV(root,'ppms')
        tagtoCSV(root,'patchcords')
        tagtoCSV(root,'sides')
        tagtoCSV(root,'ans')
        
       


if __name__ == '__main__':
        main()
                                             