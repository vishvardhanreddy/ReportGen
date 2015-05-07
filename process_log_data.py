'''
Created on Dec 08, 2014

@author: vvaka
'''
import re
import csv
import sys
import os
import shutil
import pandas as pd
import numpy as np
import time
import xml.etree.ElementTree as ET
from xlrd import open_workbook
from os import listdir
from os.path import isfile, join
from util import *

tempDir = curDir
if len(sys.argv) > 1:
    tempDir = sys.argv[1]

collectorDir = path_join(curDir, '1.Collector/%s'%tempDir)
processDir   = path_join(curDir, '2.Process/%s'%tempDir)
reportDir    = path_join(curDir, '3.Report/%s'%tempDir)
layoutDir    = path_join(collectorDir, 'layout')


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
    filename = path_join(processDir, 'nodesetup-'+tagtype+'.csv')

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
        anps = innode.findall("./anp")
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
        nodedata = innode.findall("./"+tagtype)
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

def billofMaterial(inputfile):
    outfile = path_join(processDir, 'Bill of Material.csv')
    wb = open_workbook(inputfile)
    sheet = wb.sheet_by_name('Net View (BoM)')

    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols
    Labels = sheet.row_values(8, start_colx=0, end_colx= None)

    for i in range(0, number_of_columns):
        colvalues = sheet.col_values(i, start_rowx=8, end_rowx= None)
        if colvalues[0] == 'Product ID':
            ProductID = list(colvalues)
        elif colvalues[0] == 'Description':
            Description = list(colvalues)
        elif colvalues[0] == 'Quantity':
            Quantity = list(colvalues)
            for i in range(0, len(Quantity)):
                if Quantity[i] != 'Quantity':
                    Quantity[i] = int(Quantity[i])
    for i in range(0, len(Description)):
        Description[i] = Description[i].replace(",", ";")
    csvout = zip(ProductID,Description,Quantity)
    resultFile = open(outfile,'wb')
    wr = csv.writer(resultFile, dialect='excel')
    for row in csvout:
        wr.writerow(row)
    resultFile.close()
    return None

def a2aFinalizedCircuits(inputfile):
    outfile = path_join(processDir, 'A2AFinalizedCircuits.csv')
    wb = open_workbook(inputfile)
    sheet = wb.sheet_by_name('Circuits')

    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    for i in range(0, number_of_columns):
        colvalues = sheet.col_values(i, start_rowx=0, end_rowx= None)
        if colvalues[0] == 'Wavelength':
            waveLength = list(colvalues)
        elif colvalues[0] == 'From Loc':
            fromLoc = list(colvalues)
        elif colvalues[0] == 'To Loc':
            toLoc = list(colvalues)
        elif colvalues[0] == 'Signal Rate':
            signalRate = list(colvalues)
    outList = []


    for i in range(1, len(waveLength)):
        waveL = waveLength[i].split(",")
        wLength = str(waveL[0]+str(waveL[1]))
        fromsiteandSide = fromLoc[i].split(".")
        tositeandSide = toLoc[i].split(".")
        info = (str(wLength)+','+str(fromsiteandSide[0])+','+str(fromsiteandSide[1])+','+str(tositeandSide[0])+','+str(tositeandSide[1])+','+str(signalRate[i]))

       # if info not in outList:
        outList.append(info)


    resultFile = open(outfile,'wb')
    #wr = csv.writer(resultFile, dialect='excel')
    resultFile.write('Wavelength,FromSite,FromSide,ToSite,ToSide,SignalRate\n')
    for row in outList:
        resultFile.write(row+'\n')

    resultFile.close()
    return None


def fibresDialog(inputfile):
    outfile = path_join(processDir, 'Fibres Dialog.csv')
    wb = open_workbook(inputfile)
    sheet = wb.sheet_by_name('Fibres')

    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    for i in range(0, number_of_columns):
        colvalues = sheet.col_values(i, start_rowx=2, end_rowx= None)
        if colvalues[0] == 'Name':
            listName = list(colvalues)
        elif colvalues[0] == 'Src.':
            Source = list(colvalues)
        elif colvalues[0] == 'Dst.':
            Destination = list(colvalues)
        elif colvalues[0] == 'Type':
            fiberType = list(colvalues)
        elif colvalues[0] == 'Length':
            fiberLength = list(colvalues)
        elif colvalues[0] == 'Loss SOL':
            lossSol = list(colvalues)
        elif colvalues[0] == 'Loss EOL':
            lossEol = list(colvalues)
        elif colvalues[0] == 'CD C-Band':
            cdcband = list(colvalues)
        elif colvalues[0] == 'CD L-Band':
            cdlband = list(colvalues)
        elif colvalues[0] == 'PMD':
            pmd = list(colvalues)
        elif colvalues[0] == 'QD C-Band':
            qdcband = list(colvalues)
        elif colvalues[0] == 'QD L-Band':
            qdlband = list(colvalues)
        elif colvalues[0] == 'RD':
            rd = list(colvalues)

    resultFile = open(outfile,'wb')
    #wr = csv.writer(resultFile, dialect='excel')
    resultFile.write('Name,Src,Dst,Type,Length,LossSOL,LossEOL,CDCBand,CDLBand,PMD,QDCBand,QDLBand,RD\n')
    for i in range(1, len(listName)):
        if listName[i]:
            resultFile.write(str(listName[i])+','+str(Source[i])+','+str(Destination[i])+','+str(fiberType[i])+','+str(fiberLength[i])+','+str(lossSol[i])+','+str(lossEol[i])+','+str(cdcband[i])+','+str(cdlband[i])+','+str(pmd[i])+','+str(qdcband[i])+','+str(qdlband[i])+','+str(rd[i])+'\n')
    resultFile.close()
    return None

def trafficMatrix(inputfile):
    outfile = path_join(processDir, 'Traffic Matrix.csv')
    wb = open_workbook(inputfile)
    resultFile = open(outfile,'wb')
    #wr = csv.writer(resultFile, dialect='excel')
    resultFile.write('Demand,ClServType,ProtectionType,Service,SrcSite,DstSite,SrcCard,DstCard\n')
    worksheets = wb.sheet_names()

    for worksheet_name in worksheets:
        sheet = wb.sheet_by_name(worksheet_name)

        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        for i in range(0, number_of_columns):
            colvalues = sheet.col_values(i, start_rowx=2, end_rowx= None)
            if 'Demand' in colvalues[0]:
                demandValues = list(colvalues)
                serviceValues = sheet.col_values(i+1, start_rowx=2, end_rowx= None)
            elif colvalues[0] == 'Cl. Serv. Type':
                clServType = list(colvalues)
            elif colvalues[0] == 'Protection Type':
                ProtectionType = list(colvalues)
            elif colvalues[0] == 'Service_*':
                serVice = list(colvalues)
            elif colvalues[0] == 'Src Site':
                srcSite = list(colvalues)
            elif colvalues[0] == 'Dst Site':
                dstSite = list(colvalues)
            elif colvalues[0] == 'Src Card':
                srcCard = list(colvalues)
            elif colvalues[0] == 'Dst Card':
                dstCard = list(colvalues)

        for i in range(1, len(demandValues)):
            if demandValues[i]:
                for j in range (i+1, len(serviceValues)):
                    if demandValues[j]:
                        break
                    if not demandValues[j]:
                        if 'Service_' in (serviceValues[j]):
                            resultFile.write(demandValues[i]+','+clServType[i]+','+ProtectionType[i]+','+serviceValues[j]+','+srcSite[j+1]+','+dstSite[j+1]+','+srcCard[j+1]+','+dstCard[j+1]+'\n')
                            #print demandValues[i], serVice[j]

def projData(infile):

    outfile = path_join(processDir, 'proj_data.csv')
    resultFile = open(outfile,'wb')
    inputfile = infile
    keywordlist = []

    with open(inputfile) as f:
        resultFile.write('Variable|Content\n')
        for line in f:
            result = re.search('<(.*)>', line)
            if result:
                keywordlist.append(result.group(1))

    with open (inputfile, "r") as myfile:
        data=myfile.read().replace('\n', '')

    for keyword in keywordlist:
        content = re.search(r'<'+keyword+'>=(.*)', data)
        if content:
            variables = content.group(1)
            variables = variables.lstrip()
            variable = variables.split('<')
            if variable[0] != '':
                resultFile.write(keyword+'|'+variable[0]+'\n')
            else:
                resultFile.write(keyword+'| N/A\n')

def layoutData():
    outfile = path_join(processDir, 'Layout.csv')
    layoutfiles = os.listdir(layoutDir)
    resultFile = open(outfile,'wb')
    resultFile.write('SiteName,Position,MaxPowerConsumption,TypicalPowerConsumption,UnitWeights\n')
    for file in layoutfiles:
        sitename = str(file).split('.')
        if '_' in sitename:
            sitename = str(sitename[0]).split('_')
            sitename = sitename[1]
        else:
            sitename = str(sitename[0])
        wb = open_workbook(path_join(layoutDir, file))
        worksheets = wb.sheet_names()

        for worksheet_name in worksheets:
            sheet = wb.sheet_by_name(worksheet_name)
            number_of_rows = sheet.nrows
            number_of_columns = sheet.ncols

            for i in range(0, number_of_columns):
                colvalues = sheet.col_values(i, start_rowx=2, end_rowx= None)
                if 'Name' in colvalues[0]:
                    names = list(colvalues)
                elif 'Position' in colvalues[0]:
                    positions = list(colvalues)
                elif 'Max power consumption' in colvalues[0]:
                    maxPower= list(colvalues)
                elif 'Typical power consumption' in colvalues[0]:
                    typicalPower= list(colvalues)
                elif 'Unit Weights' in colvalues[0]:
                    unitWeights = list(colvalues)

            for i in range(1, len(names)):
                if names[i]:
                    resultFile.write(sitename+','+str(positions[i])+','+str(maxPower[i])+','+str(typicalPower[i])+','+str(unitWeights[i])+'\n')


def table41():
    df = pd.read_csv(path_join(processDir, "nodesetup-sides.csv"))
    sourceSite = df.descr
    sourceSide = df.side_name
    dstSite = df.connected_to_node
    dstSide = df.connected_to_side_name
    outputDf = pd.DataFrame(columns=('Source Site','Source Side','Destination Site','Destination Side'))
    dfloc = 0
    for item in range (0, len(sourceSite)):
        Site = sourceSite[item].split('_')
        sSite = Site[0]
        sSide = sourceSide[item]
        dSite = dstSite[item]
        dSide = dstSide[item]
        outputDf.loc[dfloc] = [sSite,sSide,dSite,dSide]
        dfloc = dfloc + 1
    sortedOutputDf = outputDf.sort(["Source Site","Source Side"], ascending=[True, True])
    sortedOutputDf.to_csv(path_join(reportDir, 'table4.1.csv'), sep=',', index = False)

def table42():

    df = pd.read_csv(path_join(processDir,"nodesetup-ppms.csv"))
    outputDf = pd.DataFrame(columns=('Site','Types of Traffic'))
    dfloc = 0
    site = df.descr
    ppmporttype = df.ppmporttype
    uniqueSite = set(site)
    for item in uniqueSite:
        uSite = item.split('_')
        uSite = uSite[0]
        trafficList = []
        for i in range(0, len(site)):
            if site[i] == item:
                trafficList.append(str(ppmporttype[i]))
        while 'nan' in trafficList: trafficList.remove('nan')
        traffic = set(trafficList)
        trafficTypes = ';'.join(traffic)
        outputDf.loc[dfloc] = [uSite,trafficTypes]
        dfloc = dfloc+1
    sortedOutputDf = outputDf.sort(["Site"], ascending=[True])
    sortedOutputDf.to_csv(path_join(reportDir,'table4.2.csv'), sep=',', index = False)


def table56():
    df = pd.read_csv(path_join(processDir, "Layout.csv"))
    sName = df.SiteName
    position = df.Position
    mPower = df.MaxPowerConsumption
    tPower = df.TypicalPowerConsumption
    uniqueSName = set(sName)
    outputDf = pd.DataFrame(columns=('Site Name','Number of Racks','Max Power consumption (W)','Typical Power Consumption (W)'))
    dfloc = 0
    for site in uniqueSName:
        NoOfRacks = 0
        MaxPC = 0
        TypPC = 0
        for i in range(0, len(sName)):
            if sName[i]== site:
                NoOfRacks = NoOfRacks + 1
                MaxPC = MaxPC + float(mPower[i])
                TypPC = TypPC + float(tPower[i])
        MaxPC = ("{0:.2f}".format(MaxPC))
        TypPC = ("{0:.2f}".format(TypPC))
        outputDf.loc[dfloc] = [site,str(NoOfRacks),MaxPC,TypPC]
        dfloc = dfloc+1
        #outFile.write(site+','+str(NoOfRacks)+','+str(MaxPC)+','+str(TypPC)+'\n')
    sortedOutputDf = outputDf.sort(["Site Name"], ascending=[True])
    sortedOutputDf.to_csv(path_join(reportDir, 'table5.6.csv'), sep=',', index = False)

def table61():
    df = pd.read_csv(path_join(processDir, 'Traffic Matrix.csv'))
    sortedData = df.sort("Demand")
    demand = sortedData.Demand
    source = sortedData.SrcSite
    dst = sortedData.DstSite
    srcCard = sortedData.SrcCard
    dstCard = sortedData.DstCard
    uniqueDemand = set(demand)
    outputDf = pd.DataFrame(columns=('Service Demand','Source Site','Destination Site','Source Card', 'Destination Card' ))
    dfloc = 0
    for uitem in uniqueDemand:

        for item in range (0, len(demand)):
            if demand[item] == uitem:
                sourceSite = source[item]
                dstsite = dst[item]
                sourceCard = srcCard[item]
                destCard = dstCard[item]
        outputDf.loc[dfloc] = [uitem, sourceSite, dstsite, sourceCard, destCard ]
        dfloc = dfloc + 1
    sortedOutputDf = outputDf.sort("Service Demand")
    sortedOutputDf.to_csv(path_join(reportDir, 'table6.1.csv'), sep=',', index = False)


def table66():
    shutil.copy(path_join(processDir, 'Bill of Material.csv'), path_join(reportDir, 'table6.6.csv'))

def table71():

    df = pd.read_csv(path_join(processDir, "A2AFinalizedCircuits.csv"))
    fromSite = df.FromSite
    fromSide = df.FromSide
    toSite = df.ToSite
    toSide = df.ToSide
    wL = df.Wavelength
    outputDf = pd.DataFrame(columns=('From Site','From Side','To Site','To Side','Wavelength Count','Circuit Count'))
    dfloc = 0
    concatList = []
    concatList = fromSite+fromSide+toSite+toSide
    uConcatList = set(concatList)
    for uitem in uConcatList:
        waveLength = []
        circuitCount = 0
        for item in range (0, len(fromSite)):
            if concatList[item] == uitem:
                fSite = fromSite[item]
                fSide = fromSide[item]
                tSite = toSite[item]
                tSide = toSide[item]
                waveLength.append(wL[item])
                circuitCount = circuitCount + 1
        uwaveLength = set(waveLength)
        outputDf.loc[dfloc] = [fSite, fSide, tSite, tSide, str(len(uwaveLength)), str(circuitCount)]
        dfloc = dfloc + 1
    sortedOutputDf = outputDf.sort(["From Site","From Side"], ascending=[True, True])
    sortedOutputDf.to_csv(path_join(reportDir, 'table7.1.csv'), sep=',', index = False)

def table72():

    df = pd.read_csv(path_join(processDir, "Traffic Matrix.csv"))
    demand = df.Demand
    source = df.SrcSite
    dst = df.DstSite
    clServType = df.ClServType
    protType = df.ProtectionType
    uniqueDemand = set(demand)
    outputDf = pd.DataFrame(columns=('Source','Destination','Number of Services','Client Service Type','Protection Type'))
    dfloc = 0
    for uitem in uniqueDemand:
        serViceCount = 0
        for item in range (0, len(demand)):
            if demand[item] == uitem:
                serViceCount = serViceCount + 1
                srcsite = source[item]
                dstsite = dst[item]
                clservtype = clServType[item]
                protectionType = protType[item]
        outputDf.loc[dfloc] = [srcsite, dstsite, str(serViceCount), clservtype, protectionType]
        dfloc = dfloc + 1
    sortedOutputDf = outputDf.sort(["Source","Destination"], ascending=[True, True])
    sortedOutputDf.to_csv(path_join(reportDir, 'table7.2.csv'), sep=',', index = False)

def table81():
    df = pd.read_csv(path_join(processDir, "Fibres Dialog.csv"))
    outFile = open(path_join(reportDir, 'table8.1.csv'),'wb')
    outputDf = pd.DataFrame(columns=('Source','Destination','Fiber Type', 'Fiber Length','Fiber Loss', 'CD C-Band','CD L-Band','PMD'))
    Source = df.Src
    Dst = df.Dst
    Ftype = df.Type
    Flength = df.Length
    Loss = df.LossSOL
    CDC = df.CDCBand
    CDL = df.CDLBand
    PMD = df.PMD
    dfloc = 0
    for i in range (0, len(Source)):
        outputDf.loc[dfloc] = [Source[i], Dst[i], Ftype[i], Flength[i], Loss[i], CDC[i], CDL[i], PMD[i]]
        dfloc = dfloc + 1
    sortedOutputDf = outputDf.sort(["Source","Destination"], ascending=[True, True])
    sortedOutputDf.to_csv(path_join(reportDir, 'table8.1.csv'), sep=',', index = False)
#===============================================================================
def paramsFile():
    df = pd.read_csv(path_join(processDir, "proj_data.csv"), sep="|")
    outputFile = path_join(reportDir, 'Params.csv')
    try:
        outFile = open(outputFile,'w')
    except IOError:
        print "[ERROR] Cound't write"+outputFile
        sys.exit()
    Variable = df.Variable
    Content = df.Content
    #dfloc = 0
    paramDict = {}
    for i in range(0,len(Variable)):
        param = str(Variable[i])
        value = str(Content[i])
        if param[-1]==' ':
            param = param[:-1]
        if param and param[0]==' ':
            param = param[1:]
        if value[-1]==' ':
            value = value[:-1]
        if value and value[0] ==' ':
            value = value[1:]
        if not value:
            value = 'N/A'
        if param == 'nan':
            param = 'N/A'
        if value == 'nan':
            value = 'N/A'
        if param == 'customer':
            param = 'customer'
        param = '<'+param+'>|'
        outFile.write(param+value+'\n')
    reportDate = '<reportDate>|'
    todayDate = (time.strftime('%B %d, %Y'))
    outFile.write(reportDate+str(todayDate)+'\n')


def main():
    tree = ET.parse(path_join(collectorDir, 'neupdate.xml'))
    root = tree.getroot()
    anpCSV(root)
    tagtoCSV(root,'parameters')
    tagtoCSV(root,'ppms')
    tagtoCSV(root,'patchcords')
    tagtoCSV(root,'sides')
    tagtoCSV(root,'ans')

    mypath = collectorDir
    onlyfiles = [ f for f in listdir(mypath) if isfile(join(mypath,f)) ]
    for filename in onlyfiles:
        file = filename.lower()
        file = file.replace(" ", "")
        if file == 'billofmaterial.xls':
            billofMaterial(mypath+'/'+filename)
        elif file == 'a2afinalizedcircuits.xls':
            a2aFinalizedCircuits(mypath+'/'+filename)
        elif file == 'fibresdialog.xls':
            fibresDialog(mypath+'/'+filename)
        elif file == 'trafficmatrix.xls':
            trafficMatrix(mypath+'/'+filename)
        elif file == 'proj_data.txt':
            projData(mypath+'/'+filename)
    layoutData()

    table41()
    table42()
    table56()
    table61()
    table66()
    table71()
    table72()
    table81()
    paramsFile()

if __name__ == '__main__':
        main()


