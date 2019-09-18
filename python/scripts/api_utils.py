import TemplateRepoConf as TRConf
from urllib import request
import urllib
import json

def getCurrentData(*args, **kwargs):
    try:
        url = "http://192.168.3.11:9980/lool/templaterepo/list"
        res = urllib.request.urlopen(url)
        with open(r"C:\\Users\\Tommy\\AppData\\NDCFILE\\templatesInfo.json", "w") as outFile:
            json.dump(json.loads(res.read().decode()), outFile)
    except:
        return


def checkDiff(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    clearGrid(dialog)

    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel
    try:
        with open(TRConf.InfoFilePath, "r") as jsonFile:
            oldTemplateInfo = json.load(jsonFile)
    except:
        oldTemplateInfo = {}

    
    try:
        url = "http://192.168.3.11:9980/lool/templaterepo/list"
        res = urllib.request.urlopen(url)
        newTemplateInfo = json.loads(res.read().decode())
    except:
        oDataModel.addRow("X", ["請確認網服務","或是","伺服器是否有正確設定"])
        return

    rowCount = 1
    colorArray = []
    oldMark = int("ffffff", 16)
    newMark = int("ffb5b5", 16)

    diffJson = {}
    for nDepart in newTemplateInfo:
        if nDepart in oldTemplateInfo:
            for nTemplate in newTemplateInfo[nDepart]:
                isNew = True
                for oTemplate in oldTemplateInfo[nDepart]:
                    if nTemplate['docname'] == oTemplate['docname']:
                        if nTemplate['uptime'] == oTemplate['uptime']:
                            isNew = False
                            break
                
                rowData = []
                rowData.append(nDepart)
                rowData.append(nTemplate['docname'])
                rowData.append(nTemplate['extname'])
                rowData.append(nTemplate['uptime'])

                oDataModel.addRow(str(rowCount), rowData)
                if isNew:
                    colorArray.append(newMark)

                    if nDepart not in diffJson:
                        diffJson[nDepart] = []
                    diffJson[nDepart].append(nTemplate)
                else:
                    colorArray.append(oldMark)
                rowCount += 1
                        
        else:
            diffJson[nDepart]=[]
            for template in newTemplateInfo[nDepart]:
                rowData = []
                rowData.append(nDepart)
                rowData.append(template['docname'])
                rowData.append(template['extname'])
                rowData.append(template['uptime'])

                oDataModel.addRow(str(rowCount), rowData)
                colorArray.append(newMark)
                rowCount += 1

                diffJson[nDepart].append(template)
    oGridModel.RowBackgroundColors = colorArray

    with open(r"C:\\Users\\Tommy\\AppData\\NDCFILE\\diffInfo.json", "w") as outFile:
        json.dump(diffJson, outFile)


def insertGrid(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel
    try:
        url = "http://192.168.3.11:9980/lool/templaterepo/list"
        res = urllib.request.urlopen(url)
        jsonData = json.loads(res.read().decode())
        with open(r"C:\\Users\\Tommy\\AppData\\NDCFILE\\templatesInfo.json", "w") as outFile:
            json.dump(jsonData, outFile)
        rowCount = 1
        for department in jsonData:
            for template in jsonData[department]:
                rowData = []
                rowData.append(department)
                rowData.append(template['docname'])
                rowData.append(template['extname'])
                rowData.append(template['uptime'])

                oDataModel.addRow(str(rowCount), rowData)
                rowCount += 1

        colorArray = []
        for x in range(oDataModel.RowCount):
            if x < 10:
                colorArray.append(int("ffb5b5", 16))
            else:
                colorArray.append(int("ffffff", 16))
        oGridModel.RowBackgroundColors = colorArray

    except:
        oDataModel.addRow("X", ["請檢察網路問題 或是 伺服器設定"])


def clearButton(*args, **kwargs):
    theSource = args[0].Source
    theContext = theSource.getContext()
    gridControl = theContext.getControl("ListGrid")
    oGridModel = gridControl.Model
    oDataModel = oGridModel.GridDataModel
    oDataModel.removeAllRows()
    oGridModel.RowBackgroundColors = []

def clearGrid(dialog):
    gridControl = dialog.getControl("ListGrid")
    oGridModel = gridControl.Model
    oDataModel = oGridModel.GridDataModel
    oDataModel.removeAllRows()
    oGridModel.RowBackgroundColors = []

def syncTemplates(*args, **kwargs):
    """
    TODO
    Download the diff file from remote 
    """
    try:
        url = "http://192.168.3.11:9980/lool/templaterepo/download"
        res = urllib.request.urlopen(url)
        with open(r"C:\\Users\\Tommy\\AppData\\NDCFILE\\templates.zip", "wb") as outFile:
            outFile.write(res.read())
    except:
        return


def getButton(*args, **kwargs):
    url = "http://192.168.3.11:9980/lool/market/list"
    res = urllib.request.urlopen(url)
    data = json.loads(res.read().decode())
    theSource = args[0].Source
    theContext = theSource.getContext()
    tmpList = theContext.getControl("TemplateInfo")
    for it in data['templates']:
        tmpList.Model.insertItemText(0, it)
        tmpList.Model.setItemData(0, it)
