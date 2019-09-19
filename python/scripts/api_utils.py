import TemplateRepoConf as TRConf
from urllib import request
import urllib
import json
import os, sys

def getCurrentData(*args, **kwargs):
    try:
        url = TRConf.getAPIAddress_List()
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
        with open(TRConf.getTemplateInfoPath(), "r") as jsonFile:
            oldTemplateInfo = json.load(jsonFile)
    except:
        oldTemplateInfo = {}

    
    try:
        url = TRConf.getAPIAddress_List()
        res = urllib.request.urlopen(url)
        newTemplateInfo = json.loads(res.read().decode())
    except:
        oDataModel.addRow("X", ["請確認網服務","或是","伺服器是否有正確設定",TRConf.getAPIAddress_List()])
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

    renderInfoLabel(dialog, newTemplateInfo)
    with open(TRConf.getDiffInfoPath(), "w") as outFile:
        json.dump(diffJson, outFile)

def renderSyncResult(dialog, jsonData):
    clearGrid(dialog)
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel
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

    colorArray = [int("ffffff", 16)]
    oGridModel.RowBackgroundColors = colorArray


def renderInfoLabel(dialog, jsonData):
    oGridModel = dialog.getControl("ListGrid").Model
    TotalLabel = dialog.getControl("Total")
    ODTLabel = dialog.getControl("ODT")
    ODSLabel = dialog.getControl("ODS")
    NewLabel = dialog.getControl("New")

    totalCount = 0
    ODTCount = 0
    ODSCount = 0
    NewCount = 0

    for depart in jsonData:
        for template in jsonData[depart]:
            totalCount += 1
            extname = template['extname']
            if extname == 'ott':
                ODTCount += 1
            elif extname == "ots" :
                ODSCount += 1
    for color in oGridModel.RowBackgroundColors:
        if color == int("ffb5b5", 16):
            NewCount += 1
    TotalLabel.Model.Text = str(totalCount)
    ODTLabel.Model.Text = str(ODTCount)
    ODSLabel.Model.Text = str(ODSCount)
    NewLabel.Model.Text = str(NewCount)

def emptyDiff():
    with open(TRConf.getDiffInfoPath(), "w") as outFile:
        json.dump({}, outFile)
        outFile.close()

def saveServerSetting(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    ipAddr = dialog.getControl("ServerAddress").Model
    port = dialog.getControl("Port").Model
    httpMethod = dialog.getControl("HTTP")
    jsonData = {}
    jsonData['ServerAddress'] = ipAddr.Text
    jsonData['Port'] = port.Text
    jsonData['httpMethod'] = httpMethod.State
    with open(TRConf.getServerSettingPath(), "w") as serverSetting:
        json.dump(jsonData, serverSetting)
    dialog.endExecute()

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
    Download the diff file from remote and extract
    """
    projectDataPath = TRConf.getProjectDataPath()
    try:
        os.remove(projectDataPath + "error_urlopen.json")
    except:
        pass
    theSource = args[0].Source
    dialog = theSource.getContext()
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel

    try:

        jsonData = {}
        headers={'Content-Type':'application/json'}
        
        with open(TRConf.getDiffInfoPath(), "r") as outFile:
            jsonData = json.load(outFile)
        bindata = json.dumps(jsonData)
        bindata = bindata.encode('utf-8')
        try:
            req = urllib.request.Request(TRConf.getAPIAddress_Sync(), data=bindata, headers=headers)
            res = urllib.request.urlopen(req)
        except:
            oDataModel.addRow("X", ["請確認網服務","或是","伺服器是否有正確設定"])
            return
        
        with open(projectDataPath + "sync.zip", "wb") as zipFile:
            zipFile.write(res.read())
            zipFile.close()

        # 解壓縮檔案到使用者端的 Templates
        import zipfile
        zf = zipfile.ZipFile(projectDataPath + "sync.zip")
        zf.extractall(TRConf.getUserTemplatePath())
        emptyDiff()
        try:
            url = TRConf.getAPIAddress_List()
            res = urllib.request.urlopen(url)
            jsonData = json.loads(res.read().decode())
            with open(TRConf.getTemplateInfoPath(), "w") as outFile:
                json.dump(jsonData, outFile)
            renderSyncResult(dialog, jsonData)
            renderInfoLabel(dialog, jsonData)
        except:
            oDataModel.addRow("X", ["請確認網服務","或是","伺服器是否有正確設定"])
            return
        
    except:
        with open(projectDataPath + "error_urlopen.json", "w") as fout:
            fout.write("error\n")

def test(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel
    with open(TRConf.getProjectDataPath() + "color.txt", "w") as outFile:
        outFile.write(str(oGridModel.RowBackgroundColors))

def fileChange(*args, **kwargs):
    try:
        projectDataPath = TRConf.getProjectDataPath()
        with open(projectDataPath + "test.json", "w") as fout:
            fout.write("test\n")
    except:
        pass

def getButton(*args, **kwargs):
    url = TRConf.getAPIAddress_List()
    res = urllib.request.urlopen(url)
    data = json.loads(res.read().decode())
    theSource = args[0].Source
    theContext = theSource.getContext()
    tmpList = theContext.getControl("TemplateInfo")
    for it in data['templates']:
        tmpList.Model.insertItemText(0, it)
        tmpList.Model.setItemData(0, it)
