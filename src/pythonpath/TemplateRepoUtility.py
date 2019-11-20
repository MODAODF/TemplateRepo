import uno
import TemplateRepoConf as TRepoConf
import json
import ssl
from urllib import request
import urllib

def makeReq(url):
    context = ssl._create_unverified_context()
    return urllib.request.urlopen(url, timeout=2, context=context)

def Msgbox(Message):
    ctx = uno.getComponentContext()    
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx )
    oDoc = desktop.getCurrentComponent()
    oSM = uno.getComponentContext().getServiceManager()
    oToolkit = oSM.createInstance("com.sun.star.awt.Toolkit")
    sTipo = "infobox"
    sTitulo = "Hint"
    botones = 1
    if hasattr(oDoc, "getCurrentController"):
        oParentWin = oDoc.getCurrentController().getFrame().getContainerWindow()
        oMsgBox = oToolkit.createMessageBox(
            oParentWin, sTipo, botones, sTitulo, Message)
        oMsgBox.execute()
    else:
        oParentWin = oDoc.getFrame().getContainerWindow()
        oMsgBox = oToolkit.createMessageBox(
            oParentWin, sTipo, botones, sTitulo, Message)
        oMsgBox.execute()
    
    return None

def write2file(filename, msg):
    with open(TRepoConf.getProjectDataPath() + filename, "w") as fp:
        for item in msg:
            fp.write(str(item)+"\n")
        fp.close()

def sycnCheckNo():
    jsonData = {}
    jsonData['sync'] = False
    with open(TRepoConf.getSyncCheckResult(), "w") as syncCheckResult:
        json.dump(jsonData, syncCheckResult)

def readSycnCheck():
    with open(TRepoConf.getSyncCheckResult(), "r") as syncCheckResult:
        result = json.load(syncCheckResult)
    if result['sync']:
        return True
    else:
        return False

def createDgSyncCheck():
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateRepo.SyncCheck?location=application")
    dialog.execute()

def checkDiff(sDialog):
    dialog = sDialog
    clearGrid(dialog)

    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel

    try:
        with open(TRepoConf.getTemplateInfoPath(), "r") as jsonFile:
            oldTemplateInfo = json.load(jsonFile)
    except:
        oldTemplateInfo = {}

    
    try:
        url = TRepoConf.getAPIAddress_List()
        res = makeReq(url)
        newTemplateInfo = json.loads(res.read().decode())
    except:
        message = "請確認伺服器設定\n\n當前伺服器位置設定\n\n" + TRepoConf.getServerAddress()
        Msgbox(message)
        return False

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
    with open(TRepoConf.getDiffInfoPath(), "w") as outFile:
        json.dump(diffJson, outFile)
    return True
    
def checkConnection():
    try:
        url = TRepoConf.getAPIAddress_List()
        res = makeReq(url)
        res = None
    except:
        message = "無法連線至伺服器\n\n請確認伺服器設定或網路連線\n\n當前伺服器位置設定\n\n" + TRepoConf.getServerAddress()
        Msgbox(message)
        return False
    return True

def isDiffExist():
    diffJson={}
    try:
        with open(TRepoConf.getDiffInfoPath(), "r") as jsonFile:
            diffJson = json.load(jsonFile)
    except:
        return False
    return True

def syncTemplates(sDialog):
    """
    Download the diff file from remote and extract
    """
    # 二次確認
    if not isDiffExist():
        Msgbox("請先同步範本資訊。")
        return
    
    sycnCheckNo()
    createDgSyncCheck()
    if not readSycnCheck():
        return 

    projectDataPath = TRepoConf.getProjectDataPath()
    try:
        os.remove(projectDataPath + "error_urlopen.json")
    except:
        pass
    
    dialog = sDialog 
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel

    try:

        jsonData = {}
        headers={'Content-Type':'application/json'}
        
        with open(TRepoConf.getDiffInfoPath(), "r") as outFile:
            jsonData = json.load(outFile)

        bindata = json.dumps(jsonData)
        bindata = bindata.encode('utf-8')
        try:
            req = urllib.request.Request(TRepoConf.getAPIAddress_Sync(), data=bindata, headers=headers)
            res = makeReq(req)
            if not jsonData:
                Msgbox("已同步至最新版本。")
                return
        except:
            message = "請確認伺服器設定\n\n當前伺服器位置設定\n\n" + TRepoConf.getServerAddress()
            Msgbox(message)
            return
        
        with open(projectDataPath + "sync.zip", "wb") as zipFile:
            zipFile.write(res.read())
            zipFile.close()

        # 解壓縮檔案到使用者端的 Templates
        import zipfile
        zf = zipfile.ZipFile(projectDataPath + "sync.zip")
        zf.extractall(TRepoConf.getUserTemplatePath())
        emptyDiff()
        try:
            url = TRepoConf.getAPIAddress_List()
            res = makeReq(url)
            jsonData = json.loads(res.read().decode())
            with open(TRepoConf.getTemplateInfoPath(), "w") as outFile:
                json.dump(jsonData, outFile)
            renderSyncResult(dialog, jsonData)
            renderInfoLabel(dialog, jsonData)
            Msgbox("同步完成")
        except:
            message = "網路連線發生問題，請確認網路連線\n\n當前伺服器位置設定\n\n" + TRepoConf.getServerAddress()
            Msgbox(message)
            return
        
    except:
        with open(projectDataPath + "error_urlopen.json", "w") as fout:
            fout.write("error\n")

def clearGrid(dialog):
    gridControl = dialog.getControl("ListGrid")
    oGridModel = gridControl.Model
    oDataModel = oGridModel.GridDataModel
    oDataModel.removeAllRows()
    oGridModel.RowBackgroundColors = []

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

def emptyDiff():
    with open(TRepoConf.getDiffInfoPath(), "w") as outFile:
        json.dump({}, outFile)
        outFile.close()

def renderInfoLabel(dialog, jsonData):
    oGridModel = dialog.getControl("ListGrid").Model
    TotalLabel = dialog.getControl("Total")
    ODTLabel = dialog.getControl("ODT")
    ODSLabel = dialog.getControl("ODS")
    ODPLabel = dialog.getControl("ODP")
    NewLabel = dialog.getControl("New")

    totalCount = 0
    ODTCount = 0
    ODSCount = 0
    ODPCount = 0
    NewCount = 0

    for depart in jsonData:
        for template in jsonData[depart]:
            totalCount += 1
            extname = template['extname']
            if extname == 'ott':
                ODTCount += 1
            elif extname == "ots":
                ODSCount += 1
            elif extname == "otp":
                ODPCount += 1
    for color in oGridModel.RowBackgroundColors:
        if color == int("ffb5b5", 16):
            NewCount += 1
    TotalLabel.Model.Text = str(totalCount)
    ODTLabel.Model.Text = str(ODTCount)
    ODSLabel.Model.Text = str(ODSCount)
    ODPLabel.Model.Text = str(ODPCount)
    NewLabel.Model.Text = str(NewCount)