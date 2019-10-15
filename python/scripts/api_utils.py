import TemplateCenterConf as TCenterConf
import TemplateCenterUtility as TCenterUtility
from urllib import request
import urllib
import json
import os, sys
import uno,unohelper



def checkDiff(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    TCenterUtility.checkDiff(dialog)
    TCenterUtility.Msgbox("已同步範本資訊。")

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
    with open(TCenterConf.getServerSettingPath(), "w") as serverSetting:
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

def syncTemplates(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    TCenterUtility.syncTemplates(dialog)

def test(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel
    with open(TCenterConf.getProjectDataPath() + "color.txt", "w") as outFile:
        outFile.write(str(oGridModel.RowBackgroundColors))

def fileChange(*args, **kwargs):
    try:
        projectDataPath = TCenterConf.getProjectDataPath()
        with open(projectDataPath + "test.json", "w") as fout:
            fout.write("test\n")
    except:
        pass

def getButton(*args, **kwargs):
    url = TCenterConf.getAPIAddress_List()
    res = urllib.request.urlopen(url)
    data = json.loads(res.read().decode())
    theSource = args[0].Source
    theContext = theSource.getContext()
    tmpList = theContext.getControl("TemplateInfo")
    for it in data['templates']:
        tmpList.Model.insertItemText(0, it)
        tmpList.Model.setItemData(0, it)

def sycnCheckYes(*args, **kwargs):
    theSource = args[0].Source
    dialog = theSource.getContext()
    jsonData = {}
    jsonData['sync'] = True
    with open(TCenterConf.getSyncCheckResult(), "w") as syncCheckResult:
        json.dump(jsonData, syncCheckResult)
    dialog.endExecute()

