import sys, os
import json
InfoFilePath = r"templatesInfo.json"
DiffFilePath = r"diffInfo.json"
ServerSettingPath = r"server.json"
syncCheckResultPath = r"sync_check.json"
ProjectName  = "TemplateCenter.oxt"

def getServerAddress():
    try:
        serverSetting = {}
        with open(getServerSettingPath(), "r") as setting:
           serverSetting = json.load(setting)
        apiBaseAddr = ""
        if serverSetting['httpMethod']:
            apiBaseAddr += "http://"
        else:
            apiBaseAddr += "https://"
        apiBaseAddr += (serverSetting['ServerAddress'] + ":" + serverSetting['Port'] + "/")
        return apiBaseAddr
    except:
        return "http://127.0.0.1:22/"
    
def getAPIAddress_List():
    baseAddress = getServerAddress()
    apiAddress  = baseAddress + 'lool/templaterepo/list'
    return apiAddress 

def getAPIAddress_download():
    baseAddress = getServerAddress()
    apiAddress  = baseAddress + 'lool/templaterepo/download'
    return apiAddress

def getAPIAddress_Sync():
    baseAddress = getServerAddress()
    apiAddress  = baseAddress + 'lool/templaterepo/sync'
    return apiAddress

def getProjectDataPath():
    global ProjectName
    outPath = ""
    for path in sys.path:
        if ProjectName in path:
            outPath = path
            break
    
    outPath = outPath.split(ProjectName)[0] + ProjectName + "\\runTimeData\\"
    return outPath

def getServerSettingPath():
    prDataPath = getProjectDataPath()
    global InfoFilePath
    return prDataPath + ServerSettingPath

def getTemplateInfoPath():
    prDataPath = getProjectDataPath()
    global InfoFilePath
    return prDataPath + InfoFilePath

def getDiffInfoPath():
    prDataPath = getProjectDataPath()
    global DiffFilePath
    return prDataPath + DiffFilePath

def getUserTemplatePath():
    global ProjectName
    outPath = ""
    for path in sys.path:
        if ProjectName in path:
            outPath = path
            break
    
    outPath = outPath.split("uno_packages")[0] + "template\\" 
    if not os.path.exists(outPath):
        os.mkdir(outPath)
    return outPath

def getSyncCheckResult():
    prDataPath = getProjectDataPath()
    global syncCheckResultPath
    return prDataPath + syncCheckResultPath