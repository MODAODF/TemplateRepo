import sys
import os
import json
InfoFile = r"templatesInfo.json"
DiffFile = r"diffInfo.json"
ServerSettingFile = r"server.json"
syncCheckResultPath = r"sync_check.json"
ProjectName = "TemplateRepo"


def getProjectRootPath():
    global ProjectName
    outPath = ""
    for path in sys.path:
        if ProjectName in path and "oxt" in path:
            outPath = path
            break
    pathToken = outPath.split('\\')
    outPath = ""
    for token in pathToken:
        outPath += token
        outPath += "\\"
        if ProjectName in token:
            break
    return outPath


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
        apiBaseAddr += (serverSetting['ServerAddress'] +
                        ":" + serverSetting['Port'] + "/")
        return apiBaseAddr
    except:
        return "http://127.0.0.1:22/"


def getAPIAddress_List():
    baseAddress = getServerAddress()
    apiAddress = baseAddress + 'lool/templaterepo/list'
    return apiAddress


def getAPIAddress_download():
    baseAddress = getServerAddress()
    apiAddress = baseAddress + 'lool/templaterepo/download'
    return apiAddress


def getAPIAddress_Sync():
    baseAddress = getServerAddress()
    apiAddress = baseAddress + 'lool/templaterepo/sync'
    return apiAddress


def getProjectImagesPath():
    outPath = getProjectRootPath() + "images\\"
    return outPath


def getProjectDataPath():
    outPath = getProjectRootPath() + "runTimeData\\"
    return outPath


def getServerSettingPath():
    prDataPath = getProjectDataPath()
    return prDataPath + ServerSettingFile


def getTemplateInfoPath():
    prDataPath = getProjectDataPath()
    global InfoFile
    return prDataPath + InfoFile


def getDiffInfoPath():
    prDataPath = getProjectDataPath()
    global DiffFile
    return prDataPath + DiffFile


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
