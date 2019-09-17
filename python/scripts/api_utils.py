import urllib
import json
TemplateDialog = None

def getCurrentData(*args, **kwargs):
    url = "http://192.168.3.11:9980/lool/templaterepo/list"
    res = urllib.request.urlopen(url)
    with open(r"C:\\Users\\Tommy\\AppData\\NDCFILE\\templatesInfo.json", "w") as outFile:
        json.dump(json.loads(res.read().decode()), outFile)

def syncDiff():
    pass


def syncTemplates(*args, **kwargs):
    url = "http://192.168.3.11:9980/lool/templaterepo/download"
    res = urllib.request.urlopen(url)
    with open(r"C:\\Users\\Tommy\\AppData\\NDCFILE\\templates.zip", "wb") as outFile:
        outFile.write(res.read())


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

def writeFile(*args, **kwargs):
    import os,sys
    if not os.path.exists(r'C:\\Users\\Tommy\\AppData\\NDCFILE'):
        os.mkdir(r'C:\\Users\\Tommy\\AppData\\NDCFILE')
    with open(r'C:\\Users\\Tommy\\AppData\\NDCFILE\\getButton.txt', 'w') as f :
        if os:
            f.write("args\n")
        if os:
            f.write("kwargs")
        f.write("a")