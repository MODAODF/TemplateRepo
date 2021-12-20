import uno
import TemplateRepoConf as TRepoConf
import json
import uuid
templateDialog = None

def get_mac_address():
    return TRepoConf.getMAC()

def createDgConfig(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateRepo.Config?location=application")
    dialog.execute()
    dialog.dispose()

def createDgCommandList(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateRepo.CommandList?location=application")
    dialog.execute()
    dialog.dispose()

def createDgTemplate(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateRepo.TemplateList?location=application")
    #dialog.execute()
    dialog.EnableVisible = True
    dialog.dispose()

def createDgSetting(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateRepo.Setting?location=application")
    try:
        jsonData = {}
        with open(TRepoConf.getServerSettingPath(), "r") as serverSetting:
            jsonData = json.load(serverSetting)
        ipAddr = dialog.getControl("ServerAddress").Model
        port = dialog.getControl("Port").Model
        httpCheck = dialog.getControl("HTTP")
        httpsCheck = dialog.getControl("HTTPS")
        macaddr = dialog.getControl("MacAddress")
        macaddr.setText(get_mac_address())

        ipAddr.Text = jsonData['ServerAddress']
        port.Text = jsonData['Port']
        if jsonData['httpMethod']:
            httpCheck.State = True
            httpsCheck.State = False
        else:
            httpCheck.State = False
            httpsCheck.State = True
    except:
        ipAddr = dialog.getControl("ServerAddress").Model
        port = dialog.getControl("Port").Model
        httpMethod = dialog.getControl("HTTP")
        ipAddr.Text = "127.0.0.1"
        port.Text = "22"
        macaddr = dialog.getControl("MacAddress")
        macaddr.setText(get_mac_address())
        httpMethod.State = True
    dialog.execute()