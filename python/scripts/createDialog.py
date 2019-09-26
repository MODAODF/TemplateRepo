import uno
import TemplateCenterConf as TCenterConf
import json
templateDialog = None

def createDgConfig(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateCenter.Config?location=application")
    dialog.execute()
    dialog.dispose()

def createDgCommandList(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateCenter.CommandList?location=application")
    dialog.execute()
    dialog.dispose()

def createDgTemplate(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateCenter.TemplateList?location=application")
    #dialog.execute()
    dialog.EnableVisible = True
    dialog.dispose()


def createDgSetting(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplateCenter.Setting?location=application")
    try:
        jsonData = {}
        with open(TCenterConf.getServerSettingPath(), "r") as serverSetting:
            jsonData = json.load(serverSetting)
        ipAddr = dialog.getControl("ServerAddress").Model
        port = dialog.getControl("Port").Model
        httpCheck = dialog.getControl("HTTP")
        httpsCheck = dialog.getControl("HTTPS")

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
        httpMethod.State = True
    dialog.execute()