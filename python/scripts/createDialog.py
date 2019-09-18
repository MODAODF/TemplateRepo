import uno

templateDialog = None

def createDgConfig(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplatesMarket.Config?location=application")
    dialog.execute()
    dialog.dispose()

def createDgCommandList(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplatesMarket.CommandList?location=application")
    dialog.execute()
    dialog.dispose()

def createDgTemplate(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplatesMarket.TemplateList?location=application")
    #dialog.execute()
    dialog.EnableVisible = True
    dialog.dispose()


def createDgSetting(*args, **kwargs):
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog("vnd.sun.star.script:TemplatesMarket.Setting?location=application")
    dialog.execute()