import uno
import TemplateCenterConf as TRConf

def Msgbox(Mensaje):
    ctx = uno.getComponentContext()    
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx )
    oDoc = desktop.getCurrentComponent()
    oSM = uno.getComponentContext().getServiceManager()
    oToolkit = oSM.createInstance("com.sun.star.awt.Toolkit")
    sTipo = "infobox"
    sTitulo = "Hint"
    botones = 1
    if hasattr(oDoc, "getCurrentController"):
        oParentWin = oDoc.getCurrentController().getFrame().getContainerWindow()
        oMsgBox = oToolkit.createMessageBox(
            oParentWin, sTipo, botones, sTitulo, Mensaje)
        oMsgBox.execute()
    else:
        oParentWin = oDoc.getFrame().getContainerWindow()
        oMsgBox = oToolkit.createMessageBox(
            oParentWin, sTipo, botones, sTitulo, Mensaje)
        oMsgBox.execute()
    
    return None

def write2file(filename, msg):
    with open(TRConf.getProjectDataPath() + filename, "w") as fp:
        for item in msg:
            fp.write(str(item)+"\n")
        fp.close()


