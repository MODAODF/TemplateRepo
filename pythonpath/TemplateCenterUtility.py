import uno



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
    oParentWin = oDoc.getCurrentController().getFrame().getContainerWindow()
    oMsgBox = oToolkit.createMessageBox(
        oParentWin, sTipo, botones, sTitulo, Mensaje)
    oMsgBox.execute()
    return None



