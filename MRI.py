#  Copyright 2011 Tsutomu Uchino
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.

import unohelper
import TemplateRepoConf as TRConf
from urllib import request
import urllib, json

TEST_IMPLE_NAME = "test.test"

def checkDiff(dialog):
    
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel

    try:
        with open(TRConf.InfoFilePath, "r") as jsonFile:
            oldTemplateInfo = json.load(jsonFile)
    except:
        oldTemplateInfo = {}
    
    try:
        url = "http://192.168.3.11:9980/lool/templaterepo/list"
        res = urllib.request.urlopen(url)
        newTemplateInfo = json.loads(res.read().decode())
    except:
        oDataModel.addRow("X", ["請確認網服務","或是","伺服器是否有正確設定"])
        return


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

    with open(r"C:\\Users\\Tommy\\AppData\\NDCFILE\\diffInfo.json", "w") as outFile:
        json.dump(diffJson, outFile)

def createGrid(dialog):
    grid = dialog.getModel().createInstance(
        "com.sun.star.awt.grid.UnoControlGridModel")
    dialog.getModel().insertByName("ListGrid", grid)
    grid = dialog.getControl("ListGrid")
    xxGridModel = grid.Model
    xxGridModel.PositionX = 7
    xxGridModel.PositionY = 30
    xxGridModel.Width = 335
    xxGridModel.Height = 270
    xxGridModel.Enabled = True
    xxGridModel.ShowRowHeader=True
    oGridWidth = xxGridModel.Width
    oColumnModel = xxGridModel.ColumnModel
    oCol = oColumnModel.createColumn()

    maxColumn = 4

    oCol.Title = "部門"
    oCol.MaxWidth = oGridWidth/maxColumn
    oColumnModel.addColumn(oCol)

    oCol = oColumnModel.createColumn()
    oCol.MaxWidth = oGridWidth/maxColumn
    oCol.Title = "檔名"
    oColumnModel.addColumn(oCol)

    oCol = oColumnModel.createColumn()
    oCol.MaxWidth = oGridWidth/maxColumn
    oCol.Title = "類型"
    oColumnModel.addColumn(oCol)

    oCol = oColumnModel.createColumn()
    oCol.MaxWidth = oGridWidth/maxColumn
    oCol.Title = "最後更新時間"
    oColumnModel.addColumn(oCol)

def test(ctx, *args):
    smgr = ctx.getServiceManager()
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)

    dialog = dp.createDialog(
        "vnd.sun.star.script:TemplatesMarket.TemplateList?location=application")
    createGrid(dialog)
    checkDiff(dialog)
    dialog.execute()
    dialog.dispose()


# Registration
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    test, TEST_IMPLE_NAME, (TEST_IMPLE_NAME,),)
