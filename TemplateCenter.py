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
import TemplateCenterUtility as TCenterUtility
import TemplateCenterConf as TCenterConf
from urllib import request
import urllib
import json

def renderInfoLabel(dialog, jsonData):
    oGridModel = dialog.getControl("ListGrid").Model
    TotalLabel = dialog.getControl("Total")
    ODTLabel = dialog.getControl("ODT")
    ODSLabel = dialog.getControl("ODS")
    ODPLabel = dialog.getControl("ODP")
    NewLabel = dialog.getControl("New")

    totalCount = 0
    ODTCount = 0
    ODSCount = 0
    ODPCount = 0
    NewCount = 0

    for depart in jsonData:
        for template in jsonData[depart]:
            totalCount += 1
            extname = template['extname']
            if extname == 'ott':
                ODTCount += 1
            elif extname == "ots":
                ODSCount += 1
            elif extname == "otp":
                ODPCount += 1
    for color in oGridModel.RowBackgroundColors:
        if color == int("ffb5b5", 16):
            NewCount += 1
    TotalLabel.Model.Text = str(totalCount)
    ODTLabel.Model.Text = str(ODTCount)
    ODSLabel.Model.Text = str(ODSCount)
    ODPLabel.Model.Text = str(ODPCount)
    NewLabel.Model.Text = str(NewCount)

def checkDiff(dialog):
      
    grid = dialog.getControl("ListGrid")
    oGridModel = grid.Model
    oDataModel = oGridModel.GridDataModel

    try:
        with open(TCenterConf.getTemplateInfoPath(), "r") as jsonFile:
            oldTemplateInfo = json.load(jsonFile)
    except:
        oldTemplateInfo = {}

    try:
        url = TCenterConf.getAPIAddress_List()
        res = urllib.request.urlopen(url, timeout=2)
        newTemplateInfo = json.loads(res.read().decode())
    except:
        message = "請確認伺服器設定\n\n當前伺服器位置設定" + TCenterConf.getAPIAddress_List()
        TCenterUtility.Msgbox(message)
        oDataModel.addRow("X", ["確認網路服務", "伺服器是否有正確設定"])
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
            diffJson[nDepart] = []
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

    renderInfoLabel(dialog, newTemplateInfo)

    with open(TCenterConf.getDiffInfoPath(), "w") as outFile:
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
    xxGridModel.ShowRowHeader = True
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

from com.sun.star.task import XJob
class TemplateCenter(unohelper.Base, XJob):
        def __init__(self, ctx):
            self.ctx = ctx

        def execute(self, args):
            smgr = self.ctx.getServiceManager()
            dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)

            dialog = dp.createDialog(
                "vnd.sun.star.script:TemplateCenter.TemplateList?location=application")
            createGrid(dialog)
            checkDiff(dialog)
            dialog.execute()
            dialog.dispose()

# Registration
IMPLE_NAME = "TemplateCenter.Main"

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    TemplateCenter, IMPLE_NAME, (IMPLE_NAME,),)


