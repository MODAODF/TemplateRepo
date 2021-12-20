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

from com.sun.star.task import XJob
import unohelper
import TemplateRepoUtility as TRepoUtility
import TemplateRepoConf as TRepoConf
from urllib import request
import urllib
import json
import os,sys


def createGrid(dialog):
    grid = dialog.getModel().createInstance(
        "com.sun.star.awt.grid.UnoControlGridModel")
    imgControl = dialog.getControl("icon")
    imgModel = imgControl.Model
    imgModel.ImageURL = "file:///" + TRepoConf.getProjectImagesPath() + "templist_icon.png"
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

    oCol.Title = "類別"
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


class TemplateRepo(unohelper.Base, XJob):
    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, args):
        smgr = self.ctx.getServiceManager()
        dp = smgr.createInstanceWithContext(
            "com.sun.star.awt.DialogProvider", self.ctx)

        dialog = dp.createDialog(
            "vnd.sun.star.script:TemplateRepo.TemplateList?location=application")
        createGrid(dialog)
        if TRepoUtility.checkDiff(dialog):
            dialog.setVisible(True)
            TRepoUtility.syncTemplates(dialog)

        dialog.execute()

# Registration for StartBasic Usage
IMPLE_NAME = "TemplateRepo.Main"

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    TemplateRepo, IMPLE_NAME, (IMPLE_NAME,),)
