#!/bin/python
# -*- coding: utf-8 -*-
# mode:python; tab-width:4 -*- ex:set tabstop=4 shiftwidth=4 expandtab: -*-


# -------------------------------------------------------------------------
# Name:         monitor.py
# Purpose:      main frame for svn confirm
# Author:       Zhang Xiaoming
# Created:      2014-9-15
# Copyright:    Netease
# -------------------------------------------------------------------------

"""
Pydoc
…
"""

__version__ = "1.3"

import logging
import wx
import wx.animate

from baseframe import EditListCtrl, EditListCtrl2, wxLogText,EditListCtrl3

from ExcelReader import ExcelReader,TDataReader
from DialogFactory import AddSuiteDlg,AddSubSuiteDlg,SelectPanel
import xlrd
import xlwt
import time
import threading


titlestring = "GM Command Tool Kit"

configFilename = "config.ini"

image_list = ['client.png', 'test.png', 'client.png']


class MonitorFrame(wx.Frame):
    def __init__(self, parent):
        """
        initialize frame
        """

        wx.Frame.__init__(self, parent, -1, titlestring, size=(1200, 900), style=wx.DEFAULT_FRAME_STYLE)
        self.CenterOnScreen()

        self.initData()
        self.initUI()
        self.initLayout()
        self.initLists()
        self.initLogging()

        self.initListFirstTime()

    def initData(self):

        self.buttonDefaultSize = (88, 30)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.exceldict = None

        # get the data needed ready..  refresh and to use the data...
        pass

    def initUI(self):
        """
        initialize book_page, include quest_page, monitor_page, capture_page, date_base_page, test_suite_page, about_page
            quest_page, monitor_page includes grid
            capture_page, date_base_page, test_suite_page includes list_ctrl and tool_panel, list_ctrl display the information,
            tool_panel include the buttons
        """

        self.CreateStatusBar()
        self.SetStatusText("Welcome to X9 GM Command Tool Kit!")

        menuBar = wx.MenuBar()

        # 1st menu from left
        menu = wx.Menu()

        self.menuRefreshID = wx.NewId()
        self.menuCloseID = wx.NewId()

        menu.Append(self.menuRefreshID, "&Refresh", "Refresh data from the Excel.")
        menu.AppendSeparator()
        menu.Append(self.menuCloseID, "&Exit", "Exit.")

        # Append menu to the menu bar
        menuBar.Append(menu, "&Admin")
        self.SetMenuBar(menuBar)

        self.Bind(wx.EVT_MENU, self.OnRefreshListData, id=self.menuRefreshID)
        self.Bind(wx.EVT_MENU, self.OnClose, id=self.menuCloseID)

        # ################################  menu

        self.bookPanel = wx.Panel(self, -1)
        logobmp = wx.Image("./pic/gmlogo.png", wx.BITMAP_TYPE_PNG).ConvertToBitmap()

        self.logobmp = wx.StaticBitmap(self.bookPanel, -1, logobmp)
        self.bookPage = wx.Toolbook(self.bookPanel, -1, style=wx.BK_LEFT | wx.BORDER_DOUBLE)

        imageList = wx.ImageList(38, 40)
        # print "test1"

        for i in image_list:
            bmp = wx.Bitmap("./pic/%s" % (i), wx.BITMAP_TYPE_PNG)
            imageList.Add(bmp)

        self.bookPage.AssignImageList(imageList)
        self.bookPage.Bind(wx.EVT_TOOLBOOK_PAGE_CHANGED, self.OnPageChanged)

        self.clientPanel = wx.Panel(self.bookPage, -1)

        self.clientListPanel = wx.Panel(self.clientPanel, -1)

        self.clientCommandList = EditListCtrl(self.clientListPanel)
        self.clientCommandList.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnClientCommandListItemSelected)

        self.clientCommandRCPanel = wx.Panel(self.clientPanel, -1)  # ##################

        self.clientCommandEntityPanel = wx.Panel(self.clientCommandRCPanel, -1)

        self.eneitytextinfo = "This command need entity. Run /hl 1 ?"
        self.statictext_eneityinfo = wx.StaticText(self.clientCommandEntityPanel, -1, self.eneitytextinfo)
        # self.statictext_eneityinfo.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD))
        self.eneitysendButton = wx.Button(self.clientCommandEntityPanel, -1, u"Entity", size=(80, 50))
        self.Bind(wx.EVT_BUTTON, self.OnEneitySendButton, self.eneitysendButton)

        self.clientCommandRunPanel = wx.Panel(self.clientCommandRCPanel, -1)  # args,run

        self.commandtorun = "The command to run is:    "
        self.statictext_commandtorun = wx.StaticText(self.clientCommandRunPanel, -1, self.commandtorun)
        self.EditText_commandtorun = wx.TextCtrl(self.clientCommandRunPanel, -1, "Command Show Here", size=(3000, -1))
        # self.EditText_commandtorun.SetEditable(False)
        self.editlistcrtl2 = EditListCtrl2(self.clientCommandRunPanel)

        self.editlistcrtl2.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.OnEndEdit)
        self.editlistcrtl2.Bind(wx.EVT_LIST_BEGIN_LABEL_EDIT, self.OnBeginEdit)
        # self.editlistcrtl2.Enable(False)


        self.clientCommandContextPanel = wx.Panel(self.clientCommandRCPanel, -1)  # context
        #self.statictext_contentPara = wx.StaticText(self.clientCommandContextPanel,-1,"Context to display parameter")
        self.runButton = wx.Button(self.clientCommandContextPanel, -1, u"Run Client Command", size=(88, 80))
        self.Bind(wx.EVT_BUTTON, self.OnRunButton, self.runButton)

        self.serverPanel = wx.Panel(self.bookPage, -1)
        self.serverListPanel = wx.Panel(self.serverPanel, -1)
        self.serverCommandList = EditListCtrl(self.serverListPanel)
        self.serverCommandList.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnServerCommandListItemSelected)

        self.serverCommandRCPanel = wx.Panel(self.serverPanel, -1)  # ##################

        self.serverCommandRunPanel = wx.Panel(self.serverCommandRCPanel, -1)  # args,run

        #self.commandtorun = "The command to run is:    "
        self.statictext_commandtorun_server = wx.StaticText(self.serverCommandRunPanel, -1, self.commandtorun)
        self.EditText_commandtorun_server = wx.TextCtrl(self.serverCommandRunPanel, -1, "Command Show Here",
                                                        size=(3000, -1))
        #self.EditText_commandtorun.SetEditable(False)
        self.editlistcrtl_server2 = EditListCtrl2(self.serverCommandRunPanel)

        self.serverCommandContextPanel = wx.Panel(self.serverCommandRCPanel, -1)  # context
        self.runButton2 = wx.Button(self.serverCommandContextPanel, -1, u"Run Server Command", size=(88, 80))
        self.Bind(wx.EVT_BUTTON, self.OnRunButton2, self.runButton2)

        self.testsetPanel = wx.Panel(self.bookPage, -1)

        self.testsetListPanel = wx.Panel(self.testsetPanel, -1)
        self.testsetList = EditListCtrl(self.testsetListPanel)
        self.testsetList.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnTestSetListItemSelected)  # add the list for all..
        self.testsetList.Bind(wx.EVT_LIST_ITEM_DESELECTED,self.OnTestSetListItemDeSelected)

        self.testsetToolPanel = wx.Panel(self.testsetListPanel)
        self.addsetToolButton = wx.Button(self.testsetToolPanel, -1, u"+", size=(30, 30))
        self.delsetToolButton = wx.Button(self.testsetToolPanel, -1, u"-", size=(30, 30))
        self.Bind(wx.EVT_BUTTON, self.OnaddsetButton, self.addsetToolButton)
        self.Bind(wx.EVT_BUTTON, self.OndelsetButton, self.delsetToolButton)

        self.testsetsubListPanel = wx.Panel(self.testsetPanel, -1)
        self.testsetsubList = EditListCtrl(self.testsetsubListPanel)
        self.testsetsubList.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnTestSetSubListItemSelected)  # add the list for all..
        self.testsetsubList.Bind(wx.EVT_LIST_ITEM_DESELECTED, self.OnTestSetSubListItemDeSelected)  # add the list for all..

        self.testsetsubToolPanel = wx.Panel(self.testsetsubListPanel)
        self.addsubsetToolButton = wx.Button(self.testsetsubToolPanel, -1, u"+", size=(30, 30))
        self.delsubsetToolButton = wx.Button(self.testsetsubToolPanel, -1, u"-", size=(30, 30))
        self.Bind(wx.EVT_BUTTON, self.OnaddsubsetButton, self.addsubsetToolButton)
        self.Bind(wx.EVT_BUTTON, self.OndelsubsetButton, self.delsubsetToolButton)

        self.runsetButton = wx.Button(self.testsetsubListPanel,-1, u"Run Commands", size=(130, -1))
        self.Bind(wx.EVT_BUTTON, self.Onrunsetbutton, self.runsetButton)

        self.bookPage.AddPage(self.testsetPanel, u"T_Command", imageId=0)
        self.bookPage.AddPage(self.serverPanel, u"S_Command", imageId=1)
        self.bookPage.AddPage(self.clientPanel, u"C_Command", imageId=2)

        self.msgPanel = wx.Panel(self, -1)

        self.staticbox = wx.StaticBox(self.msgPanel, -1, "Command Used Log")
        self.msgTextCtrl = wx.TextCtrl(self.msgPanel, size=(-1, 140), style=wx.TE_PROCESS_ENTER | wx.TE_MULTILINE)

        # self.bsizer = wx.StaticBoxSizer(self.staticbox, wx.VERTICAL)
        # self.msgTextCtrl = wx.TextCtrl(self.msgPanel, size=(-1, 140), style=wx.TE_PROCESS_ENTER | wx.TE_MULTILINE)
        # self.bsizer.Add(self.msgTextCtrl, 0, wx.TOP | wx.LEFT, 10)


        # self.logListPanel = wx.Panel(self, -1)
        # self.loglist = EditListCtrl(self.logListPanel)
        # self.loglist.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnLogItemSelected)
        #
        # self.filediffPanel = wx.Panel(self, -1)
        # self.filediffPanel.Bind(wx.EVT_CONTEXT_MENU, self.OnContextMenu)
        #
        # self.filedifflist = EditListCtrl(self.filediffPanel)
        # self.filedifflist.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnfilediffItemSelected)
        #
        # self.toolPanel = wx.Panel(self.filediffPanel, -1)
        # self.refresh_button = wx.Button(self.toolPanel, -1, u"Refresh", size=(88, 50))
        # self.refresh_button.Enable(False)
        #
        # self.Bind(wx.EVT_BUTTON, self.OnRefreshButton, self.refresh_button)
        #
        # self.allconfirm_button = wx.Button(self.toolPanel, -1, u"AllConfirm", size=(88, 50))
        # self.Bind(wx.EVT_BUTTON, self.OnAllConfirmButton, self.allconfirm_button)
        # self.allconfirm_button.Enable(False)
        #
        # self.confirm_button = wx.Button(self.toolPanel, -1, u"Confirm", size=(88, 50))
        # self.Bind(wx.EVT_BUTTON, self.OnConfirmButton, self.confirm_button)
        # self.confirm_button.Enable(False)
        #
        # self.vacantPanel = wx.Panel(self.toolPanel, -1)
        #
        # self.ok_button = wx.Button(self.toolPanel, -1, u"Exit")
        # self.Bind(wx.EVT_BUTTON, self.OnOkButton, self.ok_button)

        pass


    def initLayout(self):
        """
        set layout of the picture_ctrl, page panel
        """
        self._icon = _icon = wx.EmptyIcon()
        _icon.LoadFile("./pic/pc.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(_icon)

        client_left_V_sizer = wx.BoxSizer(wx.VERTICAL)
        client_left_V_sizer.Add(self.clientCommandList, 1,
                                wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 5)
        self.clientListPanel.SetSizer(client_left_V_sizer)

        eneityPanelSizer = wx.BoxSizer(wx.HORIZONTAL)
        eneityPanelSizer.Add((30, -1))
        eneityPanelSizer.Add(self.statictext_eneityinfo, 4, wx.ALIGN_CENTER, 50)
        eneityPanelSizer.Add(self.eneitysendButton, 1, wx.ALIGN_CENTER, 50)
        eneityPanelSizer.Add((30, -1))
        self.clientCommandEntityPanel.SetSizer(eneityPanelSizer)
        # self.clientCommandEntityPanel.Enable(False)

        runpanelsizer_V = wx.BoxSizer(wx.VERTICAL)

        runpanelSizer = wx.BoxSizer(wx.HORIZONTAL)
        runpanelSizer.Add((30, -1))
        runpanelSizer.Add(self.statictext_commandtorun, 2, wx.ALIGN_CENTER, 5)
        runpanelSizer.Add(self.EditText_commandtorun, 1, wx.ALIGN_CENTER, 50)
        runpanelSizer.Add((30, -1))

        # self.runpanelSizer.Add(self.editlistcrtl2,1,wx.EXPAND, 50)

        runpanelsizer_V.Add(runpanelSizer, 1, wx.EXPAND | wx.ALIGN_CENTER, 20)
        runpanelsizer_V.Add(self.editlistcrtl2, 6, wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 25)

        self.clientCommandRunPanel.SetSizer(runpanelsizer_V)

        contentPanelSizer = wx.BoxSizer(wx.HORIZONTAL)
        contentPanelSizer.Add((20, -1))
        # contentPanelSizer.Add(self.statictext_contentPara,4,wx.ALIGN_CENTER,50)
        contentPanelSizer.Add(self.runButton, 1, wx.ALIGN_CENTER, 0)
        contentPanelSizer.Add((20, -1))
        self.clientCommandContextPanel.SetSizer(contentPanelSizer)

        client_right_V_sizer = wx.BoxSizer(wx.VERTICAL)

        client_right_V_sizer.Add(self.clientCommandEntityPanel, 1,
                                 wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)

        client_right_V_sizer.Add(self.clientCommandRunPanel, 7,
                                 wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)

        client_right_V_sizer.Add(self.clientCommandContextPanel, 2,
                                 wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        self.clientCommandRCPanel.SetSizer(client_right_V_sizer)

        client_all_H_sizer = wx.BoxSizer(wx.HORIZONTAL)
        # client_all_H_sizer.Add(self.clientCommandList)
        client_all_H_sizer.Add(self.clientListPanel, 1,
                               wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        client_all_H_sizer.Add(self.clientCommandRCPanel, 1,
                               wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        self.clientPanel.SetSizer(client_all_H_sizer)
        # ##########   the upper is for the client panel


        server_left_V_sizer = wx.BoxSizer(wx.VERTICAL)
        server_left_V_sizer.Add(self.serverCommandList, 1,
                                wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 5)
        self.serverListPanel.SetSizer(server_left_V_sizer)

        runpanelsizer_V2 = wx.BoxSizer(wx.VERTICAL)

        runpanelSizer2 = wx.BoxSizer(wx.HORIZONTAL)
        runpanelSizer2.Add((30, -1))
        runpanelSizer2.Add(self.statictext_commandtorun_server, 2, wx.ALIGN_CENTER, 5)
        runpanelSizer2.Add(self.EditText_commandtorun_server, 1, wx.ALIGN_CENTER, 50)
        runpanelSizer2.Add((30, -1))

        runpanelsizer_V2.Add(runpanelSizer2, 1, wx.EXPAND | wx.ALIGN_CENTER, 20)
        runpanelsizer_V2.Add(self.editlistcrtl_server2, 8, wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM,
                             25)
        self.serverCommandRunPanel.SetSizer(runpanelsizer_V2)

        server_contentPanelSizer2 = wx.BoxSizer(wx.HORIZONTAL)
        server_contentPanelSizer2.Add((20, -1))
        # contentPanelSizer.Add(self.statictext_contentPara,4,wx.ALIGN_CENTER,50)
        server_contentPanelSizer2.Add(self.runButton2, 1, wx.ALIGN_CENTER, 0)
        server_contentPanelSizer2.Add((20, -1))
        self.serverCommandContextPanel.SetSizer(server_contentPanelSizer2)

        server_right_V_sizer = wx.BoxSizer(wx.VERTICAL)

        server_right_V_sizer.Add(self.serverCommandRunPanel, 4,
                                 wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        server_right_V_sizer.Add(self.serverCommandContextPanel, 1,
                                 wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        self.serverCommandRCPanel.SetSizer(server_right_V_sizer)

        server_all_H_sizer = wx.BoxSizer(wx.HORIZONTAL)
        # server_all_H_sizer.Add(self.clientCommandList)
        server_all_H_sizer.Add(self.serverListPanel, 1,
                               wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        server_all_H_sizer.Add(self.serverCommandRCPanel, 1,
                               wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        self.serverPanel.SetSizer(server_all_H_sizer)

        testsettoolSizer = wx.BoxSizer(wx.HORIZONTAL)
        testsettoolSizer.Add((400, -1))
        testsettoolSizer.Add(self.addsetToolButton, 0, wx.RIGHT | wx.ALIGN_CENTER, 5)
        testsettoolSizer.Add(self.delsetToolButton, 0, wx.RIGHT | wx.ALIGN_CENTER, 5)
        self.testsetToolPanel.SetSizer(testsettoolSizer)

        testsubsettoolSizer = wx.BoxSizer(wx.HORIZONTAL)
        testsubsettoolSizer.Add((400, -1))
        testsubsettoolSizer.Add(self.addsubsetToolButton, 0, wx.RIGHT | wx.ALIGN_CENTER, 5)
        testsubsettoolSizer.Add(self.delsubsetToolButton, 0, wx.RIGHT | wx.ALIGN_CENTER, 5)
        self.testsetsubToolPanel.SetSizer(testsubsettoolSizer)

        testsetListSizer = wx.BoxSizer(wx.VERTICAL)
        testsetListSizer.Add(self.testsetToolPanel, 1,
                             wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 5)
        self.testsetListPanel.SetSizer(testsetListSizer)
        testsetListSizer.Add(self.testsetList, 10,
                             wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)

        testsubsetListSizer = wx.BoxSizer(wx.VERTICAL)
        testsubsetListSizer.Add(self.testsetsubToolPanel, 1,
                                wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 4)
        testsubsetListSizer.Add(self.testsetsubList, 7,
                                wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)

        testsubsetListSizer.Add(self.runsetButton,2,
                                wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 5)

        self.testsetsubListPanel.SetSizer(testsubsetListSizer)

        testAllSizer = wx.BoxSizer(wx.HORIZONTAL)
        testAllSizer.Add(self.testsetListPanel, 1,
                         wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 8)
        testAllSizer.Add(self.testsetsubListPanel, 1,
                         wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 8)
        self.testsetPanel.SetSizer(testAllSizer)

        pageSizer = wx.BoxSizer(wx.VERTICAL)
        pageSizer.Add(self.logobmp, 1, wx.ALL | wx.EXPAND, 0)
        pageSizer.Add(self.bookPage, 8, wx.ALL | wx.EXPAND, 5)
        self.bookPanel.SetSizer(pageSizer)

        bsizer = wx.StaticBoxSizer(self.staticbox, wx.VERTICAL)
        bsizer.Add(self.msgTextCtrl, 1,
                   wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        self.msgPanel.SetSizer(bsizer)

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.bookPanel, 9, wx.ALL | wx.EXPAND, 0)
        mainSizer.Add(self.msgPanel, 1, wx.ALL | wx.EXPAND, 0)

        self.SetSizer(mainSizer)

        pass

    def Onrunsetbutton(self,event):
        #print "On run all"

        self.runsetButton.Enable(False)

        #读取sub表格中的command指令集，逐条，逐步运行。

        lenoflist = self.testsetsubList.ItemCount

        for i in range(lenoflist):

            commandtorun = self.testsetsubList.GetItemText(i,1)

            print commandtorun

            #self.doAllCommand()
            self.doAllCommand(commandtorun)
            logging.info("GM Command has been Send!" + "    " + str(commandtorun))

        self.runsetButton.Enable(True)
        pass

    def doServerCommand(self,command):

        # This function is used to do server commands
        try:
            from utils import monkey_patch
            import x9
            from x9.system import GmCmd
            from x9 import BigWorld as BW

            self.BigWorld = BW
            print "here11111"
            self.BigWorld.player().base.doGmCommand(command.encode('gbk'))

            print "here11111"

        except Exception, err:

            if str(err) == "No module named x9":
                # print "Open X9 and restart this tool!!"
                self.showDlg("Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return

            elif str(err) == "stream has been closed":
                # print "X9 has been CLOSED!!"
                self.showDlg("X9 Has Been CLOSED!!Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return

    def doClientCommand(self,command):

        # This funcion is used to do client commands
        try:
            from utils import monkey_patch
            import x9
            from x9.system import GmCmd
            from x9 import BigWorld as BW

            self.BigWorld = BW

            GmCmd.prase_cmd(command.encode('gbk'))

        except Exception, err:

            if str(err) == "No module named x9":
                # print "Open X9 and restart this tool!!"
                self.showDlg("Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return
            elif str(err) == "stream has been closed":
                # print "X9 has been CLOSED!!"
                self.showDlg("X9 Has Been CLOSED!!Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return

    def doAllCommand(self,command):

        if command.startswith("$"):
            print "do server command"
            self.doServerCommand(command)
        elif command.startswith("/"):
            self.doClientCommand(command)


    def OnaddsubsetButton(self, event):

        print "OnaddsubsetButton"

        #添加commands，增加一条commdnas setid 和content
        print self.leftcommandid,"right"
        if self.leftcommandid is not None:
            dlg = AddSubSuiteDlg(self,self.leftcommandid)
            result = dlg.ShowModal()
            if result == wx.ID_OK:
                content = dlg.getData()

            #进行至此处，将setid和content加入到xls中

                T_reader = TDataReader()
                T_reader.writeTData_command(self.leftcommandid,content)
                del T_reader
                logging.info("Command has been added!" + "    " + str(content))
            #print setid,content
        else:
            pass

        self.OnSubSetListRefresh()

        pass

    def OndelsubsetButton(self, event):

        #print "OndelsubsetButton"
        #删除某一条command。
        #print self.rightcommandid,"right"

        if self.rightcommandid is not None:
            dlg = wx.MessageDialog(self, "Are you sure to delete  " + str(self.rightcommandid)+"  command ?",
                               "Delete...",
                               wx.YES_NO | wx.YES_DEFAULT | wx.ICON_INFORMATION
            )
            if dlg.ShowModal() == wx.ID_YES:

                # 通过选中的id，查找总的id号码。使用接口删除。
                #print self.leftcommandid,self.rightcommandid,"##########"
                T_reader = TDataReader()
                T_reader.deleteTData_command(self.leftcommandid,self.rightcommandid)
                del T_reader

                #print setid,content
            else:
                pass
        else:
            print "select to delete"

        self.OnSubSetListRefresh()

        #进行至此处。

        pass


    def OnaddsetButton(self, event):

        print "OnaddsetButton"

        # 弹出用以增加group的对话框。让用户增加新的group。同时调动后台接口，将数据写入Excel中。前台界面

        # 点击确定之后，将Excel数据验证之后写回。将屏幕面板刷新。
        #self.testsetList
        #self.testsetList.GetItemCount()

        setid = self.testsetList.GetItemCount()+1

        dlg = AddSuiteDlg(self,setid)
        result = dlg.ShowModal()
        if result == wx.ID_OK:
            content = dlg.getData()

            #进行至此处，将setid和content加入到xls中

            T_reader = TDataReader()
            T_reader.writeTData_group(setid,content)
            del T_reader
            logging.info(u"Commands Set has been added!" + "    " + unicode(content))
            #print setid,content
        else:
            pass

        self.OnSetListRefresh()


        pass

    def OndelsetButton(self, event):

        #print "OndelsetButton"

        # 弹出二次确认对话框。 点击确定之后，将内存数据和后台Excel数据同时删除。之后将屏幕刷新。
        #print "delete  "+ str(self.leftcommandid)

        dlg = wx.MessageDialog(self, "Are you sure to delete  " + str(self.leftcommandid)+"  set ?",
                               "Delete...",
                               wx.YES_NO | wx.YES_DEFAULT | wx.ICON_INFORMATION
        )
        if dlg.ShowModal() == wx.ID_YES:
            if self.leftcommandid is not None:
                T_reader = TDataReader()
                T_reader.deleteTData_group(self.leftcommandid)
                del T_reader
            else:
                pass
            #print setid,content
        else:
            pass

        self.OnSetListRefresh()
        self.OnSubSetListRefresh()

        pass

    def OnSetListRefresh(self):
        # 读取Excel 并展示

        T_reader = TDataReader()
        self.T_Data = T_reader.getTData()

        #print T_Data
        #{u'commands': {1: {u'Content': u'$mhp', u'Id': u'1', u'Set_id': u'1'}, 2: {u'Content': u'/cd', u'Id': u'2', u'Set_id': u'1'}, 3: {u'Content': u'$mhp', u'Id': u'3', u'Set_id': u'2'}, 4: {u'Content': u'/cd', u'Id': u'4', u'Set_id': u'2'}, 5: {u'Content': u'$mhp', u'Id': u'5', u'Set_id': u'2'}}, u'group': {1: {u'Content': u'\u7528\u6765\u6d4b\u8bd5\u7ec3\u4e60\u573a\u7684\u6307\u4ee4\u96c6', u'Set_id': u'1'}, 2: {u'Content': u'\u7528\u6765\u6d4b\u8bd55v5\u7684\u6307\u4ee4\u96c6', u'Set_id': u'2'}}}

        T = self.T_Data.get(u"group")
        # print len(clientData), "tset"

        self.testsetList.DeleteAllItems()

        for i in range(len(T)):

            index = self.testsetList.InsertStringItem(i, str(i))
            self.testsetList.SetStringItem(index, 0, unicode(i + 1))
            self.testsetList.SetStringItem(index, 1, T.get(i + 1).get(u'Content'))

    def OnSubSetListRefresh(self):
        T_reader = TDataReader()
        self.T_Data = T_reader.getTData()

        #print T_Data
        #{u'commands': {1: {u'Content': u'$mhp', u'Id': u'1', u'Set_id': u'1'}, 2: {u'Content': u'/cd', u'Id': u'2', u'Set_id': u'1'}, 3: {u'Content': u'$mhp', u'Id': u'3', u'Set_id': u'2'}, 4: {u'Content': u'/cd', u'Id': u'4', u'Set_id': u'2'}, 5: {u'Content': u'$mhp', u'Id': u'5', u'Set_id': u'2'}}, u'group': {1: {u'Content': u'\u7528\u6765\u6d4b\u8bd5\u7ec3\u4e60\u573a\u7684\u6307\u4ee4\u96c6', u'Set_id': u'1'}, 2: {u'Content': u'\u7528\u6765\u6d4b\u8bd55v5\u7684\u6307\u4ee4\u96c6', u'Set_id': u'2'}}}
        # print len(clientData), "tset"

        commandsheet = self.T_Data.get(u"commands")
        #print commandsheet
        commandstodisplay = []

        for i,key in enumerate(commandsheet):
            #print i ,key ,commandsheet.get(key)
            tempcommand = commandsheet.get(key)
            if tempcommand.get(u'Set_id') == unicode(self.leftcommandid):
                commandstodisplay.append(tempcommand)

        self.testsetsubList.DeleteAllItems()
        for counter in range(len(commandstodisplay)):
            index = self.testsetsubList.InsertStringItem(counter, str(counter))

            self.testsetsubList.SetStringItem(index, 0, unicode(counter + 1))
            self.testsetsubList.SetStringItem(index, 1,commandstodisplay[counter].get(u'Content'))


    def OnTestSetListItemSelected(self, event):

        self.addsubsetToolButton.Enable(True)
        self.delsubsetToolButton.Enable(True)
        self.addsetToolButton.Enable(True)
        self.delsetToolButton.Enable(True)
        commandsetid =  event.m_itemIndex + 1
        self.rightcommandid = None
        self.leftcommandid = commandsetid

        commandsheet = self.T_Data.get(u"commands")

        commandstodisplay = []

        for i,key in enumerate(commandsheet):
            #print i ,key ,commandsheet.get(key)
            tempcommand = commandsheet.get(key)
            if tempcommand.get(u'Set_id') == unicode(self.leftcommandid):
                commandstodisplay.append(tempcommand)

        self.testsetsubList.DeleteAllItems()
        for counter in range(len(commandstodisplay)):
            index = self.testsetsubList.InsertStringItem(counter, str(counter))

            self.testsetsubList.SetStringItem(index, 0, unicode(counter + 1))
            self.testsetsubList.SetStringItem(index, 1,commandstodisplay[counter].get(u'Content'))

        pass

    def OnTestSetListItemDeSelected(self,event):
        #print "OnTestSetListItemDeSelected"

        self.addsetToolButton.Enable(False)
        self.delsetToolButton.Enable(False)
        self.addsubsetToolButton.Enable(False)
        self.delsubsetToolButton.Enable(False)

        self.leftcommandid = None
        pass

    def OnTestSetSubListItemSelected(self, event):

        #print "OnTestSetsubListItemSelected"
        #右侧选中的list。
        self.addsetToolButton.Enable(False)
        self.delsetToolButton.Enable(False)
        self.addsubsetToolButton.Enable(True)
        self.delsubsetToolButton.Enable(True)
        commandid =  event.m_itemIndex + 1
        self.rightcommandid = commandid
        #print self.rightcommandid

        pass
    def OnTestSetSubListItemDeSelected(self, event):
        #print "OnTestSetsubListItemDeSelected"
        self.addsubsetToolButton.Enable(False)
        self.delsubsetToolButton.Enable(False)
        self.rightcommandid = None
        pass

    def OnRunButton2(self, event):
        #print "OnRunbutton2"

        command = self.EditText_commandtorun_server.GetValue()

        if command is not None and command != "Command Show Here":

            self.doServerCommand(command)
            logging.info("Server GM Command has been Send!" + "       " + command)
        else:
            print "nothing"

        pass


    def OnRunButton(self, event):
        # print "runbutton"

        command = str(self.EditText_commandtorun.GetValue())
        if command is not None and command != "Command Show Here":

            self.doClientCommand(command)

            logging.info("Client GM Command has been Send!" + "     " + str(command))
        else:
            print "nothing"
        pass

    def OnEndEdit(self, event):
        print "end edit"

        pass

    def OnBeginEdit(self, event):
        print "begin edit"
        pass


    def OnPaneChanged(self, event):

        print "OnPaneChanged"

        # redo the layout
        self.clientCommandRunPanel.Layout()
        # self.Layout()

        pass

    def OnEneitySendButton(self, event):

        print "OnEneitySendButton"
        # self.eneitytextinfo = "changed"
        # self.statictext_eneityinfo.SetLabel(self.eneitytextinfo)
        # self.commandtorun = "The command to run is:"
        #
        # self.buttontest2 = wx.Button(self.clientCommandRunPanel, -1, u"Confirm", size=(88, 50))
        #
        # self.runpanelSizer.Add(self.buttontest2,2, wx.ALL | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 0)
        # self.runpanelSizer.Layout()

        command = "/hl 1"

        # self.comm

        self.doClientCommand(command)

        pass

    def OnRefreshListData(self, event):
        print "OnRefreshList"

        er = ExcelReader()
        self.exceldict = er.excelData()  # get the data from the excel
        print self.exceldict
        self.OnRefreshList(event)

        pass

    def OnRefreshList(self, event):

        # use the self.exceldict to fresh list

        # {u'client': {1: {u'para_6': 0.0, u'command': u'/cd', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd51', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 1.0, u'para_1': 0.0, u'command_para_num': 0.0, u'para_3': 0.0, u'para_2': 0.0}, 2: {u'para_6': 0.0, u'command': u'/cst', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd52', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 2.0, u'para_1': 0.0, u'command_para_num': 1.0, u'para_3': 0.0, u'para_2': 0.0}, 3: {u'para_6': 0.0, u'command': u'/fps', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd53', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 3.0, u'para_1': 0.0, u'command_para_num': 0.0, u'para_3': 0.0, u'para_2': 0.0}, 4: {u'para_6': 0.0, u'command': u'/hl', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd54', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 4.0, u'para_1': 0.0, u'command_para_num': 1.0, u'para_3': 0.0, u'para_2': 0.0}}, u'server': {1: {u'para_6': 0.0, u'command': u'/cd', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd51', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 1.0, u'para_1': 0.0, u'command_para_num': 0.0, u'para_3': 0.0, u'para_2': 0.0}, 2: {u'para_6': 0.0, u'command': u'/cst', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd52', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 2.0, u'para_1': 0.0, u'command_para_num': 1.0, u'para_3': 0.0, u'para_2': 0.0}, 3: {u'para_6': 0.0, u'command': u'/fps', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd53', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 3.0, u'para_1': 0.0, u'command_para_num': 0.0, u'para_3': 0.0, u'para_2': 0.0}, 4: {u'para_6': 0.0, u'command': u'/hl', u'command_content': u'\u5185\u5bb9\u6d4b\u8bd54', u'para_5': 0.0, u'para_4': 0.0, u'para_7': 0.0, u'command_id': 4.0, u'para_1': 0.0, u'command_para_num': 1.0, u'para_3': 0.0, u'para_2': 0.0}}}

        # print "used "
        clientData = self.exceldict.get(u"client")
        # print len(clientData), "tset"

        self.clientCommandList.DeleteAllItems()

        for i in range(len(clientData)):
            #print clientData.get(i + 1)
            index = self.clientCommandList.InsertStringItem(i, str(i))

            self.clientCommandList.SetStringItem(index, 0, unicode(i + 1))
            self.clientCommandList.SetStringItem(index, 1, clientData.get(i + 1).get(u'command'))
            self.clientCommandList.SetStringItem(index, 2, clientData.get(i + 1).get(u'command_content'))

        serverData = self.exceldict.get(u"server")
        self.serverCommandList.DeleteAllItems()
        for i in range(len(serverData)):
            #print serverData.get(i + 1)
            index = self.serverCommandList.InsertStringItem(i, str(i))

            self.serverCommandList.SetStringItem(index, 0, unicode(i + 1))
            self.serverCommandList.SetStringItem(index, 1, serverData.get(i + 1).get(u'command'))
            self.serverCommandList.SetStringItem(index, 2, serverData.get(i + 1).get(u'command_content'))

        # self.loglist.DeleteAllItems()
        #
        # for i, iversion in enumerate(self.allversion):
        # oneline = self.alldata.get(unicode(iversion))
        # # print self.alldata,"^^^^^^^^^^^^^^^^",iversion,oneline,"*****************",oneline.get(u"date")
        # # print
        #
        #
        # if oneline is not None:
        # oneline_date = oneline.get(u"date")
        # oneline_message = oneline.get(u"message")
        # oneline_author = oneline.get(u"author")
        # oneline_action = oneline.get(u"action")
        # oneline_version = oneline.get(u"revision")
        #
        # # oneline_changepaths = oneline.get("chaned paths")
        #
        # action_tmp = " ".join(oneline_action)
        #
        # if oneline_message is None:
        # oneline_message = u""
        #
        # index = self.loglist.InsertStringItem(i, unicode(iversion))
        # self.loglist.SetStringItem(index, 0, unicode(oneline_version))
        # self.loglist.SetStringItem(index, 1, unicode(action_tmp))
        # # self.loglist.SetStringItem(index,2,str(oneline_action))
        # self.loglist.SetStringItem(index, 2, unicode(oneline_author))
        # self.loglist.SetStringItem(index, 3, unicode(oneline_date))
        # self.loglist.SetStringItem(index, 4, unicode(oneline_message))
        #
        # dicttosend = {}
        # dicttosend["type"] = 1
        #     dicttosend["ask"] = 2
        #     dicttosend["version"] = iversion
        #
        #     # print dicttosend, "@@@@@@@@@@@"
        #
        #     self.socketthread.setData(dicttosend)
        #
        #     pass


        pass


    def OnServerCommandListItemSelected(self, event):

        # print "OnServerCommandListItemSelected"
        command = self.exceldict.get("server").get(int(event.m_itemIndex + 1)).get("command")
        # command_need_eneity = self.exceldict.get("server").get(int(event.m_itemIndex + 1)).get("command_need_eneity")
        command_para_number = self.exceldict.get("server").get(int(event.m_itemIndex + 1)).get("command_para_num")

        self.editlistcrtl_server2.DeleteAllItems()
        if 1 <= int(command_para_number) <= 7:

            self.editlistcrtl_server2.Enable(True)
            #parameters = {}
            parameters2 = []

            for i in range(int(command_para_number)):
                keytouse = "para_" + str(i + 1)

                para_name = self.exceldict.get("server").get(int(event.m_itemIndex + 1)).get(keytouse)
                para_default = self.exceldict.get("serverdefault").get(int(event.m_itemIndex + 1)).get(keytouse)

                if isinstance(para_default, float):
                    para_default = int(para_default)
                elif isinstance(para_default, int):
                    para_default = int(para_default)
                else:
                    pass
                #parameters[para_name] = str(para_default)
                parameters2.append(unicode(para_default))
                #print para_name,para_default

                index = self.editlistcrtl_server2.InsertStringItem(i, str(i))
                self.editlistcrtl_server2.SetStringItem(index, 0, unicode(para_name))
                self.editlistcrtl_server2.SetStringItem(index, 1, unicode(para_default))
            #print parameters ,"confect"
            self.command_to_run = self.makecommand2(command, parameters2)

            self.EditText_commandtorun_server.SetLabel(self.command_to_run)

            # for i in range(int(command_para_number)):


        elif int(command_para_number) == 0:
            self.editlistcrtl_server2.Enable(False)
            self.EditText_commandtorun_server.SetLabel(command)

        pass

    def OnClientCommandListItemSelected(self, event):

        # print "OnClientCommandListItemSelected"
        # print event.m_itemIndex
        #
        # print self.exceldict

        command = self.exceldict.get("client").get(int(event.m_itemIndex + 1)).get("command")
        command_need_eneity = self.exceldict.get("client").get(int(event.m_itemIndex + 1)).get("command_need_eneity")
        command_para_number = self.exceldict.get("client").get(int(event.m_itemIndex + 1)).get("command_para_num")

        # print int(command_need_eneity), "test2"
        # print int(command_para_number), "test3"
        # print command

        # command_para_command = self.makecommand(command_para_command)

        # to make the eneitypanel ok
        if command_need_eneity == 1:
            # add the UI to send /hl 1 or /hl 0 or add this and to set the panel visiable of disable.
            self.clientCommandEntityPanel.Enable(True)

        else:
            # just use the normal UI
            self.clientCommandEntityPanel.Enable(False)

        # to flush the UI command need to run default
        self.editlistcrtl2.DeleteAllItems()
        if isinstance(command_para_number, float) and command_para_number is not None and 1 <= int(
                command_para_number) <= 7:

            self.editlistcrtl2.Enable(True)
            # parameters = {}
            parameters2 = []

            for i in range(int(command_para_number)):
                keytouse = "para_" + str(i + 1)

                para_name = self.exceldict.get("client").get(int(event.m_itemIndex + 1)).get(keytouse)
                para_default = self.exceldict.get("clientdefault").get(int(event.m_itemIndex + 1)).get(keytouse)

                if isinstance(para_default, float):
                    para_default = int(para_default)
                elif isinstance(para_default, int):
                    para_default = int(para_default)
                else:
                    pass
                # parameters[para_name] = str(para_default)
                parameters2.append(unicode(para_default))
                #print para_name,para_default

                index = self.editlistcrtl2.InsertStringItem(i, str(i))
                self.editlistcrtl2.SetStringItem(index, 0, unicode(para_name))
                self.editlistcrtl2.SetStringItem(index, 1, unicode(para_default))
            # print parameters ,"confect"
            self.command_to_run = self.makecommand2(command, parameters2)

            self.EditText_commandtorun.SetLabel(self.command_to_run)

            # for i in range(int(command_para_number)):

        elif isinstance(command_para_number, float) and command_para_number is not None and int(
                command_para_number) == 0:
            self.editlistcrtl2.Enable(False)
            self.EditText_commandtorun.SetLabel(command)

        pass

    def makecommand2(self, command, list):
        s = command
        if s is not None:
            if list is not None:
                for i in list:
                    s = s + " " + i
            else:
                s = command
        return s

    def makecommand(self, command, dict):
        # this function is to make the commands.
        s = command
        if s is not None:
            if dict is not None:
                for val in dict.values():
                    s = s + " " + val
            else:
                s = command
        return s

    def OnPageChanged(self, event):

        # print "OnPangChanged"

        pass


    def initLists(self):
        """
        initialize list_ctrl, insert column into each list_ctrl, set width;
        call updateSuiteList() and updateTestList()
        """
        # ## monitor information

        self.clientCommandList.InsertColumn(0, u"C_ID")
        self.clientCommandList.InsertColumn(1, u"C_Command")
        self.clientCommandList.InsertColumn(2, u"C_Content")

        self.clientCommandList.SetColumnWidth(0, 80)
        self.clientCommandList.SetColumnWidth(1, 130)
        self.clientCommandList.SetColumnWidth(2, 350)

        self.serverCommandList.InsertColumn(0, u"S_ID")
        self.serverCommandList.InsertColumn(1, u"S_Command")
        self.serverCommandList.InsertColumn(2, u"S_Content")

        self.serverCommandList.SetColumnWidth(0, 80)
        self.serverCommandList.SetColumnWidth(1, 130)
        self.serverCommandList.SetColumnWidth(2, 350)

        self.testsetList.InsertColumn(0,u"Set_ID")
        self.testsetList.InsertColumn(1,u"Content")
        self.testsetList.SetColumnWidth(0, 80)
        self.testsetList.SetColumnWidth(1, 430)

        self.testsetsubList.InsertColumn(0,u"Cmd_ID")
        self.testsetsubList.InsertColumn(1,u"Command")
        self.testsetsubList.SetColumnWidth(0, 80)
        self.testsetsubList.SetColumnWidth(1, 430)


        # add list all for the hahha addd here.................


        pass

    def initListFirstTime(self):
        # 这里添加首次更新的代码，主要流程为，用户点击连接按钮之后，连接成功，发送此函数，向服务器发送请求。

        er = ExcelReader()
        self.exceldict = er.excelData()
        self.OnRefreshList(wx.Event)
        self.OnSetListRefresh()

        pass


    def initLogging(self):
        '''
        Add self.msgTxtCtrl to logging system
        '''
        logger = logging.getLogger()
        handler = wxLogText(self.msgTextCtrl)
        formatter = logging.Formatter('%(levelname)s :%(asctime)s:  %(message)s')
        # #formatter = logging.Formatter('%(levelname)s :%(asctime)s %(filename)s %(lineno)s :  %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        '''
        Add debug.log to logging system
        '''
        logFile = "debug.log"
        handlerf = logging.FileHandler(logFile, 'w')
        handlerf.setFormatter(formatter)
        logger.addHandler(handlerf)
        '''
        set the level of the logger
        '''
        logger.setLevel(logging.INFO)
        pass


    def showDlg(self, content, title):
        """
        show information and error dialog
        """
        dlg = None
        if title == wx.ICON_INFORMATION:
            dlg = wx.MessageDialog(self, content, "Information", wx.OK | wx.ICON_INFORMATION)
        elif title == wx.ICON_ERROR:
            dlg = wx.MessageDialog(self, content, "Error", wx.OK | wx.ICON_ERROR)
        elif title == wx.ICON_NONE:
            dlg = wx.MessageDialog(self, content, "Infomation", wx.OK | wx.ICON_NONE)
        dlg.ShowModal()
        dlg.Destroy()

    def OnClose(self, event):
        """
        call when closing frame
        """
        # if hasattr(self,"netThread"):
        # self.showDlg("Thread is running, please stop first!", wx.ICON_ERROR)
        # return
        dlg = wx.MessageDialog(self, "Are you sure to exit??",
                               titlestring,
                               wx.YES_NO | wx.YES_DEFAULT | wx.ICON_INFORMATION
        )
        if dlg.ShowModal() == wx.ID_YES:
            self.Destroy()




class MonitorFrame2(wx.Frame):
    def __init__(self, parent):
        """
        initialize frame
        """

        wx.Frame.__init__(self, parent, -1, titlestring, size=(330, 600),
                          style=wx.CAPTION | wx.CLOSE_BOX | wx.MINIMIZE_BOX)
        self.CenterOnScreen()
        self.initData()
        self.initUI()
        self.initLayout()
        self.initCommandList()




    def initData(self):

        self.buttonDefaultSize = (88, 30)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        er = ExcelReader()
        self.exceldict = er.excelData()
        self.shortremember = []

        pass


    def initUI(self):
        """
        initialize book_page, include quest_page, monitor_page, capture_page, date_base_page, test_suite_page, about_page
            quest_page, monitor_page includes grid
            capture_page, date_base_page, test_suite_page includes list_ctrl and tool_panel, list_ctrl display the information,
            tool_panel include the buttons
        """

        self.CreateStatusBar()
        self.SetStatusText("Welcome to GM Command Tool!")

        menuBar = wx.MenuBar()

        menu = wx.Menu()

        self.menuCloseID = wx.NewId()

        menu.Append(self.menuCloseID, "&Exit", "Exit.")
        self.Bind(wx.EVT_MENU, self.OnClose, id=self.menuCloseID)
        #self.Bind(wx.EVT_MENU, self.OnMenuStartGetting, id=self.menuStartGet)
        #self.Bind(wx.EVT_MENU, self.OnMenuStopGetting, id=self.menuStopGet)

        menuBar.Append(menu, "&Admin")
        self.SetMenuBar(menuBar)


        self.commandTextPanel = wx.Panel(self,-1)
        #self.command_Text = wx.TextCtrl(self.commandTextPanel, -1, "Command Show Here", size=(3000, -1))
        self.command_Text = wx.SearchCtrl(self.commandTextPanel, -1, "Command Show Here",size=(328, -1), style=wx.TE_PROCESS_ENTER)
        self.command_Text.SetMenu(self.MakeMenu())
        self.command_Text.ShowSearchButton(True)
        self.command_Text.ShowCancelButton(True)

        self.Bind(wx.EVT_TEXT,self.OnSearch, self.command_Text)
        self.Bind(wx.EVT_SEARCHCTRL_SEARCH_BTN, self.OnShowMenu, self.command_Text)
        self.Bind(wx.EVT_SEARCHCTRL_CANCEL_BTN, self.OnCancel, self.command_Text)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnEnterSearch, self.command_Text)


        self.listPanel = wx.Panel(self, -1)
        self.commandlist = EditListCtrl3(self.listPanel)
        self.commandlist.Bind(wx.EVT_LIST_ITEM_SELECTED, self.Oncommandselected)

        self.buttonPanel = wx.Panel(self, -1)
        self.runButton = wx.Button(self.buttonPanel, -1, u"Run",size = (-1,50))
        self.Bind(wx.EVT_BUTTON, self.OnRunButton, self.runButton)



        #
        # self.filediffPanel = wx.Panel(self, -1)
        # self.filediffPanel.Bind(wx.EVT_CONTEXT_MENU, self.OnContextMenu)
        #
        # self.filedifflist = EditListCtrl(self.filediffPanel)
        # self.filedifflist.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnfilediffItemSelected)
        #
        # self.toolPanel = wx.Panel(self.filediffPanel, -1)
        # self.refresh_button = wx.Button(self.toolPanel, -1, u"Refresh", size=(88, 50))
        # self.refresh_button.Enable(False)
        #
        # self.Bind(wx.EVT_BUTTON, self.OnRefreshButton, self.refresh_button)
        #
        # self.allconfirm_button = wx.Button(self.toolPanel, -1, u"AllConfirm", size=(88, 50))
        # self.Bind(wx.EVT_BUTTON, self.OnAllConfirmButton, self.allconfirm_button)
        # self.allconfirm_button.Enable(False)
        #
        # self.confirm_button = wx.Button(self.toolPanel, -1, u"Confirm", size=(88, 50))
        # self.Bind(wx.EVT_BUTTON, self.OnConfirmButton, self.confirm_button)
        # self.confirm_button.Enable(False)
        #
        # self.vacantPanel = wx.Panel(self.toolPanel, -1)
        #
        # self.ok_button = wx.Button(self.toolPanel, -1, u"Exit")
        # self.Bind(wx.EVT_BUTTON, self.OnOkButton, self.ok_button)


        pass

    def OnShowMenu(self,event):

        self.command_Text.SetMenu(self.MakeMenu())


    def OnEnterSearch(self,event):

        inputstr = self.command_Text.GetValue()
        self.dataforsearch_temp = []

        for i in self.dataforsearch_ori:
            if (inputstr in i[0]) or (inputstr in i[1]):
                self.dataforsearch_temp.append(i)
        self.commandlist.DeleteAllItems()

        for i in range(len(self.dataforsearch_temp)):
            index = self.commandlist.InsertStringItem(i, str(i))
            temp_command = self.dataforsearch_temp[i][0]
            temp_content = self.dataforsearch_temp[i][1]
            self.commandlist.SetStringItem(index, 0, temp_command)
            self.commandlist.SetStringItem(index, 1, temp_content)
        if len(self.dataforsearch_temp) > 0 and (inputstr not in self.shortremember) and (inputstr is not u""):
            self.shortremember.append(inputstr)

        print self.shortremember
        pass

    def OnSearch(self,event):

        #如果是点击，则先将此功能屏蔽掉。

        inputstr = self.command_Text.GetValue()
        self.dataforsearch_temp = []

        for i in self.dataforsearch_ori:
            if (inputstr in i[0]) or (inputstr in i[1]):
                self.dataforsearch_temp.append(i)
        #print "OnSearch2"
        self.commandlist.DeleteAllItems()
        for i in range(len(self.dataforsearch_temp)):
            index = self.commandlist.InsertStringItem(i, str(i))
            temp_command = self.dataforsearch_temp[i][0]
            temp_content = self.dataforsearch_temp[i][1]
            self.commandlist.SetStringItem(index, 0, temp_command)
            self.commandlist.SetStringItem(index, 1, temp_content)
        self.dataforsearch_temp = []
        self.command_to_run = inputstr


    def OnCancel(self,event):

        self.initCommandList()


    def OnMenuClicked(self,event):

        itemid = event.GetId()
        Item = self.menu.FindItemById(itemid)
        self.command_Text.SetValue(Item.GetLabel())

        pass

    def MakeMenu(self):
        self.menu = wx.Menu()
        item = self.menu.Append(-1, "Recent Searches")
        item.Enable(False)
        for txt in self.shortremember:
            menuItem = self.menu.Append(-1, txt)
            self.Bind(wx.EVT_MENU, self.OnMenuClicked,menuItem)
        return self.menu



    def initLayout(self):
        """
        set layout of the picture_ctrl, page panel
        """
        self._icon = _icon = wx.EmptyIcon()
        #wx.Image("./pic/gmlogo.png", wx.BITMAP_TYPE_PNG).ConvertToBitmap()
        _icon.LoadFile("./pic/pc.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(_icon)

        commandPanelSizer = wx.BoxSizer(wx.VERTICAL)
        #commandPanelSizer.Add(self.command_Text,0,wx.ALL | wx.EXPAND, 0)

        self.commandlist.SetSizer(commandPanelSizer)

        listPanelSizer = wx.BoxSizer(wx.VERTICAL)
        listPanelSizer.Add(self.commandlist,10,wx.ALL | wx.EXPAND, 0)
        self.listPanel.SetSizer(listPanelSizer)


        buttonPanelSizer = wx.BoxSizer(wx.VERTICAL)
        buttonPanelSizer.Add(self.runButton,0,wx.ALL | wx.EXPAND, 0)
        self.buttonPanel.SetSizer(buttonPanelSizer)

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.commandTextPanel,0.5, wx.ALL | wx.EXPAND, 0)
        mainSizer.Add(self.listPanel, 9, wx.ALL | wx.EXPAND, 0)
        mainSizer.Add(self.buttonPanel,1,wx.ALL | wx.EXPAND, 0)

        self.SetSizer(mainSizer)

        pass

    def doServerCommand(self,command):

        # This function is used to do server commands
        try:
            from utils import monkey_patch
            import x9
            from x9.system import GmCmd
            from x9 import BigWorld as BW

            self.BigWorld = BW
            #print "111111111111111111111"
            #print command,"12341234"
            #self.BigWorld.player().base.doGmCommand("$warn 小明打豆豆")
            #print command,"**"
            #self.BigWorld.player().base.doGmCommand(str("$alertbulletin 中文测试").decode('utf-8').encode('gbk'))
            #command = "$alertbulletin 中文测试inner"

            #self.BigWorld.player().base.doGmCommand(str(command).decode('utf-8').encode('gbk'))
            self.BigWorld.player().base.doGmCommand(command.encode('gbk'))
            #print "1"


            #self.BigWorld.player().base.doGmCommand("$bulletin")
            #self.BigWorld.player().base.doGmCommand(u"$bulletin 中文")
            #print "@@@@@@@"

        except Exception, err:

            if str(err) == "No module named x9":
                # print "Open X9 and restart this tool!!"
                self.showDlg("Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return

            elif str(err) == "stream has been closed":
                # print "X9 has been CLOSED!!"
                self.showDlg("X9 Has Been CLOSED!!Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return

    def doClientCommand(self,command):

        # This funcion is used to do client commands
        try:
            from utils import monkey_patch
            import x9
            from x9.system import GmCmd
            from x9 import BigWorld as BW

            self.BigWorld = BW


            GmCmd.prase_cmd(command.encode('gbk'))

        except Exception, err:

            if str(err) == "No module named x9":
                # print "Open X9 and restart this tool!!"
                self.showDlg("Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return
            elif str(err) == "stream has been closed":
                # print "X9 has been CLOSED!!"
                self.showDlg("X9 Has Been CLOSED!!Open X9 And RESTART This Tool!!", wx.ICON_ERROR)
                self.Destroy()
                return

    def doAllCommand(self,command):

        if command.startswith("$"):
            print "server",command
            #print command
            self.doServerCommand(command)
        elif command.startswith("/"):
            print "client",command
            self.doClientCommand(command)
        else:
            print "else"
            print command


    def OnRunButton(self,event):

        #此处添加读取搜索框的代码。
        command = command1 = self.command_Text.GetValue()

        if self.command_to_run is not None:
            command = self.command_to_run
        #self.doAllCommand(self.command_to_run)

        print command1,"herehere",command

        self.doAllCommand(command)
        pass

    def Oncommandselected(self,event):

        #self.clickflag = 1 # 点击。

        #print "selected"
        # print int(event.m_itemIndex + 1)
        # print self.lenofclientData
        # print self.commandlist.GetItem(event.m_itemIndex,0).GetText()

        command_temp = self.commandlist.GetItem(event.m_itemIndex,0).GetText()
        content_temp = self.commandlist.GetItem(event.m_itemIndex,1).GetText()


        clientdict = self.exceldict.get("client")
        serverdict = self.exceldict.get("server")

        self.indextouse = None
        self.command_to_run = None

        if command_temp.startswith("/"):
            for keyinside in self.exceldict.get("client").keys():
                if command_temp == clientdict.get(keyinside).get(u"command") and content_temp == clientdict.get(keyinside).get(u'command_content'):
                    self.indextouse = int(clientdict.get(keyinside).get(u'command_id'))
                    ###############################
                    command_para_number = clientdict.get(keyinside).get("command_para_num")
                    #print command_para_number
                    if isinstance(command_para_number, float)  and 0 <= int(command_para_number) <= 7:

                        parameters2 = []
                        for i in range(int(command_para_number)):
                            keytouse = "para_" + str(i + 1)
                            para_default = self.exceldict.get("clientdefault").get(self.indextouse).get(keytouse)

                            if isinstance(para_default, float):
                                para_default = int(para_default)
                            elif isinstance(para_default, int):
                                para_default = int(para_default)
                            else:
                                pass
                            # parameters[para_name] = str(para_default)
                            parameters2.append(unicode(para_default))
                            #print para_name,para_default

                            self.command_to_run = self.makecommand2(command_temp, parameters2)
                        if self.command_to_run == None:
                            self.command_to_run = command_temp
                        print self.command_to_run
                            #self.command_Text.SetLabel(self.command_to_run)


            pass

        elif command_temp.startswith("$"):
            for keyinside in self.exceldict.get("server").keys():
                if command_temp == serverdict.get(keyinside).get(u"command") and content_temp == serverdict.get(keyinside).get(u'command_content'):
                    self.indextouse = serverdict.get(keyinside).get(u'command_id')


                    command_para_number = serverdict.get(keyinside).get("command_para_num")
                    #print command_para_number
                    if isinstance(command_para_number, float)  and 0 <= int(command_para_number) <= 7:

                        parameters2 = []
                        for i in range(int(command_para_number)):
                            keytouse = "para_" + str(i + 1)
                            para_default = self.exceldict.get("serverdefault").get(self.indextouse).get(keytouse)

                            if isinstance(para_default, float):
                                para_default = int(para_default)
                            elif isinstance(para_default, int):
                                para_default = int(para_default)
                            else:
                                pass

                            parameters2.append(unicode(para_default))
                            self.command_to_run = self.makecommand2(command_temp, parameters2)
                        if self.command_to_run == None:
                            self.command_to_run = command_temp
                        print self.command_to_run

            pass

        else:
            pass
        print self.command_to_run
        self.command_Text.ChangeValue(self.command_to_run)

        #print "testValue",self.command_Text.GetValue()


    def makecommand2(self, command, list):
        s = command
        if s is not None:
            if list is not None:
                for i in list:
                    s = s + " " + i
            else:
                s = command
        else:
            s = command
        return s

    def initCommandList(self):

        self.dataforsearch_ori = []

        clientData = self.exceldict.get(u"client")

        self.commandlist.DeleteAllItems()

        self.lenofclientData = len(clientData)

        for i in range(len(clientData)):
            index = self.commandlist.InsertStringItem(i, str(i))

            temp_command = clientData.get(i + 1).get(u'command')
            temp_content = clientData.get(i + 1).get(u'command_content')


            self.commandlist.SetStringItem(index, 0, temp_command)
            self.commandlist.SetStringItem(index, 1, temp_content)
            self.dataforsearch_ori.append([temp_command,temp_content])


        serverData = self.exceldict.get(u"server")

        for i in range(len(serverData)):
            index = self.commandlist.InsertStringItem(self.lenofclientData+i, str(self.lenofclientData+i))

            temp_command = serverData.get(i + 1).get(u'command')
            temp_content = serverData.get(i + 1).get(u'command_content')

            self.commandlist.SetStringItem(index, 0, temp_command)
            self.commandlist.SetStringItem(index, 1, temp_content)
            self.dataforsearch_ori.append([temp_command,temp_content])

        #print self.dataforsearch_ori



    def showDlg(self, content, title):
        """
        show information and error dialog
        """
        dlg = None
        if title == wx.ICON_INFORMATION:
            dlg = wx.MessageDialog(self, content, "Information", wx.OK | wx.ICON_INFORMATION)
        elif title == wx.ICON_ERROR:
            dlg = wx.MessageDialog(self, content, "Error", wx.OK | wx.ICON_ERROR)
        elif title == wx.ICON_NONE:
            dlg = wx.MessageDialog(self, content, "Infomation", wx.OK | wx.ICON_NONE)
        dlg.ShowModal()
        dlg.Destroy()

    def OnClose(self, event):
        """
        call when closing frame
        """
        # if hasattr(self,"netThread"):
        # self.showDlg("Thread is running, please stop first!", wx.ICON_ERROR)
        # return
        dlg = wx.MessageDialog(self, "Are you sure to exit??",
                               titlestring,
                               wx.YES_NO | wx.YES_DEFAULT | wx.ICON_INFORMATION
        )
        if dlg.ShowModal() == wx.ID_YES:
            self.Destroy()

class MainFrame(wx.Frame):

    def __init__(self, parent):
        """
        initialize frame
        """

        wx.Frame.__init__(self, parent, -1, titlestring, size=(0, 0),
                          style=wx.MINIMIZE_BOX)
        self.CenterOnScreen()
        dlg = wx.SingleChoiceDialog(
                self, u'请选择您想要使用的版本：\n  简化版：功能简化,适合单屏幕\n  扩展版：功能稍多,适合多屏幕\n\n        确认前须开启客户端', u'X9 GM指令工具：',
                [u'简化版', u'扩展版'],
                wx.CHOICEDLG_STYLE
                )
        if dlg.ShowModal() == wx.ID_OK:
            selected = dlg.GetStringSelection()
            if selected == u"简化版":
                newframe = MonitorFrame2(None)
                newframe.Show()
            elif selected == u'扩展版':
                newframe = MonitorFrame(None)
                newframe.Show()
            else:
                print "Error"
        else:
            pass

        dlg.Destroy()
        self.Destroy()


if __name__ == '__main__':

    app = wx.App(redirect=False)
    frame = MainFrame(None)
    frame.CenterOnScreen()
    app.SetTopWindow(frame)
    frame.Show()
    app.MainLoop()
