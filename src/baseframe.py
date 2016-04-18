#!/bin/python
# -*- coding: utf-8 -*-
# mode:python; tab-width:4 -*- ex:set tabstop=4 shiftwidth=4 expandtab: -*-

#-------------------------------------------------------------------------
# Name:         baseframe.py
# Purpose:      base class of frame
#
# Author:       ma xiao
# Created:      2011-1-20
# Copyright:    iscas
#-------------------------------------------------------------------------

"""
Pydoc
…
"""

__version__ = "1.0"

import wx
from wx import xrc
import wx.lib.mixins.listctrl  as  listmix
import logging
import thread


class XrcBase:
    def get_ctrls(self, cNames):
        for name in cNames:
            ctrl = xrc.XRCCTRL(self, name)
            if ctrl is None:
                logging("%s can't be found" % (name))
            setattr(self, name, ctrl)

    def init_ctrls(self):
        pass

class XrcDialog(wx.Dialog, XrcBase):
    def __init__(self, parent, resource, name):
        self.res = resource
        w = resource.LoadDialog(parent, name)
        self.PostCreate(w)
        self.init_ctrls()


class EditListCtrl(wx.ListCtrl):
    def __init__(self, parent):
        wx.ListCtrl.__init__(
            self, parent, -1,
            style=wx.LC_REPORT | wx.LC_EDIT_LABELS

        )
        #TextEditMixin.__init__(self)

        #self.SetItemCount(20)
        #self.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown)
        # self.prevalue = -1
        # self.curRow=-1
        # self.curCol=-1



        # def OnLeftDown(self, evt=None):
        #     ''' Examine the click and double
        #     click events to see if a row has been click on twice. If so,
        #     determine the current row and column and open the editor.'''
        #     TextEditMixin.OnLeftDown(self,evt)


class EditListCtrl2(wx.ListCtrl,
                   listmix.ListCtrlAutoWidthMixin,
                   listmix.TextEditMixin):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent,-1,style=wx.LC_REPORT)

        listmix.ListCtrlAutoWidthMixin.__init__(self)
        self.Populate()
        #listmix.TextEditMixin.__init__(self)

    def Populate(self):
        # for normal, simple columns, you can add them like this:
        self.InsertColumn(0, "Name")
        self.InsertColumn(1, "Parameter")

        self.SetColumnWidth(0, 150)
        self.SetColumnWidth(1, 250)

        self.currentItem = 0

class EditListCtrl3(wx.ListCtrl,
                   listmix.ListCtrlAutoWidthMixin,
                   listmix.TextEditMixin):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent,-1,style=wx.LC_REPORT)

        listmix.ListCtrlAutoWidthMixin.__init__(self)
        self.Populate()
        #listmix.TextEditMixin.__init__(self)

    def Populate(self):
        # for normal, simple columns, you can add them like this:
        self.InsertColumn(0, "Command")
        self.InsertColumn(1, "Content")

        self.SetColumnWidth(0, 80)
        self.SetColumnWidth(1, 300)

        self.currentItem = 0


class wxLogText(logging.Handler):
    def __init__(self, textCtl):
        logging.Handler.__init__(self)
        self.textCtl = textCtl
        self.thread_id = thread.get_ident()

    def emit(self, record):
        if thread.get_ident() == self.thread_id:
            #GUI
            self.textCtl.AppendText(self.format(record))
            self.textCtl.AppendText('\n')
        else:
            wx.CallAfter(self.textCtl.AppendText, self.format(record) + "\n")

