# -*- coding: utf-8 -*-

import wx
from wx import xrc

class XrcBase:
    def get_ctrls(self, cNames):
        for name in cNames:
            ctrl = xrc.XRCCTRL(self, name)
            if ctrl is None:
                ##logging("%s can't be found" % (name))
                pass
            setattr(self, name, ctrl)

    def init_ctrls(self):
        pass

class XrcDialog(wx.Dialog, XrcBase):
    def __init__(self, parent, resource, name):
        self.res = resource
        w = resource.LoadDialog(parent, name)
        self.PostCreate(w)
        self.init_ctrls()


class AddSuiteDlg(XrcDialog):

    def __init__(self, parent, id,**kw):
        '''
         xrc file
        '''
        self.id = id
        dlgResource = xrc.XmlResource("addsuitedlg.xrc")
        XrcDialog.__init__(self, parent, dlgResource, "addSuiteDlg")

    def init_ctrls(self):
        '''
        get ctrls
        '''
        ctrls = ['idTxt', 'ipTxt',"okBtn","cancelBtn"]
        self.get_ctrls(ctrls)
        self.idTxt.SetValue(str(self.id))
        self.okBtn.SetId(wx.ID_OK)
        self.okBtn.Enable(False)
        self.cancelBtn.SetId(wx.ID_CANCEL)
        self.Bind(wx.EVT_TEXT, self.EvtText, self.ipTxt)

    def EvtText(self, event):
        '''
        enable ok button when insert experiment name
        '''
        if event.GetString() != "":
            self.okBtn.Enable(True)
        else:
            self.okBtn.Enable(False)

    def getData(self):
        '''
        get data
        '''
        return self.ipTxt.GetValue()


class AddSubSuiteDlg(XrcDialog):

    def __init__(self, parent, id,**kw):
        '''
         xrc file
        '''
        self.id = id
        dlgResource = xrc.XmlResource("addsubsuitedlg.xrc")
        XrcDialog.__init__(self, parent, dlgResource, "addSuiteDlg")

    def init_ctrls(self):
        '''
        get ctrls
        '''
        ctrls = ['idTxt', 'ipTxt',"okBtn","cancelBtn"]
        self.get_ctrls(ctrls)
        self.idTxt.SetValue(str(self.id))
        self.okBtn.SetId(wx.ID_OK)
        self.okBtn.Enable(False)
        self.cancelBtn.SetId(wx.ID_CANCEL)
        self.Bind(wx.EVT_TEXT, self.EvtText, self.ipTxt)

    def EvtText(self, event):
        '''
        enable ok button when insert experiment name
        '''
        if event.GetString() != "":
            self.okBtn.Enable(True)
        else:
            self.okBtn.Enable(False)

    def getData(self):
        '''
        get data
        '''
        return self.ipTxt.GetValue()


class SelectPanel(wx.Panel):
    def __init__(self, parent, log):
        self.log = log
        wx.Panel.__init__(self, parent, -1)

        b = wx.Button(self, -1, "Create and Show a SingleChoiceDialog", (50,50))
        self.Bind(wx.EVT_BUTTON, self.OnButton, b)


    def OnButton(self, evt):
        dlg = wx.SingleChoiceDialog(
                self, 'Test Single Choice', 'The Caption',
                ['zero', 'one'],
                wx.CHOICEDLG_STYLE
                )

        if dlg.ShowModal() == wx.ID_OK:
            self.log.WriteText('You selected: %s\n' % dlg.GetStringSelection())

        dlg.Destroy()
