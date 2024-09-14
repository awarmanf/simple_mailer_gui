#!/usr/bin/python
#
# simpleMailer
#
# License
#   This program is copyleft. You have the right to freely use, modify,
#   copy, and share software, works of art, etc., on the condition that
#   these rights be granted to all subsequent users or owners. 
#
# Author
#   Arief Yudhawarman <awarmanf@yahoo.com>
#

import wx
import smtplib
import sys

global VERSION, WIDTH, HEIGHT

class myFrame(wx.Frame):

    def __init__(self, parent, title):
        wx.Frame.__init__(self, None, wx.ID_ANY, title=title, size=(WIDTH, HEIGHT))
            
        self.InitUI()
        self.Centre()
        self.Show()     

    def InitUI(self):
        global VERSION

        self.stBar = self.CreateStatusBar(1, 0)

        frame_1_statusbar_fields = ["Waiting for input..."]
        for i in range(len(frame_1_statusbar_fields)):
            self.stBar.SetStatusText(frame_1_statusbar_fields[i], i)

        panel = wx.Panel(self)

        gbSizer = wx.GridBagSizer(13, 4)

        header  = wx.StaticText(panel, label="Simple Mailer")
        header.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD, 0, ""))
        gbSizer.Add(header, pos=(0, 0), span=(1,4),
            flag=wx.ALIGN_CENTER_HORIZONTAL|wx.TOP|wx.BOTTOM, border=15)

        line = wx.StaticLine(panel)
        gbSizer.Add(line, pos=(1, 0), span=(1,4),
            flag=wx.EXPAND|wx.BOTTOM|wx.LEFT|wx.RIGHT,border=5)

        lblHost    = wx.StaticText(panel, label="SMTP Host")
        gbSizer.Add(lblHost, pos=(2, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtHost = wx.TextCtrl(panel)
        gbSizer.Add(self.txtHost, pos=(2, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        lblAccount = wx.StaticText(panel, label="Account Email")
        gbSizer.Add(lblAccount, pos=(3, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtAccount = wx.TextCtrl(panel)
        gbSizer.Add(self.txtAccount, pos=(3, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        lblPasswd  = wx.StaticText(panel, label="Password")
        gbSizer.Add(lblPasswd, pos=(4, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtPasswd = wx.TextCtrl(panel,style=wx.TE_PASSWORD)
        gbSizer.Add(self.txtPasswd, pos=(4, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        line2 = wx.StaticLine(panel)
        gbSizer.Add(line2, pos=(5, 0), span=(1, 4), 
            flag=wx.EXPAND|wx.BOTTOM|wx.LEFT|wx.RIGHT,border=5)


        lblSubject = wx.StaticText(panel, label="Subject")
        gbSizer.Add(lblSubject, pos=(6, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtSubject = wx.TextCtrl(panel)
        gbSizer.Add(self.txtSubject, pos=(6, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        lblRcpts   = wx.StaticText(panel, label="Recipients")
        gbSizer.Add(lblRcpts, pos=(7, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtRcpts = wx.TextCtrl(panel)
        gbSizer.Add(self.txtRcpts, pos=(7, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)


        lblBody = wx.StaticText(panel, label="Body:")
        gbSizer.Add(lblBody, pos=(8, 0), flag=wx.LEFT, border=10)

        self.txtBody = wx.TextCtrl(panel,-1,"\n---\n%s"%VERSION,style=wx.TE_MULTILINE)
        gbSizer.Add(self.txtBody, pos=(9, 0), span=(3, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)

        btnOk = wx.Button(panel, label="Ok", size=(90, 28))
        btnClose = wx.Button(panel, label="Close", size=(90, 28))
        self.Bind(wx.EVT_BUTTON, self.onOK, btnOk)
        self.Bind(wx.EVT_BUTTON, self.onClose, btnClose)

        gbSizer.Add(btnOk, pos=(12, 2))
        gbSizer.Add(btnClose, pos=(12, 3), flag=wx.RIGHT|wx.BOTTOM, border=10)

        gbSizer.AddGrowableCol(1)
        gbSizer.AddGrowableRow(10)
        panel.SetSizerAndFit(gbSizer)

    def onOK(self, event):

        HOST=''
        USER=''
        PASS=''
        SUBJECT=''
        RCPTS=''
        BODY=''

        if self.txtHost.GetValue() != "":
            HOST = self.txtHost.GetValue()

        else:
            wx.MessageBox("SMTP Host can't be emptied", "Warning", wx.OK | wx.ICON_WARNING)
        if self.txtAccount.GetValue() != "":
            USER = self.txtAccount.GetValue()

        else:
            wx.MessageBox("Account Email can't be emptied", "Warning", wx.OK | wx.ICON_WARNING)
        if self.txtPasswd.GetValue() != "":
            PASS = self.txtPasswd.GetValue()

        else:
            wx.MessageBox("Password can't be emptied", "Warning", wx.OK | wx.ICON_WARNING)
        if self.txtSubject.GetValue() != "":
            SUBJECT = self.txtSubject.GetValue()

        else:
            wx.MessageBox("Subject can't be emptied", "Warning", wx.OK | wx.ICON_WARNING)
        if self.txtRcpts.GetValue() != "":
            RCPTS = self.txtRcpts.GetValue()

        else:
            wx.MessageBox("Recipients can't be emptied", "Warning", wx.OK | wx.ICON_WARNING) 
        if self.txtBody.GetValue() != "":
            BODY = "\r\n".join((
                "From: %s" % USER,
                "To: %s" % RCPTS,
                "Subject: %s" % SUBJECT ,
                "X-Mailer: %s" % VERSION,
                "",
                self.txtBody.GetValue()))

        else:
            wx.MessageBox("Body email can't be emptied", "Warning", wx.OK | wx.ICON_WARNING)

        if (HOST != '') & (USER != '') & (PASS != '') & (SUBJECT != '') & (RCPTS != '') & (BODY != ''):
            self.stBar.SetStatusText("Sending email...")
            try:
                server = smtplib.SMTP(HOST,port=25,timeout=10)
            except:
                wx.MessageBox("Connection Error", "Warning", wx.OK | wx.ICON_WARNING)
            else:
                try:
                    server.login(USER, PASS)
                except:
                    wx.MessageBox("Authentication Failed", "Warning", wx.OK | wx.ICON_WARNING)
                else:
                    server.sendmail(USER, [RCPTS], BODY)
                    server.quit()
                    self.stBar.SetStatusText("Send email was success.")
                    wx.MessageBox('Your email has been sent.', 'Sent Email', wx.OK)
                    self.stBar.SetStatusText("Ready to send again.")

    def onClose(self, event):
        self.Close()


if __name__ == '__main__':

    WIDTH = 400
    HEIGHT = 500

    if (sys.platform == "win32"):
        WIDTH = 400
        HEIGHT = 625

    VERSION = 'Simple Mailer v1.0'
    app = wx.App()
    myFrame(None, title='Simple Mailer')
    app.MainLoop()


