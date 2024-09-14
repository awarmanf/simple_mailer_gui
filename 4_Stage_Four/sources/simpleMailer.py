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
import re
from os import path
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
import ConfigParser
import sys
import random
import time
import myImages

OS = sys.platform

class myFrame(wx.Frame):

    def __init__(self, parent, title, fconfig):
        
        wx.Frame.__init__(self, None, wx.ID_ANY, title=title) # or set pos=(30,140)
        self.appname = 'Simple Mailer'
        self.version = '1.2'
        self.thisrelease = '2018'
        self.identity = 'Arief Yudhawarman'
        self.email = 'awarmanf@yahoo.com'
        self.fconfig = fconfig
        if not path.exists(self.fconfig):
            wx.MessageBox("File config %s not found." % self.fconfig, "Warning", wx.OK | wx.ICON_WARNING)
            self.Close()
        else:
            self.config = ConfigParser.ConfigParser()
            self.config.read(self.fconfig)
            self.sections = self.config.sections()[1:]
            self.name = self.config.get("Identity", "name")
            self.signature = self.config.get("Identity", "signature")
            self.InitUI()
            self.Centre()
            self.Show()     


    def InitUI(self):
        
        menu = wx.Menu()

        self.itemOpen = wx.MenuItem(menu, wx.ID_ANY, "&Open\tCtrl+O","Open File Attachment")
        self.itemOpen.SetBitmap(myImages.open_item.GetBitmap())
        if OS == 'win32':
            menu.Append(self.itemOpen)
        else:
            menu.AppendItem(self.itemOpen)

        self.itemAbout = wx.MenuItem(menu, wx.ID_ANY, "&About", "About this program")
        self.itemAbout.SetBitmap(myImages.about_item.GetBitmap())
        if OS == 'win32':
            menu.Append(self.itemAbout)
        else:
            menu.AppendItem(self.itemAbout)

        menu.AppendSeparator()

        self.itemExit = wx.MenuItem(menu, wx.ID_ANY,"&Exit\tCtrl+Q","Exit program")
        self.itemExit.SetBitmap(myImages.exit_item.GetBitmap())
        if OS == 'win32':
            menu.Append(self.itemExit)
        else:
            menu.AppendItem(self.itemExit)

        menuBar = wx.MenuBar()
        menuBar.Append(menu,'File')
        self.SetMenuBar(menuBar)

        self.icon = myImages.email_icon.GetIcon()
        self.SetIcon(self.icon)

        txtSize = (200,-1)

        self.stBar = self.CreateStatusBar()
        self.stBar.SetStatusText("Waiting for input...")

        self.panel = wx.Panel(self)
    
        self.lblHeader  = wx.StaticText(self.panel, label="Simple Mailer")
        self.lblHeader.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD, 0, ""))
        self.line = wx.StaticLine(self.panel)
        
        self.lblFullName = wx.StaticText(self.panel, label="Full Name")
        self.txtFullName = wx.TextCtrl(self.panel,value=self.name,size=txtSize)
        
        self.lblAccount = wx.StaticText(self.panel, -1, "Account")
        self.cmbAccount = wx.ComboBox(self.panel, -1, choices=self.sections, style=wx.CB_DROPDOWN | wx.CB_READONLY, size=(215,28))
        
        self.lblPass  = wx.StaticText(self.panel, label="Password")
        self.txtPass = wx.TextCtrl(self.panel,style=wx.TE_PASSWORD,size=txtSize)
        self.txtPassShow = wx.TextCtrl(self.panel,size=txtSize)
        self.txtPassShow.Hide()

        bmp = myImages.eye_icon.GetBitmap()
        self.btnShow = wx.BitmapButton(self.panel, id=wx.ID_ANY, bitmap=bmp)
        self.btnShow.SetToolTip(wx.ToolTip("Show/hide password"))

        self.lblHost    = wx.StaticText(self.panel, label="SMTP Host")
        self.txtHost = wx.TextCtrl(self.panel,size=txtSize)

        self.lblSSL  = wx.StaticText(self.panel, label="Security Type")
        self.rdoBtn1 = wx.RadioButton(self.panel, label='Plain',style=wx.RB_GROUP)
        self.rdoBtn2 = wx.RadioButton(self.panel, label='TLS')
        self.rdoBtn3 = wx.RadioButton(self.panel, label='SSL')

        self.lblSubject = wx.StaticText(self.panel, label="Subject")
        self.txtSubject = wx.TextCtrl(self.panel)
        self.txtSubject.SetFocus()
        
        self.lblRcpts   = wx.StaticText(self.panel, label="Recipients")
        self.txtRcpts = wx.TextCtrl(self.panel)
        self.txtRcpts.SetToolTip(wx.ToolTip("Add another recipient separated with comma"))        
        
        self.lblFile = wx.StaticText(self.panel, label="File Attachment")        
        self.txtFile = wx.TextCtrl(self.panel)

        self.lblBody = wx.StaticText(self.panel, label="Body:")
        self.txtBody = wx.TextCtrl(self.panel,-1,self.signature,style=wx.TE_MULTILINE)

        self.btnSubmit = wx.Button(self.panel, label="Submit", size=(90, 28))
        self.btnClose = wx.Button(self.panel, label="Close", size=(90, 28))
        
        self.__set_properties()
        self.__do_layout()
        
        self.Bind(wx.EVT_MENU, self.OnOpen, self.itemOpen)
        self.Bind(wx.EVT_MENU, self.OnAbout, self.itemAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, self.itemExit)
        self.Bind(wx.EVT_COMBOBOX, self.OnSelect, self.cmbAccount)
        self.Bind(wx.EVT_BUTTON, self.togglePassword, self.btnShow)
        self.Bind(wx.EVT_BUTTON, self.onSubmit, self.btnSubmit)
        self.Bind(wx.EVT_BUTTON, self.onClose, self.btnClose)


    def __set_properties(self):
        self.selected = 0
        self.password_shown= False
        self.security = 'Plain'
        self.txtCtrlPass = self.txtPass
        self.cmbAccount.SetSelection(self.selected)
        self.txtHost.SetValue(self.config.get(self.sections[self.selected],'host'))
        self.txtCtrlPass.SetValue(self.config.get(self.sections[self.selected],'password'))
        self.txtBody.SetFont(wx.Font(9, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, ""))


    def __do_layout(self):
        
        topSizer = wx.BoxSizer(wx.VERTICAL)
        rdoBtnSizer = wx.BoxSizer(wx.HORIZONTAL)
        passSizer = wx.BoxSizer(wx.HORIZONTAL)

        if OS == 'win32':
            gbs_x = 13
            txtBody_spanx = 1
            btnSubmit_posx = 12
            btnClose_posx = 12
            GrowableRow = 11
        else:
            gbs_x = 15
            txtBody_spanx = 3
            btnSubmit_posx = 14
            btnClose_posx = 14
            GrowableRow = 12

        gbSizer = wx.GridBagSizer(gbs_x, 4)
        
        gbSizer.Add(self.lblHeader, pos=(0, 0), span=(1,4),
            flag=wx.ALIGN_CENTER_HORIZONTAL|wx.TOP|wx.BOTTOM, border=10)
        gbSizer.Add(self.line, pos=(1, 0), span=(1,4),
            flag=wx.EXPAND|wx.BOTTOM|wx.LEFT|wx.RIGHT,border=5)
        
        gbSizer.Add(self.lblFullName, pos=(2, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        gbSizer.Add(self.txtFullName, pos=(2, 1), span=(1, 3), flag=wx.RIGHT, border=20)
        
        gbSizer.Add(self.lblAccount, pos=(3, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        gbSizer.Add(self.cmbAccount, pos=(3, 1), span=(1, 3), 
            flag=wx.LEFT|wx.RIGHT, border=5)
        
        gbSizer.Add(self.lblPass, pos=(4, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        passSizer.Add(self.txtPass)
        passSizer.Add(self.txtPassShow)
        passSizer.Add(self.btnShow,flag=wx.LEFT,border=5)
        gbSizer.Add(passSizer, pos=(4, 1), span=(1, 3), flag=wx.LEFT|wx.RIGHT, border=5)

        gbSizer.Add(self.lblHost, pos=(5, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        gbSizer.Add(self.txtHost, pos=(5, 1), span=(1, 3), 
            flag=wx.LEFT|wx.RIGHT, border=5)

        gbSizer.Add(self.lblSSL, pos=(6, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        rdoBtnSizer.Add(self.rdoBtn1,flag=wx.RIGHT,border=5)
        rdoBtnSizer.Add(self.rdoBtn2,flag=wx.RIGHT,border=5)
        rdoBtnSizer.Add(self.rdoBtn3)
        gbSizer.Add(rdoBtnSizer, pos=(6,1), span=(1, 3), flag=wx.LEFT|wx.RIGHT, border=5)

        gbSizer.Add(self.lblSubject, pos=(7, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        gbSizer.Add(self.txtSubject, pos=(7, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)        

        gbSizer.Add(self.lblRcpts, pos=(8, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        gbSizer.Add(self.txtRcpts, pos=(8, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)        

        gbSizer.Add(self.lblFile, pos=(9, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        gbSizer.Add(self.txtFile, pos=(9, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        gbSizer.Add(self.lblBody, pos=(10, 0), flag=wx.LEFT, border=10)
        gbSizer.Add(self.txtBody, pos=(11, 0), span=(txtBody_spanx, 4), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)

        gbSizer.Add(self.btnSubmit, pos=(btnSubmit_posx, 2), flag=wx.RIGHT, border=5)
        gbSizer.Add(self.btnClose, pos=(btnClose_posx, 3), flag=wx.RIGHT|wx.BOTTOM, border=25)

        gbSizer.AddGrowableCol(2)
        gbSizer.AddGrowableRow(GrowableRow)

        topSizer.Add(gbSizer, proportion=1, flag=wx.ALL|wx.EXPAND, border=5)
        
        self.panel.SetSizer(topSizer)
        topSizer.Fit(self)
        
    #
    # Methods
    #

    def getSecType(self):
        if bool(self.rdoBtn1.GetValue()):
            security = 'Plain'
        if bool(self.rdoBtn2.GetValue()):
            security = 'TLS'
        if bool(self.rdoBtn3.GetValue()):
           security = 'SSL'
        return security


    def checkConnection(self, HOST):
        self.security = self.getSecType()
        try:
            if self.security in ['Plain','TLS'] :
                return smtplib.SMTP(HOST,port=587,timeout=10)
            if self.security == 'SSL':
                return smtplib.SMTP_SSL(HOST,timeout=10)
        except:
            wx.MessageBox("Connection Error", "Warning", wx.OK | wx.ICON_WARNING)
            self.stBar.SetStatusText("Waiting for input...")


    def sentEmail(self, server, USER, LRCPTS, BODY='', FILE=''):
        try:
            if (FILE == ''):
                server.sendmail(USER, LRCPTS, BODY)
            else:
                server.sendmail(USER, LRCPTS, FILE)
        except smtplib.SMTPRecipientsRefused as e:
            result = e.recipients
            err_msg = []
            for value in result:
                err_msg.append(result[value][1])
            str = '\n'.join(err_msg)
            server.quit()
            wx.MessageBox(str,"Error sending email", wx.OK | wx.ICON_WARNING)
            self.stBar.SetStatusText("Ready to send again.")
        except:
            server.quit()
            wx.MessageBox('Unknown exception',"Error sending email", wx.OK | wx.ICON_WARNING)
            self.stBar.SetStatusText("Ready to send again.")
        else:
            server.quit()
            self.stBar.SetStatusText("Send email was success.")
            wx.MessageBox('Your email has been sent.', 'Sent Email', wx.OK)
            self.stBar.SetStatusText("Ready to send again.")


    #
    # Event handlers
    #

    def OnSelect(self, event):
        
        self.selected = self.cmbAccount.GetSelection()
        self.txtHost.SetValue(self.config.get(self.sections[self.selected],'host'))
        self.txtCtrlPass.SetValue(self.config.get(self.sections[self.selected],'password'))
        event.Skip()


    def onSubmit(self, event):

        USER = self.sections[self.selected]
        PASS = self.txtCtrlPass.GetValue()
        HOST = self.txtHost.GetValue()
        SUBJECT = self.txtSubject.GetValue()
        RCPTS = self.txtRcpts.GetValue()
        FILE = self.txtFile.GetValue()
        BODY = self.txtBody.GetValue()
        self.security = self.getSecType()
        messageID = str(int( (1 - random.random())*1E+16 )) + '$' + str(time.time()).split('.')[0] + '$' + USER

        # Check recipient
        if (SUBJECT==""):
            if (HOST=="") | (USER=="") | (PASS=="") | (RCPTS==""):
                wx.MessageBox("To check recipient all entries should be filled exclude Subject+Body.", "Warning", wx.OK | wx.ICON_WARNING)
            else:
                # code for check recipient
                LRCPTS = RCPTS.split(',')
                RCPT = LRCPTS[0]
                STATUS = '250'
                server = self.checkConnection(HOST)
                if (server):
                    try:
                        if self.security == 'TLS':
                            server.ehlo()
                            if server.has_extn('STARTTLS'):
                                server.starttls()
                                server.ehlo()
                        server.login(USER, PASS)
                    except:
                        wx.MessageBox("Authentication Failed", "Warning", wx.OK | wx.ICON_WARNING)
                        self.stBar.SetStatusText("Waiting for input...")
                        server.quit()
                    else:
                        server.mail(USER)
                        result = server.rcpt(RCPT)
                        if STATUS in str(result[0]):
                            wx.MessageBox('<%s>: Recipient address allowed'%RCPT, "Information", wx.OK)
                        else:
                            wx.MessageBox(result[1], "Warning", wx.OK | wx.ICON_WARNING)
                        server.quit()
        else:
            if (HOST=="") | (USER=="") | (PASS=="") | (RCPTS=="") | (BODY==""):
                wx.MessageBox("To sent email all entries should be filled.", "Warning", wx.OK | wx.ICON_WARNING)
            else:
                # code for sent email
                LRCPTS = RCPTS.split(',')            
                HEADER = "\r\n".join((
                    "From: %s <%s>" % (self.name, USER),
                    "To: %s" % RCPTS,
                    "Subject: %s" % SUBJECT ,
                    "Date: %s" % formatdate(localtime=True),
                    "Message-ID: <%s>" % messageID,
                    "X-Mailer: %s version %s" % (self.appname,self.version)))

                self.stBar.SetStatusText("Sending email...")
                # validate email address
                I = 0 # number of emails
                J = 0 # number of emails success to validate
                for email in LRCPTS:
                    I = I + 1
                    match = re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,6})$',email)
                    if match == None:
                        wx.MessageBox("Bad syntax for %s" % email, "Warning", wx.OK | wx.ICON_WARNING)
                        self.stBar.SetStatusText("Waiting for input...")
                    else:
                        J = J + 1

                if I == J:
                    server = self.checkConnection(HOST)
                    if (server):
                        try:
                            if self.security == 'TLS':
                                server.ehlo()
                                if server.has_extn('STARTTLS'):
                                    server.starttls()
                                    server.ehlo()
                            server.login(USER, PASS)
                        except:
                            server.quit()
                            wx.MessageBox("Authentication Failed", "Warning", wx.OK | wx.ICON_WARNING)
                            self.stBar.SetStatusText("Waiting for input...")
                        else:
                            if (FILE == ''):
                                BODY = HEADER + '\r\n\r\n' + BODY
                                self.sentEmail(server, USER, LRCPTS, BODY=BODY)
                            else:
                                header = 'Content-Disposition', 'attachment; filename="%s"' % path.basename(FILE)
                                if path.exists(FILE):
                                    msg = MIMEMultipart()
                                    msg["From"] = USER
                                    msg["Subject"] = SUBJECT
                                    msg["Date"] = formatdate(localtime=True)
                                    msg["Message-ID"] = '<' + messageID + '>'
                                    msg["X-Mailer"] = self.appname + ' version ' + self.version
                                    if BODY:
                                        msg.attach( MIMEText(BODY) )
                                    msg["To"] = RCPTS
                                    attachment = MIMEBase('application', "octet-stream")
                                    try:
                                        with open(FILE, "rb") as fh:
                                            data = fh.read()
                                        attachment.set_payload( data )
                                        encoders.encode_base64(attachment)
                                        attachment.add_header(*header)
                                        msg.attach(attachment)
                                    except IOError:
                                        wx.MessageBox("Error opening attachment file %s" % FILE, "Warning", wx.OK | wx.ICON_WARNING)
                                        self.stBar.SetStatusText("Waiting for input...")
                                    else:
                                        self.sentEmail(server, USER, LRCPTS, FILE=msg.as_string())
                                else:
                                    wx.MessageBox("Fle %s doesn't exists" % FILE, "Warning", wx.OK | wx.ICON_WARNING)
                                    self.stBar.SetStatusText("Waiting for input...")


    def onClose(self, event):
        self.Close()
        sys.exit(0)


    def togglePassword(self,event):
        self.txtPass.Show(self.password_shown)
        self.txtPassShow.Show(not self.password_shown)
        if not self.password_shown:
            self.txtPassShow.SetValue(self.txtPass.GetValue())
            self.txtPassShow.SetFocus()
            self.txtCtrlPass = self.txtPassShow
        else:
            self.txtPass.SetValue(self.txtPassShow.GetValue())
            self.txtPass.SetFocus()
            self.txtCtrlPass = self.txtPass
        self.txtPass.GetParent().Layout()
        self.password_shown= not self.password_shown


    def OnOpen(self, event):
        """Open a file attachment"""
        dlg = wx.FileDialog(self, "Choose a file", "", "", "*.*", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:        
            self.txtFile.SetValue(dlg.GetPath())


    def OnAbout(self, event):
        dialog = wx.MessageDialog(self,
            'A Simple Mailer in wxPython version ' + self.version + '\n\n'
            + 'Copyright ' + u'\xa9 ' + self.thisrelease \
            + ' by ' + self.identity + '\n' + self.email, \
            'About ' + self.appname, wx.OK)
        dialog.ShowModal()


    def OnExit(self,exit):
        self.Close()
        sys.exit(0)

class myApp(wx.App):
    def OnInit(self):
        self.locale = wx.Locale(wx.LANGUAGE_DEFAULT)
        myFrame(None, title='Simple Mailer', fconfig='config.ini')
        return True

if __name__ == '__main__':
    app = myApp(0)
    app.MainLoop()

