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
import sys
import random
import time
import myImages

OS = sys.platform

class myFrame(wx.Frame):

    def __init__(self, parent, title):
        wx.Frame.__init__(self, None, wx.ID_ANY, title=title)
        self.appname = 'Simple Mailer'
        self.version = '1.1'
        self.thisrelease = '2018'
        self.author = 'Arief Yudhawarman'
        self.email = 'awarmanf@yahoo.com'
        self.InitUI()
        self.Centre()
        self.Show()     

    def InitUI(self):

        # Setting up the menu
        menu = wx.Menu()

        itemOpen = wx.MenuItem(menu, wx.ID_ANY, "&Open\tCtrl+O","Open File Attachment")
        itemOpen.SetBitmap(myImages.open_item.GetBitmap())
        self.MenuAppend(OS,menu,itemOpen)

        itemAbout = wx.MenuItem(menu, wx.ID_ANY, "&About", "About this program")
        itemAbout.SetBitmap(myImages.about_item.GetBitmap())
        self.MenuAppend(OS,menu,itemAbout)

        menu.AppendSeparator()

        itemExit = wx.MenuItem(menu, wx.ID_ANY,"&Exit\tCtrl+Q","Exit program")
        itemExit.SetBitmap(myImages.exit_item.GetBitmap())
        self.MenuAppend(OS,menu,itemExit)

        menuBar = wx.MenuBar()
        menuBar.Append(menu,'File')
        self.SetMenuBar(menuBar)

        self.icon = myImages.email_icon.GetIcon()
        self.SetIcon(self.icon)

        txtSize = (200,-1)

        self.stBar = self.CreateStatusBar(1, 0)

        frame_1_statusbar_fields = ["Waiting for input..."]
        for i in range(len(frame_1_statusbar_fields)):
            self.stBar.SetStatusText(frame_1_statusbar_fields[i], i)

        panel = wx.Panel(self)
    
        topSizer = wx.BoxSizer(wx.VERTICAL)
        passSizer = wx.BoxSizer(wx.HORIZONTAL)
        btnSizer =  wx.BoxSizer(wx.HORIZONTAL)

        if OS == 'win32':
            gbsX = 12
            spanX = 1
            posX = 11
        else:
            gbsX = 15
            spanX = 4
            posX = 14

        gbSizer = wx.GridBagSizer(gbsX, 4)

        lblHeader  = wx.StaticText(panel, label="Simple Mailer")
        lblHeader.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD, 0, ""))
        gbSizer.Add(lblHeader, pos=(0, 0), span=(1,4),
            flag=wx.ALIGN_CENTER_HORIZONTAL|wx.TOP, border=10)

        line = wx.StaticLine(panel)
        gbSizer.Add(line, pos=(1, 0), span=(1,4),
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT,border=5)

        lblAccount = wx.StaticText(panel, label="Account Email")
        gbSizer.Add(lblAccount, pos=(2, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtAccount = wx.TextCtrl(panel,size=txtSize)
        gbSizer.Add(self.txtAccount, pos=(2, 1), span=(1, 3), 
            flag=wx.LEFT|wx.RIGHT, border=5)

        lblPass  = wx.StaticText(panel, label="Password")
        self.txtPass = wx.TextCtrl(panel,style=wx.TE_PASSWORD,size=txtSize)
        self.txtPassShow = wx.TextCtrl(panel,size=txtSize)
        self.txtPassShow.Hide()

        bmp = myImages.eye_icon.GetBitmap()
        self.btnShow = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmp)
        self.btnShow.SetToolTip(wx.ToolTip("Show/hide password"))

        gbSizer.Add(lblPass, pos=(3, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        passSizer.Add(self.txtPass)
        passSizer.Add(self.txtPassShow)
        passSizer.Add(self.btnShow,flag=wx.LEFT,border=5)
        gbSizer.Add(passSizer, pos=(3, 1), span=(1, 3), flag=wx.LEFT|wx.RIGHT, border=5)

        lblHost    = wx.StaticText(panel, label="SMTP Host")
        gbSizer.Add(lblHost, pos=(4, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtHost = wx.TextCtrl(panel,size=txtSize)
        gbSizer.Add(self.txtHost, pos=(4, 1), span=(1, 3), 
            flag=wx.LEFT|wx.RIGHT, border=5)

        lblSSL  = wx.StaticText(panel, label="Security Type")
        gbSizer.Add(lblSSL, pos=(5, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)

        self.rdoBtn1 = wx.RadioButton(panel, label='Plain',style=wx.RB_GROUP)
        self.rdoBtn2 = wx.RadioButton(panel, label='TLS')
        self.rdoBtn3 = wx.RadioButton(panel, label='SSL')
        gbSizer.Add(self.rdoBtn1, pos=(5,1),flag=wx.LEFT|wx.RIGHT,border=5)
        gbSizer.Add(self.rdoBtn2, pos=(5,2),flag=wx.LEFT|wx.RIGHT,border=5)
        gbSizer.Add(self.rdoBtn3, pos=(5,3),flag=wx.LEFT|wx.RIGHT,border=5)

        lblSubject = wx.StaticText(panel, label="Subject")
        gbSizer.Add(lblSubject, pos=(6, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtSubject = wx.TextCtrl(panel)
        gbSizer.Add(self.txtSubject, pos=(6, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        lblRcpts   = wx.StaticText(panel, label="Recipients")
        gbSizer.Add(lblRcpts, pos=(7, 0), flag=wx.LEFT|wx.ALIGN_CENTER_VERTICAL, border=10)
        self.txtRcpts = wx.TextCtrl(panel)
        self.txtRcpts.SetToolTip(wx.ToolTip("Add another recipient separated with comma"))
        gbSizer.Add(self.txtRcpts, pos=(7, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        lblFile = wx.StaticText(panel, label="File Attachment")
        gbSizer.Add(lblFile, pos=(8, 0), flag=wx.LEFT, border=10)

        self.txtFile = wx.TextCtrl(panel)
        gbSizer.Add(self.txtFile, pos=(8, 1), span=(1, 3), 
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        lblBody = wx.StaticText(panel, label="Body:")
        gbSizer.Add(lblBody, pos=(9, 0), flag=wx.LEFT, border=10)

        self.txtBody = wx.TextCtrl(panel,-1,"TIA\n\n---\n\nSent by %s %s"%(self.appname,self.version),style=wx.TE_MULTILINE)

        gbSizer.Add(self.txtBody, pos=(10, 0), span=(spanX, 4), flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)

        btnLogin = wx.Button(panel, label="Test Login", size=(90, 28))
        btnSubmit = wx.Button(panel, label="Submit", size=(70, 28))
        btnClose = wx.Button(panel, label="Close", size=(70, 28))

        btnSizer.Add(btnLogin,flag=wx.RIGHT,border=5)
        btnSizer.Add(btnSubmit,flag=wx.LEFT|wx.RIGHT,border=5)
        btnSizer.Add(btnClose,flag=wx.LEFT,border=5)

        gbSizer.Add(btnSizer, pos=(posX,1),span=(1,3),flag=wx.RIGHT, border=30)

        gbSizer.AddGrowableCol(2)
        gbSizer.AddGrowableRow(10)

        topSizer.Add(gbSizer, proportion=1, flag=wx.ALL|wx.EXPAND, border=5)
        
        panel.SetSizer(topSizer)
        topSizer.Fit(self)

        self.__set_properties()

        self.Bind(wx.EVT_MENU, self.OnOpen, itemOpen)
        self.Bind(wx.EVT_MENU, self.OnAbout, itemAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, itemExit)
        self.Bind(wx.EVT_BUTTON, self.togglePassword, self.btnShow)
        self.Bind(wx.EVT_BUTTON, self.onLogin, btnLogin)
        self.Bind(wx.EVT_BUTTON, self.onSubmit, btnSubmit)
        self.Bind(wx.EVT_BUTTON, self.onClose, btnClose)


    def __set_properties(self):
        self.password_shown = False
        self.security = 'Plain'
        self.txtCtrlPass = self.txtPass


    #
    # Methods
    #

    def MenuAppend(self,OS,menu,item):
        if OS == 'win32':
            menu.Append(item)
        else:
            menu.AppendItem(item)

    def getSecType(self):
        if bool(self.rdoBtn1.GetValue()):
            security = 'Plain'
        if bool(self.rdoBtn2.GetValue()):
            security = 'TLS'
        if bool(self.rdoBtn3.GetValue()):
           security = 'SSL'
        return security


    def checkConnection(self, HOST, port=None):
        self.security = self.getSecType()
        try:
            if self.security in ['Plain'] :
                return smtplib.SMTP(HOST,port=port,timeout=10)
            if self.security in ['TLS'] :
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

    def onLogin(self, event):
        HOST = self.txtHost.GetValue()
        USER = self.txtAccount.GetValue()
        PASS = self.txtCtrlPass.GetValue()
        self.security = self.getSecType()

        if (HOST == "") | (USER == "") | (PASS == "") :
            wx.MessageBox("Account email, password & host should be filled.", "Warning", wx.OK | wx.ICON_WARNING)
        else:
            server = self.checkConnection(HOST,port=587)
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
                else:
                    wx.MessageBox("Authentication Success", "Info", wx.OK | wx.ICON_INFORMATION)
                server.quit()


    def onSubmit(self, event):
        HOST = self.txtHost.GetValue()
        USER = self.txtAccount.GetValue()
        PASS = self.txtCtrlPass.GetValue()
        SUBJECT = self.txtSubject.GetValue()
        RCPTS = self.txtRcpts.GetValue()
        BODY = self.txtBody.GetValue()
        FILE = ''
        self.security = self.getSecType()   
        messageID = str(int( (1 - random.random())*1E+16 )) + '$' + str(time.time()).split('.')[0] + '$' + USER

        # Check recipient
        if (SUBJECT==""):
            if (HOST!="") & (USER!="") & (RCPTS!=""):
                LRCPTS = RCPTS.split(',')
                RCPT = LRCPTS[0]
                STATUS = '250'
                if (PASS != ""):
                    server = self.checkConnection(HOST,port=587)
                    if (server):
                        if self.security == 'TLS':
                            server.ehlo()
                            if server.has_extn('STARTTLS'):
                                server.starttls()
                                server.ehlo()
                        try:
                            server.login(USER, PASS)
                        except:
                            wx.MessageBox("Authentication Failed", "Warning", wx.OK | wx.ICON_WARNING)
                            self.stBar.SetStatusText("Waiting for input...")
                            server.quit()
                        else:
                            server.mail(USER)
                            result = server.rcpt(RCPT)
                            if STATUS in str(result[0]):
                                wx.MessageBox('<%s>: %s'% (RCPT,result[1]), "Information", wx.OK)
                            else:
                                wx.MessageBox(result[1], "Warning", wx.OK | wx.ICON_WARNING)
                            server.quit()
                else:
                    server = self.checkConnection(HOST,port=25)
                    server.ehlo()
                    server.mail(USER)
                    result = server.rcpt(RCPT)
                    if STATUS in str(result[0]):
                        wx.MessageBox('<%s>: %s'% (RCPT,result[1]), "Information", wx.OK)
                    else:
                        wx.MessageBox(result[1], "Warning", wx.OK | wx.ICON_WARNING)
                    server.quit()
            else:
                wx.MessageBox("To check recipient all entries should be filled exclude Subject+Body.", "Warning", wx.OK | wx.ICON_WARNING)
        else:
            if (HOST=="") | (USER=="") | (PASS=="") | (RCPTS=="") | (BODY==""):
                wx.MessageBox("To sent email all entries should be filled.", "Warning", wx.OK | wx.ICON_WARNING)
            else:
                LRCPTS = RCPTS.split(',')
                HEADER = "\r\n".join((
                    "From: %s" % USER,
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
                    server = self.checkConnection(HOST,port=587)
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
                            if (self.txtFile.GetValue() == ''):
                                BODY = HEADER + '\r\n\r\n' + BODY
                                self.sentEmail(server, USER, LRCPTS, BODY=BODY)
                            else:
                                FILE = self.txtFile.GetValue()
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

    def OnOpen(self, event):
        """Open a file attachment"""
        dlg = wx.FileDialog(self, "Choose a file", "", "", "*.*", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:        
            self.txtFile.SetValue(dlg.GetPath())
        dlg.Destroy()

    def OnAbout(self, event):
        dialog = wx.MessageDialog(self,
            'A Simple Mailer in wxPython version ' + self.version + '\n\n'
            + 'Copyright ' + u'\xa9 ' + self.thisrelease \
            + ' by ' + self.author + '\n' + self.email, \
            'About Simple Mailer' + self.appname, wx.OK)
        dialog.ShowModal()
        dialog.Destroy()

    def OnExit(self,exit):
        self.Close(True)

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


class myApp(wx.App):
    def OnInit(self):
        self.locale = wx.Locale(wx.LANGUAGE_DEFAULT)
        self.frame = myFrame(None, title='Simple Mailer')
        return True

if __name__ == '__main__':
    app = myApp(0)
    app.MainLoop()


