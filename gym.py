import wx
import wx.lib.scrolledpanel
import sqlite3
import os
import shutil
from xlwt import Workbook
import xlrd
from win32com.client import constants, Dispatch
import wx.grid
from dateutil.relativedelta import relativedelta
from datetime import datetime
import urllib2
import cookielib
from getpass import getpass
import sys
#from stat import *
import  cStringIO
import urllib2
from datetime import date


class MainFrame(wx.Frame):
    def __init__(self,parent,id,title):
        screenSize = wx.DisplaySize()
        screenWidth = screenSize[0]
        screenHeight = screenSize[1]
        wx.Frame.__init__(self,parent,id,title,size=screenSize)

        panel1=wx.Panel(self,-1,style=wx.SIMPLE_BORDER,size=(screenWidth,screenHeight/9))    
        panel2 = wx.lib.scrolledpanel.ScrolledPanel(self,-1, size=(screenWidth,screenHeight-screenHeight/9), pos=(0,screenHeight/9), style=wx.SIMPLE_BORDER)   
        panel2.SetupScrolling()

        nb=wx.Notebook(panel2)

        page1=FirstPage(nb) 
        page2=SecondPage(nb)              
        page3=ThirdPage(nb)
        page4=FourthPage(nb)
        page5=FifthPage(nb)
        page6=SixthPage(nb)             
        page7=SeventhPage(nb)


        nb.AddPage(page1,"Form Filling")
        nb.AddPage(page2,"Search Member")
        nb.AddPage(page3,"Renew Membership")
        nb.AddPage(page4,"Due Date Info")
        nb.AddPage(page5,"upload image seperately")
        nb.AddPage(page6,"Backup/Restore")
        nb.AddPage(page7,"Delete Entry")
        
        nb.SetPadding( padding=(50,3))

        heading=wx.StaticText(panel1,-1,label='GYM Software',pos=(screenWidth/3,20))
        font=wx.Font(30,wx.DECORATIVE, wx.ITALIC, wx.NORMAL)
        heading.SetFont(font)

        sizer=wx.BoxSizer()
        sizer.Add(nb,10,wx.EXPAND)
        panel2.SetSizer(sizer)

        panel1.SetBackgroundColour("#3299CC")

        self.Show(True)
        
class FirstPage(wx.Panel): 
    def __init__(self, parent): 
        wx.Panel.__init__(self,parent,size=wx.DefaultSize)
        screenSize = wx.DisplaySize()
        sw = screenSize[0]
        sh = screenSize[1]

        panel=wx.lib.scrolledpanel.ScrolledPanel(self,size=(sw-sw/2,sh-sh/3+50),style=wx.SIMPLE_BORDER,pos=(50,10))
        #panel=wx.lib.scrolledpanel.ScrolledPanel(self)
        panel.SetupScrolling()

        sizer = wx.GridBagSizer(10, 10)

        text0 = wx.StaticText(panel, label="Enrollment/Serial Number")
        sizer.Add(text0, pos=(0, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc0 = wx.TextCtrl(panel)
        sizer.Add(self.tc0, pos=(0, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)
        
        text1 = wx.StaticText(panel, label="Name")
        sizer.Add(text1, pos=(1, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc1 = wx.TextCtrl(panel)
        sizer.Add(self.tc1, pos=(1, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text2 = wx.StaticText(panel, label="Father's/Husband's Name")
        sizer.Add(text2, pos=(2, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc2 = wx.TextCtrl(panel)
        sizer.Add(self.tc2, pos=(2, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text3 = wx.StaticText(panel, label="Date of Birth")
        sizer.Add(text3, pos=(3, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc3 = wx.TextCtrl(panel)
        sizer.Add(self.tc3, pos=(3, 1),span=(1,10), flag=wx.TOP|wx.EXPAND, border=5)

        text31 = wx.StaticText(panel, label="Date Format : dd/mm/yyyy")
        sizer.Add(text31, pos=(3, 11), flag=wx.TOP|wx.LEFT, border=10)
        
        text4 = wx.StaticText(panel, label="Home Address")
        sizer.Add(text4, pos=(4, 0) ,flag=wx.TOP|wx.LEFT, border=10)

        self.tc4 = wx.TextCtrl(panel)
        sizer.Add(self.tc4, pos=(4, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text5 = wx.StaticText(panel, label="Office Address")
        sizer.Add(text5, pos=(5, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc5 = wx.TextCtrl(panel)
        sizer.Add(self.tc5, pos=(5, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text6 = wx.StaticText(panel, label="Gender")
        sizer.Add(text6, pos=(6, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc6 = wx.TextCtrl(panel)
        sizer.Add(self.tc6, pos=(6, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text7 = wx.StaticText(panel, label="Contact No.")
        sizer.Add(text7, pos=(7, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc7 = wx.TextCtrl(panel)
        sizer.Add(self.tc7, pos=(7, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)


        text8 = wx.StaticText(panel, label="Residence No.")
        sizer.Add(text8, pos=(8, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc8 = wx.TextCtrl(panel)
        sizer.Add(self.tc8, pos=(8, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text9 = wx.StaticText(panel, label="Occupation")
        sizer.Add(text9, pos=(9, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc9 = wx.TextCtrl(panel)
        sizer.Add(self.tc9, pos=(9, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text10 = wx.StaticText(panel, label="Date of Joining")
        sizer.Add(text10, pos=(10, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc10 = wx.TextCtrl(panel)
        sizer.Add(self.tc10, pos=(10, 1),span=(1,10), flag=wx.TOP|wx.EXPAND, border=5)

        text10 = wx.StaticText(panel, label="Date Format : dd/mm/yyyy")
        sizer.Add(text10, pos=(10, 11), flag=wx.TOP|wx.LEFT, border=10)

        text11 = wx.StaticText(panel, label="Fee Paid")
        sizer.Add(text11, pos=(11, 0), flag=wx.TOP|wx.LEFT, border=10)

        self.tc11 = wx.TextCtrl(panel)
        sizer.Add(self.tc11, pos=(11, 1),span=(1,15), flag=wx.TOP|wx.EXPAND, border=5)

        text12 = wx.StaticText(panel, label="Membership Time")
        sizer.Add(text12, pos=(12, 0), flag=wx.TOP|wx.LEFT, border=10)

        time = ['1 month', '2 month', '3 month', '4 month','5 month','6 month','7 month','8 month','9 month','10 month','11 month','1 year','2 year','3 year']
        self.combo = wx.ComboBox(panel, choices = time,pos=(12, 1))
        sizer.Add(self.combo, pos=(12, 1), flag=wx.TOP|wx.EXPAND, border=5)

        button=wx.Button(panel,label="Save",size=(100,50))
        button.Bind(wx.EVT_BUTTON,self.Save)

        sizer.Add(button, pos=(13,1), flag=wx.TOP|wx.EXPAND, border=5)

        imgtext = wx.StaticText(self, label="Upload Image",pos=(sw-sw/2+200,100))
        self.tcimg = wx.TextCtrl(self,pos=(sw-sw/2+200,150))
        bimg=wx.Button(self,label="Select Image",size=(100,30),pos=(sw-sw/2+200,200))
        bimg.Bind(wx.EVT_BUTTON,self.Select)
        imginfotext = wx.StaticText(self, label="Only select a .jpg type image",pos=(sw-sw/2+200,250))
        
        panel.SetSizer(sizer)

    def Select(self,event):
        
        dlg=wx.FileDialog(self,"Choose a .png image file","","","*.jpg")
        if dlg.ShowModal()==wx.ID_OK:
            path=dlg.GetPath()
        path=dlg.GetPath()
        folder,filename=os.path.split(path)
        dlg.Destroy()
        self.tcimg.SetValue(filename)

        self.path=path
      

    def Save(self,event):

        self.EnrolmentNo=self.tc0.GetValue()  
        self.Name=self.tc1.GetValue()
        self.ForHname=self.tc2.GetValue()
        self.Dob=self.tc3.GetValue()
        self.HAddress=self.tc4.GetValue()
        self.OAddress=self.tc5.GetValue()
        self.Gender=self.tc6.GetValue()
        self.ContactNo=self.tc7.GetValue()
        self.ResidenceNo=self.tc8.GetValue()
        self.Occupation=self.tc9.GetValue()
        self.Doj=self.tc10.GetValue()           
        self.FeePaid=self.tc11.GetValue()
        self.MembershipTime=self.combo.GetValue()   

        
        self.ImageName=self.tcimg.GetValue()

        if len(self.EnrolmentNo)==0:
            msgbox = wx.MessageBox('WARNING : Please fill Enrollment Number field', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        elif len(self.Doj)==0:
            msgbox = wx.MessageBox('WARNING : Please fill Date of Joining field', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        elif len(self.MembershipTime)==0:
            msgbox = wx.MessageBox('WARNING : Please fill Membership Time field', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        else:
        
            conn=sqlite3.connect("MemberData.db")
            c=conn.cursor()
            c.execute('Select * from gym1 where EnrolmentNo="%s"'%(self.EnrolmentNo))
            data=c.fetchall()
            conn.close()


            if len(data)==0:                                                       #If this is new enrolment no. entry
                conn=sqlite3.connect("MemberData.db")
                c=conn.cursor()
                c.execute("INSERT INTO gym1 VALUES (?, ?, ?, ?,?, ?, ?, ?,?,?,?,?,?)",(self.EnrolmentNo,self.Name,self.ForHname,self.Dob,\
                            self.HAddress,self.OAddress,self.Gender,self.ContactNo,self.ResidenceNo,self.Occupation,self.Doj,self.FeePaid,self.MembershipTime))

                conn.commit()
                conn.close()

                if len(self.ImageName)==0:
                    msgbox = wx.MessageBox('NO IMAGE UPLOADED', 
                               'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
                else:
                    conn=sqlite3.connect("MemberData.db")
                    c=conn.cursor()
                    c.execute("INSERT INTO gym2 VALUES (?, ?)",(self.EnrolmentNo,self.ImageName))
                    conn.commit()
                    conn.close()

                    shutil.copy2(self.path, 'C:\Users\Nishant\Desktop\gym software')


                msgbox = wx.MessageBox('SAVE SUCCESSFULL!', 
                               'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

                

            else:                                                               # If this enrolment no. entry already exists
                msgbox = wx.MessageBox('WARNING : This Enrollment Number already exits', 
                           'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)


        
Eno=0

class SecondPage(wx.Panel): 
    def __init__(self, parent): 
        wx.Panel.__init__(self,parent,size=wx.DefaultSize)
        
        text1 = wx.StaticText(self, label="Search for a Member",pos=(250,50))
        font=wx.Font(20,wx.DECORATIVE, wx.ITALIC, wx.NORMAL)
        text1.SetFont(font)

        
        text2 = wx.StaticText(self, label="Enter Enrolment No.",pos=(100,200))
        self.tc1 = wx.TextCtrl(self,pos=(100,250))
        
        button=wx.Button(self,label="Search",size=(100,30),pos=(350,250))
        button.Bind(wx.EVT_BUTTON,self.Search)

    def Search(self,event):

        global Eno
        Eno=self.tc1.GetValue()

        conn=sqlite3.connect("MemberData.db")
        c=conn.cursor()
        c.execute('Select * from gym1 where EnrolmentNo="%s"'%(Eno))
        data=c.fetchall()
        conn.close()

        if len(data)==0:
            msgbox = wx.MessageBox('This Enrollment Number does not exits', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        else:
            app1=MyApp1()
            app1.MainLoop()
            
            
class MyApp1(wx.App):
    def OnInit(self):
        frame=ShowDetails(None,-1,'Gym Software')
        frame.Show()
        frame.Maximize(True)
        return True

class ShowDetails(wx.Frame):
    def __init__(self,parent,id,title):
        screenSize = wx.DisplaySize()
        screenWidth = screenSize[0]
        screenHeight = screenSize[1]
        wx.Frame.__init__(self,parent,id,title,size=screenSize)

        panel=wx.Panel(self,size=(300,screenHeight),pos=(0,0),style=wx.SIMPLE_BORDER)
        panel1 = wx.lib.scrolledpanel.ScrolledPanel(self,-1, size=(screenWidth,screenHeight+400), pos=(300,0), style=wx.SIMPLE_BORDER)  
        panel1.SetupScrolling()
        
        global Eno
        
        conn=sqlite3.connect("MemberData.db")
        c=conn.cursor()
        c.execute('Select * from gym1 where EnrolmentNo="%s"'%(Eno))
        data=c.fetchall()
        conn.close()
        
        sizer = wx.GridBagSizer(10, 10)

        text0 = wx.StaticText(panel, label="Enrolment/Serial Number")
        sizer.Add(text0, pos=(0, 0), flag=wx.TOP|wx.LEFT, border=10)

        t0 = wx.StaticText(panel, label=data[0][0])
        sizer.Add(t0, pos=(0, 1), flag=wx.TOP|wx.LEFT, border=10)
        
        text1 = wx.StaticText(panel, label="Name")
        sizer.Add(text1, pos=(1, 0), flag=wx.TOP|wx.LEFT, border=10)

        t1 = wx.StaticText(panel, label=data[0][1])
        sizer.Add(t1, pos=(1, 1), flag=wx.TOP|wx.LEFT, border=10)

        text2 = wx.StaticText(panel, label="Father's/Husband's Name")
        sizer.Add(text2, pos=(2, 0), flag=wx.TOP|wx.LEFT, border=10)

        t2 = wx.StaticText(panel, label=data[0][2])
        sizer.Add(t2, pos=(2, 1), flag=wx.TOP|wx.LEFT, border=10)

        text3 = wx.StaticText(panel, label="Date of Birth")
        sizer.Add(text3, pos=(3, 0), flag=wx.TOP|wx.LEFT, border=10)

        t3 = wx.StaticText(panel, label=data[0][3])
        sizer.Add(t3, pos=(3, 1), flag=wx.TOP|wx.LEFT, border=10)


        text4 = wx.StaticText(panel, label="Home Address")
        sizer.Add(text4, pos=(4, 0) ,flag=wx.TOP|wx.LEFT, border=10)

        t4 = wx.StaticText(panel, label=data[0][4])
        sizer.Add(t4, pos=(4, 1) ,flag=wx.TOP|wx.LEFT, border=10)

        text5 = wx.StaticText(panel, label="Office Address")
        sizer.Add(text5, pos=(5, 0), flag=wx.TOP|wx.LEFT, border=10)

        t5 = wx.StaticText(panel, label=data[0][5])
        sizer.Add(t5, pos=(5, 1), flag=wx.TOP|wx.LEFT, border=10)

        text6 = wx.StaticText(panel, label="Gender")
        sizer.Add(text6, pos=(6, 0), flag=wx.TOP|wx.LEFT, border=10)

        t6 = wx.StaticText(panel, label=data[0][6])
        sizer.Add(t6, pos=(6, 1), flag=wx.TOP|wx.LEFT, border=10)

        text7 = wx.StaticText(panel, label="Contact No.")
        sizer.Add(text7, pos=(7, 0), flag=wx.TOP|wx.LEFT, border=10)

        t7 = wx.StaticText(panel, label=data[0][7])
        sizer.Add(t7, pos=(7, 1), flag=wx.TOP|wx.LEFT, border=10)


        text8 = wx.StaticText(panel, label="Residence No.")
        sizer.Add(text8, pos=(8, 0), flag=wx.TOP|wx.LEFT, border=10)

        t8 = wx.StaticText(panel, label=data[0][8])
        sizer.Add(t8, pos=(8, 1), flag=wx.TOP|wx.LEFT, border=10)

        text9 = wx.StaticText(panel, label="Occupation")
        sizer.Add(text9, pos=(9, 0), flag=wx.TOP|wx.LEFT, border=10)

        t9 = wx.StaticText(panel, label=data[0][9])
        sizer.Add(t9, pos=(9, 1), flag=wx.TOP|wx.LEFT, border=10)

        text10 = wx.StaticText(panel, label="Date of Joining")
        sizer.Add(text10, pos=(10, 0), flag=wx.TOP|wx.LEFT, border=10)

        t10 = wx.StaticText(panel, label=data[0][10])
        sizer.Add(t10, pos=(10, 1), flag=wx.TOP|wx.LEFT, border=10)
 
        text11 = wx.StaticText(panel, label="Fee Paid")
        sizer.Add(text11, pos=(11, 0), flag=wx.TOP|wx.LEFT, border=10)

        t11 = wx.StaticText(panel, label=data[0][11])
        sizer.Add(t11, pos=(11, 1), flag=wx.TOP|wx.LEFT, border=10)

        text12 = wx.StaticText(panel, label="Membership Time")
        sizer.Add(text12, pos=(12, 0), flag=wx.TOP|wx.LEFT, border=10)

        t12 = wx.StaticText(panel, label=data[0][12])
        sizer.Add(t12, pos=(12, 1), flag=wx.TOP|wx.LEFT, border=10)

        
        
        conn=sqlite3.connect("MemberData.db")
        c=conn.cursor()
        c.execute('Select Image from gym2 where EnrolmentNo="%s"'%(Eno))
        imagename=c.fetchall()
        conn.close()

        if len(imagename)==0:
            msgbox = wx.MessageBox('No image uploaded for this member!', 
                           'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        else:

            path='C:/MemberImages'
            
           
            
            try:
                # pick a .jpg file you have in the working folder
                imageFile = path+"/"+imagename[0][0]
                data = open(imageFile, "rb").read()
                # convert to a data stream
                stream = cStringIO.StringIO(data)
                # convert to a bitmap
                bmp = wx.BitmapFromImage( wx.ImageFromStream( stream ))
                # show the bitmap, (5, 5) are upper left corner coordinates
                wx.StaticBitmap(panel1, -1, bmp, (10, 10))
                
                # alternate (simpler) way to load and display a jpg image from a file
                # actually you can load .jpg  .png  .bmp  or .gif files
                jpg1 = wx.Image(imageFile, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
                # bitmap upper left corner is in the position tuple (x, y) = (5, 5)
                wx.StaticBitmap(panel1, -1, jpg1, (5,5), (jpg1.GetWidth(), jpg1.GetHeight()))
                #wx.StaticBitmap(panel1, -1, jpg1, (5 , 5), size=(screenWidth-300,screenHeight))
            except IOError:
                msgbox = wx.MessageBox('Error Loading Image!', 
                           'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

            
            
        panel.SetSizer(sizer)

        sizerpanel=wx.BoxSizer(wx.HORIZONTAL)
        sizerpanel.Add(panel)
        sizerpanel.Add(panel1)
        self.SetSizer(sizerpanel)

        
        
                         
class ThirdPage(wx.Panel): 
    def __init__(self, parent): 
        wx.Panel.__init__(self,parent,size=wx.DefaultSize)

        text1 = wx.StaticText(self, label="Enter Enrollment No.",pos=(100,100))
        self.tc1 = wx.TextCtrl(self,pos=(250,100))

        text2 = wx.StaticText(self, label="Enter New Fee Paid",pos=(100,150))
        self.tc2 = wx.TextCtrl(self,pos=(250,150))
        
        text3 = wx.StaticText(self, label="Enter date of Fee Paid",pos=(100,200))
        self.tc3 = wx.TextCtrl(self,pos=(250,200))

        text35 = wx.StaticText(self, label="Date Format : dd/mm/yyyy",pos=(400,200))
        
        text4 = wx.StaticText(self, label="Membership Time",pos=(100,250))

        time = ['1 month', '2 month', '3 month', '4 month','5 month','6 month','7 month','8 month','9 month','10 month','11 month','1 year','2 year','3 year'] 
        self.combo = wx.ComboBox(self, choices = time,pos=(250,250))
        

        button=wx.Button(self,label="Update Memebership",size=(150,30),pos=(250,350))
        button.Bind(wx.EVT_BUTTON,self.UpdateMembership)

    def UpdateMembership(self,event):

        tc1=self.tc1.GetValue()
        tc2=self.tc2.GetValue()
        tc3=self.tc3.GetValue()
        tc4=self.combo.GetValue()

        if len(tc1)==0 or len(tc2)==0 or len(tc3)==0 or len(tc4)==0:
            msgbox = wx.MessageBox('WARNING : Please fill all the fields', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        else:

            conn=sqlite3.connect("MemberData.db")
            c=conn.cursor()
            c.execute('Select * from gym1 where EnrolmentNo="%s"'%(str(tc1)))
            data=c.fetchall()
            conn.close()


            if len(data)==0:
                msgbox = wx.MessageBox('WARNING : This enrollment number does not exist.', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
            else:
                
                conn=sqlite3.connect('MemberData.db')
                c=conn.cursor()
                c.execute('SELECT FeePaid FROM gym1')
                c.execute('UPDATE gym1 SET FeePaid="%s" WHERE EnrolmentNo="%s" '%(str(tc2),str(tc1)))
                conn.commit()
                c.execute('UPDATE gym1 SET Doj="%s" WHERE EnrolmentNo="%s" '%(str(tc3),str(tc1)))
                conn.commit()
                c.execute('UPDATE gym1 SET MembershipTime="%s" WHERE EnrolmentNo="%s" '%(str(tc4),str(tc1)))
                conn.commit()
                conn.close()

                conn=sqlite3.connect('MemberData.db')     #delete entry from duedatedatabase
                c=conn.cursor()
                c.execute('DELETE FROM gym3 WHERE EnrolmentNo="%s"'%(str(tc1)))
                conn.commit()
                conn.close()

                msgbox = wx.MessageBox('CHANGES SAVED', 
                               'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
            

class FourthPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self,parent,size=wx.DefaultSize)
        
        pnl2 = wx.lib.scrolledpanel.ScrolledPanel(self,3, size=(1200,500), pos=(10,10), style=wx.SIMPLE_BORDER)
        pnl2.SetupScrolling()

        
        self.DueDateDatabase()
        self.sendsms()
        self.birthdaysms()
        
        conn=sqlite3.connect('MemberData.db')
        c=conn.cursor()
        c.execute('SELECT * FROM gym3')
        data=c.fetchall()
        conn.close()

        count=0
        for i in data:
            count+=1
            
        grid = wx.grid.Grid(pnl2, -1)
        grid.CreateGrid(count, 8)

        grid.SetColSize(0, 120)
        grid.SetColSize(1, 200)
        grid.SetColSize(2, 150)
        grid.SetColSize(3, 150)
        grid.SetColSize(4, 120)
        grid.SetColSize(5, 120)
        grid.SetColSize(6, 150)
        
        grid.SetColLabelValue(0, "Enrollment No.")
        grid.SetColLabelValue(1, "Name")
        grid.SetColLabelValue(2, "Contact No.")
        grid.SetColLabelValue(3, "Residence No.")
        grid.SetColLabelValue(4, "Date of Joining")
        grid.SetColLabelValue(5, "Fee Paid")
        grid.SetColLabelValue(6, "Membership Time")
        grid.SetColLabelValue(7, "Status")
        
        

        z=-1
        for i in data:
            z+=1
            for j in range(0,8):
                a=str(i[j])
                grid.SetCellValue(z,j,a)
                grid.SetReadOnly(z, j)

        
        

        gbox = wx.BoxSizer()
        gbox.Add(grid, 1, wx.EXPAND)
        pnl2.SetSizer(gbox)

    def DueDateDatabase(self):

        conn=sqlite3.connect('MemberData.db')
        c=conn.cursor()
        c.execute('SELECT EnrolmentNo,Name,ContactNo,ResidenceNo,Doj,FeePaid,MembershipTime FROM gym1')
        data=c.fetchall()
        conn.close()

        for i in data:
            doj=i[4]
            MT=i[6]        #MembershipTime

            if MT[2]=='m' or MT[3]=='m':

                m=int(MT[:2])

                d = doj[:2] + "/" + doj[3:5] + "/" + doj[6:]

                date = datetime.strptime(d, "%d/%m/%Y")            #convert to standard datetime
                date=date+relativedelta(months=m)                  #add date
                #date=date.strftime('%d/%m/%Y')                     #converts to normal date but it isn't required here

                
            elif MT[2]=='y':

                y=int(MT[:2])

                d = doj[:2] + "/" + doj[3:5] + "/" + doj[6:]

                date = datetime.strptime(d, "%d/%m/%Y")            #convert to standard datetime
                date=date+relativedelta(years=y)                   #add date
                #date=date.strftime('%d/%m/%Y')                     #converts to normal date but it isn't required here

            present=datetime.now()
            if present>date:
                
                status='Not Send'
                a=list(i)
                a.append(str(status))
                conn=sqlite3.connect('MemberData.db')
                c=conn.cursor()
                c.execute('Select EnrolmentNo from gym3')
                enos=c.fetchall()
                flag=1
                for j in enos:
                    if j[0]==i[0]:
                        flag=0

                if flag==1:
                    c.execute('INSERT INTO gym3 VALUES (?,?,?,?,?,?,?,?)', a)
                    conn.commit()
                conn.close()
        
    def sendsms(self):

        if self.internet_on()==True:
            conn=sqlite3.connect('MemberData.db')
            c=conn.cursor()
            c.execute('Select * from gym3 where Status="Not Send"')
            data=c.fetchall()
            conn.close()

            #print data

            #------------------------------sms for owner

            for i in data:
                try:
                    
                    message = 'GYM Membership Renewal Alert: Membership for Enrollment Number "%s" is finished. Name "%s" Contact no. "%s"'%(i[0],i[1],i[2])
                    number = "98XXXXXXXX"

                    if __name__ == "__main__":    
                        username = "98XXXXXXXX"
                        passwd = "XXXXXX"

                    #logging into the sms site
                    url ='http://site24.way2sms.com/Login1.action?'
                    data = 'username='+username+'&password='+passwd+'&Submit=Sign+in'

                    #For cookies

                    cj= cookielib.CookieJar()
                    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))

                    #Adding header details
                    opener.addheaders=[('User-Agent','Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2062.120')]
                    try:
                        usock =opener.open(url, data)
                    except IOError:
                        print "error"

                    jession_id =str(cj).split('~')[1].split(' ')[0]
                    send_sms_url = 'http://site24.way2sms.com/smstoss.action?'
                    send_sms_data = 'ssaction=ss&Token='+jession_id+'&mobile='+number+'&message='+message+'&msgLen=136'
                    opener.addheaders=[('Referer', 'http://site25.way2sms.com/sendSMS?Token='+jession_id)]

                    try:
                        sms_sent_page = opener.open(send_sms_url,send_sms_data)
                        
                    except IOError:
                        print "error"


                    #-----------------------------------------sms for member

                    doj=i[4]
                    MT=i[6]        #MembershipTime

                    if MT[2]=='m' or MT[3]=='m':

                        m=int(MT[:2])

                        d = doj[:2] + "/" + doj[3:5] + "/" + doj[6:]


                        date = datetime.strptime(d, "%d/%m/%Y")            #convert to standard datetime
                        date=date+relativedelta(months=m)                  #add date
                        #date=date.strftime('%d/%m/%Y')                     #converts to normal date but it isn't required here

                        
                    elif MT[2]=='y':

                        y=int(MT[:2])

                        d = doj[:2] + "/" + doj[3:5] + "/" + doj[6:]

                        date = datetime.strptime(d, "%d/%m/%Y")            #convert to standard datetime
                        date=date+relativedelta(years=y)                   #add date
                        #date=date.strftime('%d/%m/%Y')                     #converts to normal date but it isn't required here


                    date=date.strftime('%d/%m/%Y')                     #convert to normal date
                    
                    message = 'Dear valued customer. Your membership payment at GYM has expired on "%s". Please ignore if already paid "stay healthy with us".'%(date)
                    number = i[2]

                    if __name__ == "__main__":    
                        username = "98XXXXXXXX"
                        passwd = "XXXXXX"

                    #logging into the sms site
                    url ='http://site24.way2sms.com/Login1.action?'
                    data = 'username='+username+'&password='+passwd+'&Submit=Sign+in'

                    #For cookies

                    cj= cookielib.CookieJar()
                    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))

                    #Adding header details
                    opener.addheaders=[('User-Agent','Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2062.120')]
                    try:
                        usock =opener.open(url, data)
                    except IOError:
                        print "error"

                    jession_id =str(cj).split('~')[1].split(' ')[0]
                    send_sms_url = 'http://site24.way2sms.com/smstoss.action?'
                    send_sms_data = 'ssaction=ss&Token='+jession_id+'&mobile='+number+'&message='+message+'&msgLen=136'
                    opener.addheaders=[('Referer', 'http://site25.way2sms.com/sendSMS?Token='+jession_id)]

                    try:
                        sms_sent_page = opener.open(send_sms_url,send_sms_data)
                        
                    except IOError:
                        print "error"

                    
                    conn=sqlite3.connect('MemberData.db')
                    c=conn.cursor()
                    #c.execute('Select from gym3 where Status="Not Send"')
                    c.execute('UPDATE gym3 SET Status="Sent" where EnrolmentNo="%s"'%(i[0]))
                    conn.commit()
                    c.execute('Select * from gym3')
                    data=c.fetchall()
                    #print data
                    conn.close()

                except:
                    msgbox = wx.MessageBox('error sending message!', 
                                   'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP) 
        else:
            pass
        
    def internet_on(self):
        try:
            urllib2.urlopen('http://216.58.192.142', timeout=1)
            return True
        except urllib2.URLError as err: 
            return False

    def birthdaysms(self):
        from datetime import date
        d=date.today()
        d=str(d)
        dm=d[5:7]#month
        dd=d[8:10]#date

        conn=sqlite3.connect('MemberData.db')
        c=conn.cursor()
        c.execute('Select * from gym1')
        data=c.fetchall()
        conn.close()

        for i in data:
            date=i[3][0:2]
            month=i[3][3:5]
            if dm==month and dd==date:
                try:
                    message = 'GYM Member birthday: Today is birthday of "%s" ,Contact no. "%s" ,Enrollment Number "%s"'%(i[1],i[2],i[0])
                    number = "98XXXXXXXX"

                    if __name__ == "__main__":    
                        username = "98XXXXXXXX"
                        passwd = "XXXXXX"

                    #logging into the sms site
                    url ='http://site24.way2sms.com/Login1.action?'
                    data = 'username='+username+'&password='+passwd+'&Submit=Sign+in'

                    #For cookies

                    cj= cookielib.CookieJar()
                    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))

                    #Adding header details
                    opener.addheaders=[('User-Agent','Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2062.120')]
                    try:
                        usock =opener.open(url, data)
                    except IOError:
                        print "error"

                    jession_id =str(cj).split('~')[1].split(' ')[0]
                    send_sms_url = 'http://site24.way2sms.com/smstoss.action?'
                    send_sms_data = 'ssaction=ss&Token='+jession_id+'&mobile='+number+'&message='+message+'&msgLen=136'
                    opener.addheaders=[('Referer', 'http://site25.way2sms.com/sendSMS?Token='+jession_id)]

                    try:
                        sms_sent_page = opener.open(send_sms_url,send_sms_data)
                        
                    except IOError:
                        print "error"

                except:
                    pass

    

class FifthPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self,parent,size=wx.DefaultSize)

        text0 = wx.StaticText(self, label="Enrollment/Serial Number", pos=(100,100))
        self.tc0 = wx.TextCtrl(self,pos=(300,100))
        
        imgtext = wx.StaticText(self, label="Upload Image",pos=(100,200))
        self.tcimg = wx.TextCtrl(self,pos=(100,250))
        bimg=wx.Button(self,label="Select Image",size=(100,30),pos=(100,300))
        bimg.Bind(wx.EVT_BUTTON,self.Select)
        imginfotext = wx.StaticText(self, label="Only select a .jpg type image",pos=(100,350))

        save=wx.Button(self,label="Save Image",size=(100,30),pos=(100,400))
        save.Bind(wx.EVT_BUTTON,self.SaveImage)

    def Select(self,event):
        
        dlg=wx.FileDialog(self,"Choose a .png image file","","","*.jpg")
        if dlg.ShowModal()==wx.ID_OK:
            path=dlg.GetPath()
        path=dlg.GetPath()
        folder,filename=os.path.split(path)
        dlg.Destroy()
        self.tcimg.SetValue(filename)

        self.path=path

    def SaveImage(self,event):

        self.EnrolmentNo=self.tc0.GetValue()
        self.ImageName=self.tcimg.GetValue()
        
        conn=sqlite3.connect("MemberData.db")
        c=conn.cursor()
        c.execute('Select * from gym2 where EnrolmentNo="%s"'%(self.EnrolmentNo))
        data=c.fetchall()
        conn.close()

        conn=sqlite3.connect("MemberData.db")
        c=conn.cursor()
        c.execute('Select * from gym1 where EnrolmentNo="%s"'%(self.EnrolmentNo))
        data1=c.fetchall()
        conn.close()

        if len(data1)==0:
            msgbox = wx.MessageBox('WARNING : This Enrollment Number does not exits in main database', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        else:
            if len(data)==0:                                                       #If this is new enrolment no. entry
                conn=sqlite3.connect("MemberData.db")
                c=conn.cursor()
                c.execute("INSERT INTO gym2 VALUES (?, ?)",(self.EnrolmentNo,self.ImageName))
                conn.commit()
                conn.close()

                shutil.copy2(self.path, 'C:\Users\Nishant\Desktop\gym software')


                msgbox = wx.MessageBox('SAVE SUCCESSFULL!', 
                               'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

            else:                                                               # If this enrolment no. entry already exists
                msgbox = wx.MessageBox('WARNING : This Enrollment Number already exits', 
                           'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

fname=''
foldername=''
        
class SixthPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self,parent,size=wx.DefaultSize)

        t1 = wx.StaticText(self, label="backup all data to a excel file", pos=(100,100))
        b1=wx.Button(self,label="backup",size=(100,30),pos=(300,100))
        b1.Bind(wx.EVT_BUTTON,self.backup)

        t2 = wx.StaticText(self, label="import backup file1", pos=(100,200))
        self.tc1 = wx.TextCtrl(self,pos=(220,200))
        b2=wx.Button(self,label="choose file1",size=(100,30),pos=(350,200))
        b2.Bind(wx.EVT_BUTTON,self.choosefile1)
        b3=wx.Button(self,label="import file1",size=(100,30),pos=(350,300))
        b3.Bind(wx.EVT_BUTTON,self.importdata1)

        t2 = wx.StaticText(self, label="import backup file2", pos=(100,400))
        self.tc2 = wx.TextCtrl(self,pos=(220,400))
        b2=wx.Button(self,label="choose file2",size=(100,30),pos=(350,400))
        b2.Bind(wx.EVT_BUTTON,self.choosefile2)
        b3=wx.Button(self,label="import file2",size=(100,30),pos=(350,500))
        b3.Bind(wx.EVT_BUTTON,self.importdata2)

    def backup(self,event):
        dlg=wx.DirDialog(self,"Choose a directory",style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            path=dlg.GetPath()
        dlg.Destroy()

        wb=Workbook()
        sheet1=wb.add_sheet('Sheet 1')

        conn=sqlite3.connect('MemberData.db')
        c=conn.cursor()
        c.execute('''SELECT * FROM gym1''')
        data=c.fetchall()
        conn.close()

        coloumname=('EnrolmentNo','Name','ForHname','Dob','HAddress','OAddress','Gender','ContactNo','ResidenceNo','Occupation','Doj','FeePaid','MembershipTime')

        for j in range(0,13):
                sheet1.write(0,j,coloumname[j])

        row=0
        for i in data:
            row+=1
            for j in range(0,13):
                sheet1.write(row,j,i[j])


        wb.save(os.path.join(path, 'GYM DATABASE BACKUP1.xls'))

        
        wb=Workbook()
        sheet1=wb.add_sheet('Sheet 1')

        conn=sqlite3.connect('MemberData.db')
        c=conn.cursor()
        c.execute('''SELECT * FROM gym2''')
        data=c.fetchall()
        conn.close()

        coloumname=('EnrolmentNo','Image')

        for j in range(0,2):
                sheet1.write(0,j,coloumname[j])

        row=0
        for i in data:
            row+=1
            for j in range(0,2):
                sheet1.write(row,j,i[j])


        wb.save(os.path.join(path, 'GYM DATABASE BACKUP2.xls'))

        msgbox = wx.MessageBox('Product Database Backup Successful', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

    def choosefile1(self,event):
        
        dlg=wx.FileDialog(self,"Choose a media file","","","*.xls")
        if dlg.ShowModal()==wx.ID_OK:
            path=dlg.GetPath()
            folder,filename=os.path.split(path)
        global foldername
        global fname
        fname=filename
        foldername=folder
        
        self.tc1.SetValue(fname)
        
        dlg.Destroy()

    def choosefile2(self,event):
        
        dlg=wx.FileDialog(self,"Choose a media file","","","*.xls")
        if dlg.ShowModal()==wx.ID_OK:
            path=dlg.GetPath()
            folder,filename=os.path.split(path)
        global foldername
        global fname
        fname=filename
        foldername=folder
        
        self.tc2.SetValue(fname)
        
        dlg.Destroy()

    def importdata1(self,event):
        global foldername
        global fname
        fname='\\'+fname

        XLS_FILE = foldername + fname

        workbook = xlrd.open_workbook(XLS_FILE)
        sheet = workbook.sheet_by_name('Sheet 1')
        n=sheet.nrows
        #print 'no.of rows',n
        n+=1
        ROW_SPAN = [2, n]                                       # 2nd argument needs to be dynamically changed to add more enteries.Currently its not.   
        COL_SPAN = [1,14]                                       # Same as above. but to increase coloumns.
        app = Dispatch("Excel.Application")
        app.Visible = True
        ws = app.Workbooks.Open(XLS_FILE).Sheets(1)
        
        exceldata = [[ws.Cells(row, col).Value 
                      for col in xrange(COL_SPAN[0], COL_SPAN[1])] 
                     for row in xrange(ROW_SPAN[0], ROW_SPAN[1])]
        conn = sqlite3.connect('MemberData.db')
        c = conn.cursor()
        c.execute('''SELECT EnrolmentNo FROM gym1''')
        data=c.fetchall()

        a=[]
        #row[17]=str('')
        if len(data)==0:
            f=1
        else:
            f=0
        for row in exceldata:
            for i in data:
                if i[0]==row[0]:
                    f=0
                    break
                else:
                    f=1
            if f==1:
                a=[]
                k=0
                for i in range(0,13):
                    a.append(str(row[i]))
                    k+=1
                c.execute('INSERT INTO gym1 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', a)
                conn.commit()
                
        conn.close()
        #c.execute('DELETE FROM emp3')
        #m=Message(self)
        msgbox = wx.MessageBox('IMPORT SUCCESSFUL!', 
                       'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

    def importdata2(self,event):
        global foldername
        global fname
        fname='\\'+fname

        XLS_FILE = foldername + fname

        workbook = xlrd.open_workbook(XLS_FILE)
        sheet = workbook.sheet_by_name('Sheet 1')
        n=sheet.nrows
        #print 'no.of rows',n
        n+=1
        ROW_SPAN = [2, n]                                       # 2nd argument needs to be dynamically changed to add more enteries.Currently its not.   
        COL_SPAN = [1,3]                                       # Same as above. but to increase coloumns.
        app = Dispatch("Excel.Application")
        app.Visible = True
        ws = app.Workbooks.Open(XLS_FILE).Sheets(1)
        
        exceldata = [[ws.Cells(row, col).Value 
                      for col in xrange(COL_SPAN[0], COL_SPAN[1])] 
                     for row in xrange(ROW_SPAN[0], ROW_SPAN[1])]
        conn = sqlite3.connect('MemberData.db')
        c = conn.cursor()
        c.execute('''SELECT EnrolmentNo FROM gym2''')
        data=c.fetchall()

        conn1=sqlite3.connect('MemberData.db')
        c1=conn1.cursor()
        c1.execute('Select EnrolmentNo from gym1')
        enos=c1.fetchall()
        conn1.close()

        if len(enos)==0:
                msgbox = wx.MessageBox('First IMPORT Main Database', 
                           'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
                flag=0
        
        a=[]

        if len(data)==0:
            f=1
        else:
            f=0
        for row in exceldata:
            for i in data:                                 #to check if enrollment no already exists in gym2 database
                if i[0]==row[0]:
                    f=0
                    break
                else:
                    f=1

            
            flag=0  
            for j in enos:           #To ensure that enrollment no. already exits in gym1 database. It should already exist to avoid backup bugs.

                if int(j[0])==int(row[0]):
                    flag=1
                    break
                else:
                    flag=0

            if flag!=1 and len(enos)!=0:
                msgbox = wx.MessageBox('"%d" This Enrollment Number does not exits in Main Database!'%(int(row[0])), 
               'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
            
            if f==1 and flag==1:
                a=[]
                k=0
                for i in range(0,2):
                    a.append(str(row[i]))
                    k+=1
                c.execute('INSERT INTO gym2 VALUES (?,?)', a)
                conn.commit()
                
        conn.close()
        #c.execute('DELETE FROM emp3')
        #m=Message(self)
        if f==1 and flag==1:
            msgbox = wx.MessageBox('IMPORT SUCCESSFUL!', 
                           'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
            
        
class SeventhPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self,parent,size=wx.DefaultSize)

        text0 = wx.StaticText(self, label="Enrollment/Serial Number", pos=(100,100))
        self.tc0 = wx.TextCtrl(self,pos=(300,100))

        delete=wx.Button(self,label="Delete Entry",size=(120,40),pos=(100,200))
        delete.Bind(wx.EVT_BUTTON,self.DeleteEntry)

    def DeleteEntry(self,event):

        dlg = wx.MessageDialog(self, "Do you really want to delete this Entry?","Confirm Delete", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:

            Sno=self.tc0.GetValue()

            conn=sqlite3.connect("MemberData.db")
            c=conn.cursor()
            try:
                c.execute('delete from gym1 where EnrolmentNo="%s"'%(Sno))
                conn.commit()
                msgbox = wx.MessageBox('Entry Deleted from data', 
                                   'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
            except:
                msgbox = wx.MessageBox('datadoesnotexits in main data', 
                               'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
            
            finally:
                c.execute('delete from gym2 where EnrolmentNo="%s"'%(Sno))
                conn.commit()
                msgbox = wx.MessageBox('Entry Deleted from imageDatabase data', 
                                   'Message', wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

            conn=sqlite3.connect("MemberData.db")
            c=conn.cursor()
            try:
                c.execute('delete from gym3 where EnrolmentNo="%s"'%(Sno))
                conn.commit()
            except:
                pass
                
            
class MyApp(wx.App):
    def OnInit(self):
        frame=MainFrame(None,-1,'Gym Software')
        frame.Show()
        frame.Maximize(True)
        return True

app=MyApp()
app.MainLoop()
