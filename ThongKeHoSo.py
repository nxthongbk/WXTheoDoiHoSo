#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from selenium import webdriver
from time import sleep
import openpyxl
from openpyxl.styles import Font
import wx

wb = openpyxl.load_workbook('ma.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')


fontStyle = Font(size = "10")



class Example(wx.Frame):

    def __init__(self, parent, title):
        super(Example, self).__init__(parent, title=title)

        self.InitUI()
        self.Centre()

    def InitUI(self):

        panel = wx.Panel(self)

        sizer = wx.GridBagSizer(5, 5)

        text1 = wx.StaticText(panel, label="CHI CỤC THUẾ GÒ VẤP")
        sizer.Add(text1, pos=(0, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM,
            border=15)

        icon = wx.StaticBitmap(panel, bitmap=wx.Bitmap('thue.jpg'))
        sizer.Add(icon, pos=(0, 4), flag=wx.TOP|wx.RIGHT|wx.ALIGN_RIGHT,
            border=5)

        line = wx.StaticLine(panel)
        sizer.Add(line, pos=(1, 0), span=(1, 5),
            flag=wx.EXPAND|wx.BOTTOM, border=10)

        text2 = wx.StaticText(panel, label="Năm")
        sizer.Add(text2, pos=(2, 0), flag=wx.LEFT, border=10)

        tc1 = wx.TextCtrl(panel)
        sizer.Add(tc1, pos=(2, 1), span=(1, 3), flag=wx.TOP|wx.EXPAND)

        text3 = wx.StaticText(panel, label="File Nhập")
        sizer.Add(text3, pos=(3, 0), flag=wx.LEFT|wx.TOP, border=10)

        tc2 = wx.TextCtrl(panel)
        sizer.Add(tc2, pos=(3, 1), span=(1, 3), flag=wx.TOP|wx.EXPAND,
            border=5)

        button1 = wx.Button(panel, label="Browse...")
        sizer.Add(button1, pos=(3, 4), flag=wx.TOP|wx.RIGHT, border=5)

        text4 = wx.StaticText(panel, label="File Xuất")
        sizer.Add(text4, pos=(4, 0), flag=wx.TOP|wx.LEFT, border=10)

        combo = wx.ComboBox(panel)
        sizer.Add(combo, pos=(4, 1), span=(1, 3),
            flag=wx.TOP|wx.EXPAND, border=5)

        button2 = wx.Button(panel, label="Browse...")
        sizer.Add(button2, pos=(4, 4), flag=wx.TOP|wx.RIGHT, border=5)

        sb = wx.StaticBox(panel, label="Nâng Cao")

        boxsizer = wx.StaticBoxSizer(sb, wx.VERTICAL)
        boxsizer.Add(wx.CheckBox(panel, label="Xuất Tất Cả"),
            flag=wx.LEFT|wx.TOP, border=5)
        boxsizer.Add(wx.CheckBox(panel, label="Xuất hồ sơ thành công"),
            flag=wx.LEFT, border=5)
        boxsizer.Add(wx.CheckBox(panel, label="Xuất hồ sơ hoàn"),
            flag=wx.LEFT|wx.BOTTOM, border=5)
        sizer.Add(boxsizer, pos=(5, 0), span=(1, 5),
            flag=wx.EXPAND|wx.TOP|wx.LEFT|wx.RIGHT , border=10)

        button3 = wx.Button(panel, label='Help')
        sizer.Add(button3, pos=(7, 0), flag=wx.LEFT, border=10)

        button4 = wx.Button(panel, label="Ok")
        sizer.Add(button4, pos=(7, 3))
        button4.Bind(wx.EVT_BUTTON,self.OnClicked)

        button5 = wx.Button(panel, label="Cancel")
        sizer.Add(button5, pos=(7, 4), span=(1, 1),flag=wx.BOTTOM|wx.RIGHT, border=10)

        sizer.AddGrowableCol(2)

        panel.SetSizer(sizer)
        sizer.Fit(self)
        
    def run_hs(self):
        wb = openpyxl.load_workbook('ma.xlsx')
        sheet = wb.get_sheet_by_name('Sheet1')

        from openpyxl.styles import Font
        fontStyle = Font(size = "10")


        driver = webdriver.Firefox(executable_path=r'geckodriver.exe')
        driver.get('http://www.vnpost.vn/')
        sleep(1)
        driver.find_element_by_xpath("//span[@class=\"caret\"]").click()
        sleep(1)
        driver.find_element_by_xpath("//input[@name=\"dnn$ctl11$txtTrackItem\"]").send_keys("RJ714626641VN")
        sleep(1)
        driver.find_element_by_xpath("//button[@id=\"dnn_ctl11_btnTrackItem\"]").click()
        a = driver.find_element_by_xpath("//label[text()=\"Trạng thái\"]/..//strong").text;
        print(a)
        dict={}
        # for item in ['RJ714626655VN','RJ714608603VN','RJ714608634VN','RJ714608617VN','RJ714571276VN','RJ714571293VN','RJ714571280VN','RJ714582415VN','RJ714582407VN'
        # 'RJ714615269VN',
        # 'RJ714585037VN',
        # 'RJ714585023VN',
        # 'RJ714585010VN' 
        # ]:

        for i in range(8,22):
            sleep(5)
            item = sheet['AB' + str(i)].value
            update_item =  'AC' + str(i)
            driver.find_element_by_xpath("//input[@name=\"dnn$ctr734$View$uc$txtItemCode\"]").clear()
            driver.find_element_by_xpath("//input[@name=\"dnn$ctr734$View$uc$txtItemCode\"]").send_keys(item)
            driver.find_element_by_xpath("//button[@id=\"dnn_ctr734_View_uc_btnSearch\"]").click()
            a = driver.find_element_by_xpath("//label[text()=\"Trạng thái\"]/..//strong").text;
            #sheet['A1'] = 'Hello world!'
            #sheet[update_item] = a
            dict.update({item:a})
            sheet[update_item] = a
            
            sheet[update_item].font = fontStyle

        print(dict)

        #worksheet= myworkbook['Sheet1'] 
        #worksheet['A21']='We are writing to B4'
        #sheet[update_item] 
        wb.save('ma.xlsx')


        driver.quit()
        driver.close 
        
    def OnClicked(self, event): 
        btn = event.GetEventObject().GetLabel() 
        self.run_hs()
    
    
      
def main():

    app = wx.App()
    ex = Example(None, title="Thống Kê Hồ Sơ - Chi Cục Thuế Gò Vấp")
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()