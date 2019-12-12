# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a program to reduce the work that classifying the interns from each hospital.
"""

import xlsxwriter
import openpyxl
from timeutil import timeutil
TimeUtil = timeutil.TimeUtil
import pandas


class RotatingWriter(object):
    def __init__(self):
    
        pass
    
    
    
    def setFormats(self):
        bodyFormat = self.workbook.add_format({
            'bold':     False,
            'size':    14,            
            'border':   1,
            'align':    'center',
            'valign':   'vcenter',
            'text_wrap' : 1,
            'num_format': 'mm-dd',
#            'fg_color': '#D7E4BC',
        })  
        self.bodyFormat = bodyFormat

        titleFormat = self.workbook.add_format({
            'bold':     True,
            'size':    22,
            #'border':   6,
            'align':    'center',
            'valign':   'vcenter',
            #'fg_color': '#D7E4BC',
        })          
        self.titleFormat = titleFormat
    
    def createXls(self, fileNamePara):
        """
        Create the work sheet and set the format.
        """
        workbook = xlsxwriter.Workbook(fileNamePara)   
        worksheet = workbook.add_worksheet('sheet')             

        self.workbook = workbook
        self.worksheet = worksheet
        
        self.setFormats()

    def setColWidth(self):

        colWidth = [8.38 , 8.38 , 8.38 , 8.38 , 8.38 , 8.38 , 
                    8.38 , 8.38 , 8.38 , 8.38 , 8.38 , 8.38 , 
                    8.38 , 8.38 , 8.38 , 8.38 , 8.38 , 8.38 , 8.38 ,
                     ]        
        for i in range(len(colWidth)):
            self.worksheet.set_column(i, i, colWidth[i])        
            
        self.worksheet.set_column(0, 0, 16)          

    def setRowHeight(self):

        rowHeight = [20.25, 20.25, 76.5, 80, 30, 30, 30, 30, 30, 30, 30, 30, 30,30,
                     30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30,30,
                     30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30,30,
                     30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30,30,
                     ]
        for i in range(len(rowHeight)):
            self.worksheet.set_row(i, rowHeight[i])
                    
    
    def receivePandasDataFrame(self, xlsxNamePara):
        """
        Read the original excel as pandas dataFrame.
        """
        self.dataFrame = pandas.read_excel(xlsxNamePara)


    def iteratelyCopy(self, hospitalNamePara):
        """
        iterate the dataFrame and format the cells.
        """
        
        rows, cols  = obj.dataFrame.shape
        for i in range(rows):
            for j in range(cols):

                if pandas.isnull(obj.dataFrame.iloc[i][j]):
                    self.worksheet.write(i+1, j, u"", self.bodyFormat)
                else:
                    temp = RotatingWriter.deHospitalizeNames(self.dataFrame.iloc[i][j], hospitalNamePara)
                    self.worksheet.write(i+1, j, temp, self.bodyFormat)
        

    def writeContentForRecord(self):
        """
        write the special lines
        """
        self.worksheet.merge_range('A1:S1', u"广州市越秀区白云街社区卫生服务中心",  self.titleFormat)
        self.worksheet.merge_range('A2:S2', u"白云社卫全科医师规培社区轮转实习安排时间表（2019年11月1日~2020年4月30日）",  self.titleFormat) 
        self.worksheet.merge_range('A3:C3', u"时间\日期",  self.bodyFormat) 
        self.worksheet.merge_range('A4:C4', u"带教老师",  self.bodyFormat)   
#       self.worksheet.write('A4', u"时间",self.bodyFormat)



        
        
    def save(self):
        self.workbook.close()

    @staticmethod
    def deHospitalizeNames(namePara, selectPara):
        """
            select all the cells that ends with selectPara, and then replace it and return the formatted unicode.
        """
        if type(namePara) == unicode:
            if "-" in namePara and "--" not in namePara:
                if selectPara in namePara:
                    return namePara.replace(selectPara, u"")
                else:
                    return u""
        else:
            return namePara

    @staticmethod
    def tester():
        print "test dehospitalizeName"
        name = u"刘刘刘-中一"
        print RotatingWriter.deHospitalizeNames(name, u"-中一")
        print RotatingWriter.deHospitalizeNames(name, u"-省医")
        print RotatingWriter.deHospitalizeNames(name, u"-广医")

if __name__ == "__main__":

    obj = RotatingWriter()
    obj.receivePandasDataFrame(u"规培时间安排.xlsx")
    
    obj.createXls(u"轮科广医版.xlsx")
    obj.setColWidth()
    obj.setRowHeight()
    obj.iteratelyCopy(u"-广医")
    obj.writeContentForRecord()
    obj.save()    

    obj.createXls(u"轮科省医版.xlsx")
    obj.setColWidth()
    obj.setRowHeight()
    obj.iteratelyCopy(u"-省医")
    obj.writeContentForRecord()
    obj.save()    
    
    obj.createXls(u"轮科中一版.xlsx")
    obj.setColWidth()
    obj.setRowHeight()
    obj.iteratelyCopy(u"-中一")
    obj.writeContentForRecord()
    obj.save()    

