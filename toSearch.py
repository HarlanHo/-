# -*- coding: utf-8 -*-
from mysqldb import *
import json
import os
from openpyxl import load_workbook
import sys
import PySide2
from PySide2.QtWidgets import QApplication, QMessageBox,QFileDialog
from PySide2.QtUiTools import QUiLoader
import xlrd

dirname = os.path.dirname(PySide2.__file__)
plugin_path = os.path.join(dirname, 'plugins', 'platforms')
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path


#生成窗口
class mainWindows:
    def __init__(self):
        # 从文件中加载UI

        self.ui = QUiLoader().load('ui/mainwindow.ui')#主窗口
        self.ui.pushButton_1.clicked.connect(self.toSearch)
        self.ui.pushButton_2.clicked.connect(self.toUploadWindows)
        self.ui.pushButton_3.clicked.connect(self.toAboutWindows)
        self.ui.lineEdit.textChanged.connect(self.toSearch)


        self.uiUpload = QUiLoader().load('ui/upload.ui')    #提交试题的窗口
        self.uiUpload.pushButton_2.clicked.connect(self.toUpload)
        self.uiUpload.pushButton_1.clicked.connect(self.openExcel)


        self.uiAbout = QUiLoader().load('ui/about.ui')      #关于的窗口


        #连接数据库
        self.query=db('127.0.0.1',3306,'xxx','xxx123','xxx')
        self.query.connect()

    def toSearch(self):
        if self.ui.lineEdit.text()=="":
            return
        self.ui.textBrowser.clear()            #先清空文本框内容
        keys=self.ui.lineEdit.text()         #获取输入框的内容

        #把关键词中的单双引号过滤转换为多关键字搜索
        keys.replace("'",'"')       
        keysList=keys.split('"')
        sql='SELECT * FROM datas WHERE '
        #sql="SELECT * FROM datas WHERE FIND_IN_SET('%s',question);" 
        t_="AND ".join("question LIKE '%"+key+"%' "  for key in keysList if key!="")
        sql=sql+t_+";"


        print(sql)
        res=self.query.query(sql)
        print(res)
        for question in res:
            print(question)
            if question!="":
                self.ui.textBrowser.append("题目"+str(question[0])+"、"+question[3]+"\n\n"+"答案:"+question[4]+"\n\n")

    def toUploadWindows(self):
        self.uiUpload.show()

    def toAboutWindows(self):
        self.uiAbout.show()

    def toUpload(self):
      try:
        path=self.uiUpload.lineEdit.text() 
        if path=="":
            QMessageBox.about(self.uiUpload,'错误',"路径及文件名不能为空！")
            return
        try:            #容错,防止读取错误
            xlsFile = xlrd.open_workbook(path, formatting_info=True)
            xlsSheet = xlsFile.sheet_by_name("Sheet1")
        except:
            QMessageBox.about(self.uiUpload,'错误',"打开文件失败。")
            return
        rows=xlsSheet.nrows          #xls的行数
        cols=xlsSheet.ncols          #xls的列数
        sqlNum="select count(*) as TOTAL from datas;"
        countNum=self.query.query(sqlNum)[0][0]     #读取datas表里记录数目
        
        for r in range(1,rows):

            countNum=countNum+1
            valueRow=xlsSheet.row_values(r)
            id_=valueRow[0]
            type_=valueRow[1]
            level=valueRow[2]
            question=valueRow[3]
            answer=valueRow[4]
            #sqlIsExist=""

            sqlUploadQue="INSERT INTO datas(id,type,level,question,answer) values(%s,%s,%s,'%s','%s');" % (countNum,type_,level,question,answer)
            res=self.query.query(sqlUploadQue)
        QMessageBox.about(self.uiUpload,'成功',"试题上传成功。")
      except:
        QMessageBox.about(self.uiUpload,'错误',"试题上传失败！")







    def openExcel(self):

        filePathDialog=QFileDialog(self.uiUpload)
        filePath = filePathDialog.getOpenFileName(self.uiUpload, "标题")
        if len(filePath)!=0:
            self.uiUpload.lineEdit.clear()
            self.uiUpload.lineEdit.setText(filePath[0])



app = QApplication([])
win = mainWindows()
win.ui.show()
app.exec_()




