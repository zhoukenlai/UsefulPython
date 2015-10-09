#-*- coding:utf-8 -*-

from HTMLParser import HTMLParser
#from openpyxl import Workbook
#from openpyxl.writer.excel import ExcelWriter
import datetime
from xlwt import *
import re
import cx_Oracle
import os

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.ZHS16GBK'
 
class MyHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.parserdata = []
        self.sdate=""
        self.flag=False
 
    def handle_starttag(self, tag, attrs):
        pass
        #print tag

    def handle_data(self,data):
        data=(data.decode('gb2312')).strip()
        if len(data)>0:
            i=len(self.parserdata)-1
            m=re.search(u'^(\d*)、(.*)',data)
            n=re.search(u'(\d*)月(\d*)日$',data)
            if bool(n):
                self.sdate=data;
            if bool(m):
                #self.parserdata.append(self.sdate+"|A"+data+"|A");
                self.parserdata.append(data);
                self.flag=True
            if bool(m)==False and bool(n)==False :
                if i>=0:
                    if len(self.parserdata[i])>10 and self.flag:
                        self.parserdata[i] += "|A"
                        self.flag=False 
                    self.parserdata[i] += data

    def handle_endtag(self, tag):
        pass

    def SaveExcel2003(self):
        wb = Workbook(optimized_write = True)
        ws = wb.create_sheet()

        ws.append([u"标题",u"内容"])
        for x in hp.parserdata:
            ws.append(x.split('|A'))

        wb.save('out1111.xlsx')

    def SaveExcel2007(self):
        book = Workbook(encoding='utf-8')
        sheet = book.add_sheet('Sheet',cell_overwrite_ok=True)

        i=0
        for x in hp.parserdata:
            ls = x.split('|A')
            if len(ls)!=2:
                pass
            else:
                sheet.write(i, 0, ls[0])
                sheet.write(i, 1, ls[1])
                i = i+1

        now = datetime.datetime.now()
        book.save('xxx'+now.strftime('%Y%m%d%H%M')+'.xls')

def getDB():
    conn = cx_Oracle.connect('username/pwd@db')
    cursor = conn.cursor()
    sql = u"""select * from tab"""
    cursor.execute(sql)
    row=cursor.fetchall()

    book = Workbook(encoding='utf-8')
    sheet = book.add_sheet('Sheet',cell_overwrite_ok=True)

    i=0
    for x in row:
        if len(x)!=5:
            pass
        else:
            sheet.write(i, 0, x[0].decode('gbk'))
            sheet.write(i, 1, x[1].decode('gbk'))
            sheet.write(i, 2, x[2].decode('gbk'))
            sheet.write(i, 3, x[3].decode('gbk'))
            sheet.write(i, 4, x[4].decode('gbk'))
            i = i+1

    now = datetime.datetime.now()
    book.save('xxx'+now.strftime('%Y%m%d')+'.xls')

 
if __name__ == "__main__":

    getDB();

    f = open(u"xxx.htm",'r')
    html_code = f.read()
    hp = MyHTMLParser()
    hp.feed(html_code)
    
    hp.SaveExcel2007();
        
    hp.close()


