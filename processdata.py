#!/usr/bin/env python
# author:wujianqiang
# date:  2018-11-19

from configparser import ConfigParser
import sys,logging
import xlrd,xlsxwriter
logging.basicConfig(level=logging.ERROR,
                    filename="access.log",
                    format='%(asctime)s-%(levelname)s::%(message)s')
class Processdata(object):
    def __init__(self):
        config = ConfigParser()
        config.read("config.ini")
        self.key = config.get("key","key")
        self.title = config.get("title","title")
        self.fromfile = config.get("fromfile","file")
        self.fromsheet = config.get("fromsheet","sheet")
        self.tofile = config.get("tofile","file")
        self.tosheet = config.get("tosheet","sheet")

    def getalldatas(self):
        try:
            rd = xlrd.open_workbook(self.fromfile)
        except:
            logging.error("this file (%s) is no exits"%self.fromfile)
            sys.exit(0)
        try:
            st = rd.sheet_by_name(self.fromsheet)
        except:
            logging.error("this sheet (%s) is no exits"%self.fromsheet)
            sys.exit(0)
        rows = st.nrows
        titles = st.row_values(0)
        datas = []
        for row in range(1,rows):
            values = st.row_values(row)
            datas.append(values)
        return titles, datas

    def parserdata(self):
        titles, datas = self.getalldatas()
        key = self.key.split(";")
        title = self.title.split(";")
        key_index = []
        for k in key:
            if k in titles:
                key_index.append(titles.index(k))
            else:
                logging.error("this key (%s) is no exits"%k)
                sys.exit(0)
        title_index = []
        for t in title:
            if t in titles:
                title_index.append(titles.index(t))
            else:
                logging.error("this title (%s) is no exits"%t)
                sys.exit(0)
        todatas = {}
        for data in datas:
            todatas.setdefault(tuple([data[i] for i in key_index]),[]).append([data[i] for i in title_index])
        return todatas

    def writedata(self):
        wk = xlsxwriter.Workbook(self.tofile)
        wt = wk.add_worksheet(self.tosheet)
        datas = self.parserdata()
        title = self.title.split(";")
        title.append("count")
        j = 0
        for t in title:
            wt.write(0,j,t)
            j += 1
        i = 1
        for v in datas.values():
            wt.write_row(i,0,v[0])
            wt.write(i,len(v[0]),len(v))
            i += 1
        wk.close()

if __name__ == "__main__":
    pd = Processdata()
    pd.writedata()
