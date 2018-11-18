from threading import Thread
import time
import xlrd   #这个包需要下载
from xlrd import xldate_as_tuple  
import os
import sys
from openpyxl import Workbook  #这个包需要下载
import requests
import time
from bs4 import BeautifulSoup as bs
from lxml import etree
import json
from datetime import datetime
 
class StockSpider(Thread):
	def __init__(self, stock,datapath):
		
		super(StockSpider, self).__init__()
		self.stock_id = stock["number"].split(".")[0]
		self.date = datetime(*xldate_as_tuple(stock["date"] , 0))
		self.enddate = datetime(*xldate_as_tuple(stock["date"]+60 , 0))
		self.stock_time =  self.date.strftime('%Y-%m-%d')    #excel中存储的日期格式是纯数字，要变
		self.end_stock_time =  self.enddate.strftime('%Y-%m-%d')    #excel中存储的日期格式是纯数字，要变
		
		self.datapath = datapath
		self.url = "http://www.cninfo.com.cn/new/fulltextSearch/full?searchkey="+self.stock_id+"&sdate="+self.stock_time+"&edate="+self.end_stock_time+"&isfulltext=false&sortName=nothing&sortType=desc&pageNum=1"
		self.stock_dir = datapath+"/"+self.stock_id+"_"+self.stock_time
		self.headers = {
			'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36'
		}
		

	def make_dirs(self,name):
		isExists=os.path.exists(name)
		if not isExists:
			os.makedirs(name)

	def savePDF(self,filepath,pdfURL):
		#res = requests.get('http://lxml.de/lxmldoc-4.1.1.pdf')
		pdfURL = "http://static.cninfo.com.cn/"+pdfURL

		if ".pdf" not in pdfURL:
			print("非pdf文件")
			return
			
		print(filepath)
		print(pdfURL)
		try:
			res = requests.get(pdfURL)
			res.encoding = res.apparent_encoding
			with open(filepath, 'wb') as f:
				f.write(res.content)
				f.close()
		except:
			print("保存pdf出错")

	def run(self):
		self.parse_page()

	def parse_page(self):
		files =  os.listdir(self.datapath) 
		if self.stock_id+"_"+self.stock_time in files:
			return
		
		try:
			response=requests.post(self.url,self.headers)
			jsonStr = response.content.decode('utf8')
			obj = json.loads(jsonStr)
		except:
			print("url连接错误")

		
		array = obj["announcements"]      #取出数据元素
		if len(array) >0 :
			self.make_dirs(self.stock_dir)  #连接成功才创建文件夹
			print(self.url)
			for passage in array:
				
				if "证监会" in passage["announcementTitle"] or "证券监督管理" in passage["announcementTitle"] :
					name = passage["announcementTitle"]
					if "：" in name:
						name = name.split("：")[1]
					pdfPATH = self.stock_dir+"/"+name+".pdf"
					self.savePDF(pdfPATH,passage["adjunctUrl"])
		else:
			print("找不到资源")

def readDataExcel(filepath):
  data = xlrd.open_workbook(filepath)
  table = data.sheets()[0]  # 获取excel中第一个sheet表
  datalist = []
  row = table.nrows
  col = len(table.row_values(0))-1
  totalData = []

  for i in range(1,table.nrows):
      row_data = {}
      row_data["number"] = table.row_values(i)[0]
      row_data["name"]  = table.row_values(i)[1]   
      row_data["date"]  =  table.row_values(i)[2]

      totalData.append(row_data)
  
  return totalData

def main():
	datapath =  sys.path[0]+"/Mydata"
	stockList = readDataExcel("相关资料简略版.xlsx")

	# 保存线程
	Thread_list = []
	# 创建并启动线程
	for stock in stockList:
		try:
			p = StockSpider(stock,datapath)
			p.start()
			Thread_list.append(p)
		except:
			print("错误")
			print(stock )
		
	# 让主线程等待子线程执行完成
	for i in Thread_list:
		i.join()


 
if __name__=="__main__":
	start = time.time()
	main()
	print ('[info]耗时：%s'%(time.time()-start))