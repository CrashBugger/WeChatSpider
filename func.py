import datetime
import sys
import time

import requests
from requests import Response
import xlwt
from xlwt import Worksheet


class Function():
    def __init__(self, cookie, token, starTime, endTime):
        """初始化url，依据情况修改
         params中
         fakeid---->每个公众号id
         begin->页数有关
         ajax->不用修改
         """
        self.url = "https://mp.weixin.qq.com/cgi-bin/appmsg"
        self.headers = {
            "cookie": cookie,
            "referer": f"https://mp.weixin.qq.com/cgi-bin/appmsg?t=media/appmsg_edit_v2&action=edit&isNew=1&type=77&createType=0&token={token}&lang=zh_CN",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.30",
        }
        self.data = []
        self.params = {
            "action": "list_ex",
            "begin": "",
            "count": "5",
            "fakeid": "MzIxNDc0MjU0OA==",
            "type": "9",
            "query": "",
            "token": token,
            "lang": "zh_CN",
            "f": "json",
            "ajax": "1"
        }
        self.startTime = datetime.datetime.strptime(starTime, "%Y-%m-%d").timestamp()
        self.endTime = datetime.datetime.strptime(endTime, "%Y-%m-%d").timestamp()
        self.data = []
        self.isDone = False

    def getData(self, resp: Response):
        """拿到原始数据  依据情况修改"""
        data = resp.json()
        list = data['app_msg_list']
        return list

    def parseAndSaveData(self, list: list):
        """数据解析   依据情况修改"""
        """此处返回元组列表 0为date 1为标题"""
        for item in list:
            if self.judge(item):
                title = item["title"]
                dateDay = self.transfer_time(float(item["create_time"]))
                one = [dateDay, title]
                self.data.append(one)

    def judge(self, item: dict):
        """数据判断，此处为日期判断  依据情况修改"""
        time = float(item["create_time"])
        if time < self.startTime:
            self.isDone = True
        if self.endTime >= time >= self.startTime:
            return True
        else:
            return False

    def getBegin(self, pageNum):
        """根据begin参数变化情况改写
        每个微信公众号可能不同"""
        return (pageNum - 1) * 5

    def getResp(self, pageNum: int):
        self.params["begin"] = self.getBegin(pageNum)
        resp = requests.get(url=self.url, headers=self.headers, params=self.params)
        return resp

    def transfer_time(self, s) -> str:
        aa = time.ctime(s)
        bb = aa.split()
        cc = (bb[-1] + "-" + bb[1] + "-" + bb[2]).replace('Jan', '1').replace('Feb', '2').replace('Mar', '3'). \
            replace('Apr', '4').replace('May', '5').replace('Jun', '6').replace('Jul', '7').replace('Aug', '8') \
            .replace('Sep', '9').replace('Oct', '10').replace('Nov', '11').replace('Dec', '12')
        return cc

    def saveData(self):
        workbook = xlwt.Workbook(encoding="utf-8")
        sheet: Worksheet = workbook.add_sheet("1")
        sheet.write_merge(0, 0, 0, 5, "国家网络安全学院网站内容审核登记表")
        sheet.write(1, 0, label="序号")
        sheet.write(1, 1, label="日期")
        sheet.write(1, 2, label="栏目")
        sheet.write(1, 3, label="内容")
        sheet.write(1, 4, label="审核人")
        sheet.write(1, 5, label="上传人")
        rowCount = 2
        for item in self.data:
            sheet.write(rowCount, 0, rowCount - 1)
            sheet.write(rowCount, 1, item[0])
            title = item[1].split('|')
            if len(title) > 1:
                sheet.write(rowCount, 2, title[0])
            title2 = item[1].split('丨')
            if len(title2) > 1:
                sheet.write(rowCount, 2, title2[0])
            sheet.write(rowCount, 3, item[1])
            rowCount += 1
        workbook.save("1.xls")
