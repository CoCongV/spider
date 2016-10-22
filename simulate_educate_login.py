#-*-coding:utf-8-*-
import urllib
import urllib2
import cookielib
import re
from datetime import datetime
import xlwt
from bs4 import BeautifulSoup

class JDZXY(object):

    def __init__(self, name, stuid, password):
        self.name = name
        str_now = str(datetime.now())
        pattern = re.compile(r'\s')
        datetime_now = re.sub(pattern, '%20', str_now)[:-7]
        pattern_access = re.compile(r'[\-\s\:]')
        AccessID =re.sub(pattern_access, r'', str_now)
        self.loginUrl = 'http://218.65.59.52/JwXs/LoginCheck.asp?datetime=' + datetime_now
        self.gradeUrl = 'http://218.65.59.52/jwxs/Xsxk/Xk_CjZblist.asp?flag=find'
        self.stuid = stuid
        self.cookies = cookielib.CookieJar()
        self.postdata = urllib.urlencode({
            'loginLb': 'xs',
            'Account': self.stuid,
            'PassWord': password,
            'x': '17',
            'y': '14',
            'AccessID': AccessID,
        })
        self.handler = urllib2.HTTPCookieProcessor(self.cookies)
        self.opener = urllib2.build_opener(self.handler)

    def getPage(self):
        request = urllib2.Request(url=self.loginUrl, data=self.postdata)
        result = self.opener.open(request)
        # print result.read().decode('gbk')

    def getGrade(self):
        self.grade_time = raw_input(u'查询学期，格式:xxxx-xxxx-x(2015-2016-2)\n')
        postdata = urllib.urlencode({
            'XH': self.stuid,
            'Xnxqh': self.grade_time,
        })
        requset = urllib2.Request(url=self.gradeUrl, data=postdata)
        grade = self.opener.open(requset)
        grade = grade.read().decode('gbk')
        print grade
        return grade

    def buildXls(self, grade):
        soup = BeautifulSoup(grade, 'lxml')
        items = soup.form.find_all('td')
        num = 0
        row = 0
        col = 0
        wb =xlwt.Workbook()
        sheet = wb.add_sheet(sheetname='%s' % 'grade')
        for item in items:
            text = item.getText()
            sheet.write(row, col, text)
            num += 1
            if num%12 == 0 and num != 0:
                row += 1
                col = 0
                print item.getText(), '\n'
            else:
                col += 1
                print item.getText()
        wb.save(u'成绩查询%s.xls' % self.stuid)

    def start(self):
        self.getPage()
        grade = self.getGrade()
        self.buildXls(grade)

jdz = JDZXY(name='grade', stuid='xxxxxxxxxx', password='xxxxxxxxxxx')
jdz.start()
