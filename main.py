#!/usr/bin/python3

import argparse
import json
import operator
import os
import re
import requests
import time
import sys

from datetime import datetime
from datetime import timedelta
from configparser import ConfigParser

# third lib
import requests
import xlsxwriter

from pyvirtualdisplay import Display
from selenium import webdriver

# my lib
from lib.demail import Email


_now = datetime.now()
last_month_last_date = _now.replace(day=1) - timedelta(days=1)
start_date_str = "%s-%s-01" %(last_month_last_date.year, last_month_last_date.month)
end_date_str = "%s-%s-%s"  % (last_month_last_date.year, last_month_last_date.month, last_month_last_date.day)
#start_date_str = "2016-5-1"
#end_date_str = "2016-5-31"

SEND_DAY = 1  # day of every month
USER_CONF_PATH = "%s/.AutoScriptConfig/tower-overtime-reportor/user.ini" % os.getenv("HOME")
CAL_URL = "https://tower.im/teams/35e3a49a6e2e40fa919070f0cd9706c8/calendar_events/?start=%s&end=%s" % (start_date_str, end_date_str)
OVERTIME_CALENDAR_GUID = "b96e5a357a884c7e8c5c2ab12858dd02"

MAIL_RECEIVERS = "yinghongli@deepin.com,zhangfengling@deepin.com,zhangmingzhu@deepin.com,wangdi@deepin.com"
#MAIL_RECEIVERS = "tangcaijun@deepin.com"
MAIL_CC = "tangcaijun@deepin.com"

BASE_TOWER_URL = "tower.im/api/v2"
TOWER_API = "https://%s" % BASE_TOWER_URL


class ConfigController:

    def __init__(self):
        self.tower_token = ""

    def get_login_info(self):
        config = ConfigParser()
        config.read(USER_CONF_PATH)
        username = config["USER"]["UserName"]
        passwd = config["USER"]["UserPWD"]
        return username, passwd


    def get_tower_token(self):
        if self.tower_token == "":
            config = ConfigParser()
            config.read(USER_CONF_PATH)
            username = config.get("USER", "UserName")
            passwd = config.get("USER", "UserPWD")
            client_id = config.get("DEEPIN", "ClientId")
            client_secret = config.get("DEEPIN", "ClientSecret")

            url = "https://%s:%s@%s/oauth/token" % (client_id, client_secret, BASE_TOWER_URL)
            d = {"grant_type":"password", "username": username, "password": passwd}
            success, data = self.__sendRequest(url, d)

            if success:
                self.tower_token = data.get("access_token")
            else:
                print("E: get tower access token error", file=sys.stderr)

        return self.tower_token


    def __sendRequest(self, url, d={}, h={}, method='POST'):
        if method == 'POST':
            resp = requests.post(url, data=d, headers=h)
        elif method == 'GET':
            resp = requests.get(url, data=d, headers=h)
        else:
            print("request method not supported")
            return False, None

        if resp.ok:
            return True, resp.json()

        print ("E: send request error: %s" % resp.text, file=sys.stderr)
        return False, None


class BrowserController:

    def __init__(self):
        self.browser = webdriver.Firefox()
        self.cc = ConfigController()
        (username, passwd) = self.cc.get_login_info()
        self.login(username, passwd)


    def login(self, username, passwd):

        print("login to tower ...")
        login_url = "https://tower.im/users/sign_in"
        self.browser.get(login_url)
        unEL = self.browser.find_element_by_id("email")
        pwdEL = self.browser.find_element_by_name("password")
        unEL.send_keys(username)
        pwdEL.send_keys(passwd)
        unEL.submit()

        time.sleep(5)

        # check login status
        if self.browser.current_url == "https://tower.im/teams/35e3a49a6e2e40fa919070f0cd9706c8/projects/":
            print ("login successfully")
            return True

        else:
            print ("login error, current url (%s) does not match." % self.browser.current_url)
            return False


    def get_calendar_events(self):
        self.browser.get(CAL_URL)
        source = self.browser.page_source
        el = self.browser.find_element_by_tag_name("body")
        text = el.text
        data = json.loads(text)
        return data


class OvertimeAnalyze:

    def __init__(self):
        self.cc = ConfigController()

    def work(self):
        #with open("data.json") as fp:
        #    data = json.load(fp)
        print("analyzing calendar ...")
        overtime_datas = self.analyze()
        month_str = self.get_month_str()
        file_name = "%s_overtime.xlsx" % (month_str)

        print("writing to excel ...")
        self.write_excel(overtime_datas, file_name)

        if self.check_send_day():
            subject = "%s 加班信息统计" % month_str
            print("sending email ...")
            self.send_email(subject=subject, files=[file_name])
        else:
            print("sorry, today is not the email-sending day, abort sending email.")

        print("finish.")

    
    def get_month_str(self):
        tmp_str_list = start_date_str.split("-")
        month_str = "%s-%s" % (tmp_str_list[0], tmp_str_list[1])
        return month_str


    def analyze(self, cal_data=None):

        overtime_datas = []

        if not cal_data:
            bc = BrowserController()
            cal_data = bc.get_calendar_events()

        from pprint import pprint
        pprint(cal_data)

        for item in cal_data.get("calendar_events", []).copy():
            data = {}
            caleventable_guid = item.get("caleventable_guid")
            if caleventable_guid != OVERTIME_CALENDAR_GUID:
                continue
            guid = item.get("guid")

            nickname = item.get("creator").get("nickname")

            starts_at = item.get("starts_at", "")
            ends_at = item.get("ends_at", "")
            time_re = re.compile("(\d+-\d+-\d+)T(\d+:\d+)*")
            starts_at = " ".join(time_re.findall(starts_at)[0])
            ends_at = " ".join(time_re.findall(ends_at)[0])

            content = item.get("content", "")

            token = self.cc.get_tower_token()
            reminders = self.get_reminders(token, guid)

            data["nickname"] = nickname
            data["starts_at"] = starts_at
            data["ends_at"] = ends_at
            data["content"] = content
            data["reminders"] = reminders

            overtime_datas.append(data)

        return overtime_datas


    def get_reminders(self, token, guid):
        h = {
                "Authorization":"Bearer %s" % (token, )
            }

        url = "https://tower.im/api/v2/events/%s" % (guid, )

        r = requests.get(url=url, headers=h)
        j = r.json()
        comments = j.get("comments")
        if len(comments):
            content = comments[0].get("content")
            a = re.compile(">@(\w+)<")
            namelist = a.findall(content)
            #print(namelist)
            return " ".join(namelist)

        return ""


    def write_excel(self, overtime_datas, file_name="./overtime.xlsx"):
        COLUMN = {
                "nickname": "A",
                "content": "B",
                "starts_at": "C",
                "ends_at": "D",
                "reminders": "E",
        }

        # sort data
        overtime_datas = sorted(overtime_datas, key=lambda k: k["starts_at"])

        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet()

        self.prepare_headers(workbook, worksheet)

        # formats
        common_format = workbook.add_format()
        common_format.set_text_wrap()
        common_format.set_align("top")

        row = 3  # start row
        for data in overtime_datas:
            for title, content in data.items():
                pos = "%s%d" %(COLUMN.get(title), row)
                worksheet.write(pos, content, common_format)
            row += 1


    def prepare_headers(self, workbook, worksheet):

        # write table header
        constant = {
                "B1": "%s 加班信息统计" % (self.get_month_str()),
                "A2": "姓名",
                "B2": "内容",
                "C2": "起始时间",
                "D2": "结束时间",
                "E2": "部门负责人",
                }

        # set content column width
        worksheet.set_column('A:A', 6) 
        worksheet.set_column('B:B', 30) 
        worksheet.set_column('C:E', 15) 

        # format
        basic_format = workbook.add_format()
        basic_format.set_align("vcenter")
        basic_format.set_align("center")
        basic_format.set_bold()

        # write basic headers
        for p, content in constant.items():
            worksheet.write(p, content, basic_format)

    def check_send_day(self):
        now = datetime.now()
        if now.day == SEND_DAY:
            return True

        return False


    def send_email(self, subject, files):
        e = Email()
        content = ""
        e.send(MAIL_RECEIVERS, subject, content, CC=MAIL_CC, files=files, auto_close=True, use_footer=True)


if __name__ == "__main__":
    display = Display(visible=0, size=(1366, 768))
    display.start()
    oa = OvertimeAnalyze()
    oa.work()
    display.stop()
