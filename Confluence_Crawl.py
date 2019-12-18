# -*-coding: utf-8 -*-

import re
import os
import socket

import requests
from bs4 import BeautifulSoup

import libs.secret_key as secret_key
from libs.Deal_Xlsx import Deal_Xlsx


class Confluence_Crawl:
    def __init__(self):
        self.s = requests.Session()
        self.login()

    def login(self):
        LOG_IN_INFO = {
            "os_username": secret_key.id,
            "os_password": secret_key.pwd,
            "login": "로그인",
            "os_destination": "/index.action",
        }
        print("logging_in")
        req = self.s.post(
            "http://ldps-sltn-alb-atlassian-01-1012132539.ap-northeast-2.elb.amazonaws.com:8090/login.action?os_destination=%2Findex.action&permissionViolation=true",
            data=LOG_IN_INFO,
        )
        if req.status_code == 200 and req.ok:
            # soup = BeautifulSoup(req.content, 'html.parser')
            # print(soup.prettify())
            return True
        else:
            return False

    def get_digital_business(self):
        DB_URL = "http://ldps-sltn-alb-atlassian-01-1012132539.ap-northeast-2.elb.amazonaws.com:8090/pages/viewpage.action?pageId=6348907"
        req = self.s.get(DB_URL)
        if req.status_code == 200 and req.ok:
            # soup = BeautifulSoup(req.content, 'html.parser')
            # print(soup.prettify())
            return True

    def get_db_list(self):
        DB_LIST_URL = "http://ldps-sltn-alb-atlassian-01-1012132539.ap-northeast-2.elb.amazonaws.com:8090/pages/viewpage.action?pageId=17146586"
        req = self.s.get(DB_LIST_URL)
        if req.status_code == 200 and req.ok:
            # soup = BeautifulSoup(req.content, 'html.parser')
            # print(soup.prettify())
            return req

    def confluence_crawl(self):
        if self.login():
            req = self.get_db_list()
            soup = BeautifulSoup(req.content, "html.parser")
            # soup = soup.prettify(formatter = 'html')
            tr_list = soup.select("table.wrapped.fixed-table.confluenceTable tr")
            return tr_list
        else:
            print("status code not 200")

    def write_xlsx(self):
        pass

    def single_line(self, raw_str):
        return " ".join(raw_str.split())

    def uni_to_utf8(self, unicode_str):
        return str(unicode_str.encode("utf-8"))

    def exit(self):
        self.s.close()


if __name__ == "__main__":
    cc = Confluence_Crawl()
    excel = Deal_Xlsx()
    cur_sheet = excel.set_cur_sheet("디지털사업부문")

    temp_index = ""
    tr_list = cc.confluence_crawl()
    idx = 0
    while idx < len(tr_list):
        for tr in tr_list[1:]:
            td = tr.select("td")
            if len(td) == 4:
                index = temp_index
                content = td[0]
                origin = td[1]
                date = td[2]
                team = td[3]
            else:
                index = td[0]
                content = td[1]
                # print(content)
                origin = td[2]
                date = td[3]
                team = td[4]
                temp_index = index
                # print(index, content, origin, date, team)

            for n, i in enumerate([index, content, origin, date, team]):
                i = i.prettify(formatter="html")
                i = re.sub("<br.*>", "!!!!!", i)
                i = re.sub("<p>", "!!!!!", i)
                i = re.sub("<.*>", "", i)
                i = re.sub("\n", "", i)
                i = re.sub("&middot;", "·", i)
                i = re.sub("&rarr;", "→", i)
                i = re.sub("&nbsp;", "", i)
                i = re.sub("!!!!!", "\n", i)

                i = i.strip()

                excel.write(n + 1, idx + 4, i)
                excel.set_border(n + 1, idx + 4)
                # print(i)
                if n == 1:
                    excel.align(n + 1, idx + 4, "general", "center", True)
                else:
                    excel.align(n + 1, idx + 4, "center", "center", True)

            idx += 1

    cur_path = os.path.dirname(os.path.abspath(__file__))
    savs_path = os.path.join(cur_path, "result.xlsx")
    excel.wb.save("result.xlsx")
    cc.exit()
