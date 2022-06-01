# -*- coding:utf-8 -*-
import json
import requests
import pandas
import arrow
import base64
from Crypto.Cipher import PKCS1_v1_5 as Cipher_pksc1_v1_5
from Crypto.PublicKey import RSA


def encrpt(password):
    public_key = """-----BEGIN PUBLIC KEY-----
    MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBAJ/d1OkCHxhTV0AkVdJhGNrvneX1biuLvBEmT0PJzX4H5zCXmVBZpWuJVyqSGaY0GHMTdYgvLtA+8giTmlD9mLUCAwEAAQ==
    -----END PUBLIC KEY-----
    """
    rsakey = RSA.importKey(public_key)
    cipher = Cipher_pksc1_v1_5.new(rsakey)
    cipher_text = base64.b64encode(cipher.encrypt(password.encode()))
    return cipher_text.decode()


class AmroWeb:
    session = requests.session()
    header = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
    }
    indexUrl = "http://me.sichuanair.com/login.shtml"
    loginUrl = "http://me.sichuanair.com/api/v1/security/loginSC"
    vcodeUrl = "http://me.sichuanair.com/verifyScCode"
    attnUrl = "http://me.sichuanair.com/api/v1/plugins/LM_ATTN_LIST"
    loginPayload = {"username": "员工号", "userPassword": encrpt("密码"), "vCode": ""}

    def __init__(self):
        AmroWeb.session.headers.update(self.header)

    def update_cookies(self, cookies_dict):
        AmroWeb.session.cookies = requests.utils.cookiejar_from_dict(cookies_dict, cookiejar=None, overwrite=True)

    def is_login(self):
        datenow = arrow.now()
        flightDate = datenow.format('YYYY-MM-DD')
        flightDate1 = datenow.shift(days=1).format('YYYY-MM-DD')
        payload = {
            "base4code": "ZUUU", "flightDate": flightDate, "flightDate1": flightDate1, "actype1": "()",
            "page": "1", "rows": "1"
        }
        r = AmroWeb.session.post("http://me.sichuanair.com/api/v1/plugins/LM_FLIGHT_SEARCH_LIST", data=payload)
        if json.loads(r.text)['code'] != 200:
            r = AmroWeb.session.post(self.loginUrl, data=self.loginPayload)
            if json.loads(r.text)['code'] != 200:
                return False
        return True

    def get_vcode(self):
        AmroWeb.session.cookies = requests.utils.cookiejar_from_dict({}, cookiejar=None, overwrite=True)
        AmroWeb.session.get(self.indexUrl)
        r = AmroWeb.session.get(self.vcodeUrl)
        return r.content

    def login(self, vcode=''):
        self.loginPayload['vCode'] = vcode
        AmroWeb.session.post(self.loginUrl, data=self.loginPayload)

    def getteams(self, team=2):
        teamurl = "http://me.sichuanair.com/api/v1/plugins/LM_QUERY_TEAM"
        payload = None
        if team == 2:
            payload = {
                "baseCode": "CD01", "deptNo": "000009007005", "FunctionCode": "LM_QUERY_TEAM"
            }
        if team == 4:
            payload = {
                "baseCode": "CD01", "deptNo": "000009007007", "FunctionCode": "LM_QUERY_TEAM"
            }
        r = self.session.post(teamurl, data=payload)
        teams = json.loads(r.text)["data"]
        return teams

    def getemps(self, team=2):
        empsurl = "http://me.sichuanair.com/api/v1/plugins/LM_BASEINFO_LIST"
        emps = []
        payload, teams = None, None
        if team == 2:
            payload = {
                "baseCode": "CD01", "deptNo": "000009007005", "name": "",
                "page": "1", "rows": "1000", "teamNo": team['VALUE']
            }
            teams = self.getteams()
        if team == 4:
            payload = {
                "baseCode": "CD01", "deptNo": "000009007007", "name": "",
                "page": "1", "rows": "1000", "teamNo": team['VALUE']
            }
            teams = self.getteams(team=4)
        for team in teams:
            r = self.session.post(empsurl, data=payload)
            empsdict = json.loads(r.text)["data"]
            emps += [{"NAME": x['NAME'], "TEAM": team['TEXT'], 'MEMO1': x.get('MEMO')} for x in empsdict]
        return pandas.DataFrame(emps)

    def getempsx(self, team=2):
        empsurl = "http://me.sichuanair.com/api/v1/plugins/LM_BASEINFO_LIST"
        emps = []
        payload = None
        if team == 2:
            payload = {
                "baseCode": "CD01", "deptNo": "000009007005", "name": "",
                "page": "1", "rows": "1000", "teamNo": ""
            }
        if team == 4:
            payload = {
                "baseCode": "CD01", "deptNo": "000009007007", "name": "",
                "page": "1", "rows": "1000", "teamNo": ""
            }
        r = self.session.post(empsurl, data=payload)
        empsdict = json.loads(r.text)["data"]
        emps += [{'NAME': x['NAME'], 'MEMO1': x.get('MEMO')} for x in empsdict]
        return pandas.DataFrame(emps)

    def getattn1(self, team=2):
        """
        未打卡和休假的
        # type:attn/search
        # amrod定义的shift:D/N/Y/Z
        :return:
        """
        shifts = ['D', 'N', 'N', 'D']
        today = arrow.now().format("YYYY-MM-DD")
        payload = None
        if team == 2:
            days = arrow.get(today + 'T00:00:00.000+08:00') - arrow.get('2020-12-30T00:00:00.000+08:00')
            shift = shifts[days.days % 4]
            if days.days % 4 == 2:
                today = arrow.now().shift(days=-1).format("YYYY-MM-DD")
            if days.days % 4 == 3:
                today = arrow.now().shift(days=1).format("YYYY-MM-DD")
            payload = {
                "type": "attn", "attndate": today, "baseCode": "CD01", "deptNo": "000009007005", "shift": shift,
                "refPkid": ""
            }
        if team == 4:
            days = arrow.get(today + 'T00:00:00.000+08:00') - arrow.get('2020-12-28T00:00:00.000+08:00')
            shift = shifts[days.days % 4]
            if days.days % 4 == 2:
                today = arrow.now().shift(days=-1).format("YYYY-MM-DD")
            if days.days % 4 == 3:
                today = arrow.now().shift(days=1).format("YYYY-MM-DD")
            payload = {
                "type": "attn", "attndate": today, "baseCode": "CD01", "deptNo": "000009007007", "shift": shift,
                "refPkid": ""
            }
        r = self.session.post(self.attnUrl, data=payload)
        attn1 = json.loads(r.text).get("data", [])
        attn1 = pandas.DataFrame(attn1)
        attn1['attn'] = ''
        if 'LG' not in list(attn1):
            attn1['LG'] = ''
        attn1['attn'][attn1['LG'].isnull()] = 'N'
        attn1['label'] = ''
        attn1['label'][attn1['LG'].isnull()] = 'label-warning'
        return attn1

    def getattn2(self, team=2):
        """
        已正确打卡的
        :return:
        """
        shifts = ['D', 'N', 'N', 'D']
        today = arrow.now().format("YYYY-MM-DD")
        payload = None
        if team == 2:
            days = arrow.get(today + 'T00:00:00.000+08:00') - arrow.get('2020-12-30T00:00:00.000+08:00')
            shift = shifts[days.days % 4]
            if days.days % 4 == 2:
                today = arrow.now().shift(days=-1).format("YYYY-MM-DD")
            if days.days % 4 == 3:
                today = arrow.now().shift(days=1).format("YYYY-MM-DD")
            payload = {
                "type": "search", "attndate": today, "baseCode": "CD01", "deptNo": "000009007005", "shift": shift,
                "refPkid": "", "name": "", "name1": "", "empNo": ""
            }
        if team == 4:
            days = arrow.get(today + 'T00:00:00.000+08:00') - arrow.get('2020-12-28T00:00:00.000+08:00')
            shift = shifts[days.days % 4]
            if days.days % 4 == 2:
                today = arrow.now().shift(days=-1).format("YYYY-MM-DD")
            if days.days % 4 == 3:
                today = arrow.now().shift(days=1).format("YYYY-MM-DD")
            payload = {
                "type": "search", "attndate": today, "baseCode": "CD01", "deptNo": "000009007007", "shift": shift,
                "refPkid": "", "name": "", "name1": "", "empNo": ""
            }
        r = self.session.post(self.attnUrl, data=payload)
        attn2 = json.loads(r.text).get("data", [])
        attn2 = pandas.DataFrame(attn2)
        attn2['attn'] = 'Y'
        attn2['label'] = 'label-success'
        return attn2

    def getattn3(self, team=2):
        """
        错打卡的
        :return:
        """
        shifts = ['D', 'N', 'Y', 'Z']
        today = arrow.now().format("YYYY-MM-DD")
        attn3 = []
        payload = None
        if team == 2:
            days = arrow.get(today + 'T00:00:00.000+08:00') - arrow.get('2020-12-30T00:00:00.000+08:00')
            shifts.remove(['D', 'N', 'N', 'D'][days.days % 4])
            if days.days % 4 == 2:
                today = arrow.now().shift(days=-1).format("YYYY-MM-DD")
            if days.days % 4 == 3:
                today = arrow.now().shift(days=1).format("YYYY-MM-DD")
            for shift in shifts:
                payload = {
                    "type": "search", "attndate": today, "baseCode": "CD01", "deptNo": "000009007005", "shift": shift,
                    "refPkid": "", "name": "", "name1": "", "empNo": ""
                }
                r = self.session.post(self.attnUrl, data=payload)
                attn3 += json.loads(r.text).get("data", [])
        if team == 4:
            days = arrow.get(today + 'T00:00:00.000+08:00') - arrow.get('2020-12-28T00:00:00.000+08:00')
            shifts.remove(['D', 'N', 'N', 'D'][days.days % 4])
            if days.days % 4 == 2:
                today = arrow.now().shift(days=-1).format("YYYY-MM-DD")
            if days.days % 4 == 3:
                today = arrow.now().shift(days=1).format("YYYY-MM-DD")
            for shift in shifts:
                payload = {
                    "type": "search", "attndate": today, "baseCode": "CD01", "deptNo": "000009007007", "shift": shift,
                    "refPkid": "", "name": "", "name1": "", "empNo": ""
                }
                r = self.session.post(self.attnUrl, data=payload)
                attn3 += json.loads(r.text).get("data", [])
        if attn3:
            attn3 = pandas.DataFrame(attn3)
            attn3['attn'] = 'F'
            attn3['label'] = 'label-danger'
        else:
            attn3 = pandas.DataFrame(attn3)
        return attn3


if __name__ == "__main__":
    amro = AmroWeb()
    amro.login()
