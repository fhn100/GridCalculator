import requests
import json

def parse_cookies(cookie_string):
    """安全地解析 cookie 字符串"""
    cookies = {}
    if cookie_string:
        try:
            # 处理 '; ' 或 ';' 分隔符
            cookie_parts = cookie_string.split(';')
            for part in cookie_parts:
                part = part.strip()
                if '=' in part:
                    key, value = part.split('=', 1)
                    cookies[key] = value
        except Exception as e:
            print(f"警告：解析 Cookie 时出错: {e}")
    return cookies

class APIClient:
    def __init__(self, user_id, fund_key, cookie, start_date, end_date, headers=None):
        self.user_id = user_id
        self.fund_key = fund_key
        self.cookie = cookie
        self.start_date = start_date
        self.end_date = end_date
        self.headers = headers or {
            'User-Agent': 'Mozilla/5.0',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            # 可以根据需要添加更多默认头
        }

    def _get_session(self):
        cookies = parse_cookies(self.cookie)
        session = requests.Session()
        session.cookies.update(cookies)
        return session

    def _send_request(self, url, data):
        session = self._get_session()
        try:
            response = session.post(url=url, data=data, headers=self.headers, timeout=30)
            response.raise_for_status() # 如果状态码不是 200，会抛出异常
            return response
        except requests.exceptions.RequestException as e:
            raise Exception(f"网络请求失败: {e}")

    def get_stock_history(self):
        url = "https://tzzb.10jqka.com.cn/caishen_httpserver/tzzb/caishen_fund/stock_position/v1/stock_history_query"
        data = {
            "userid": self.user_id,
            "fundkey": self.fund_key,
            "stock_code": "",
            "stock_account": "",
            "start_date": self.start_date,
            "end_date": self.end_date,
            "from_pc": "1"
        }
        response = self._send_request(url, data)
        return response

    def _get_stock_position(self):
        """
        获取股票持仓信息，用于获取股票名称。
        """
        url = "https://tzzb.10jqka.com.cn/caishen_httpserver/tzzb/caishen_fund/pc/asset/v1/stock_position"
        # 注意：请求体中同时包含了 userid 和 user_id，根据示例数据推测可能都需要。
        data = {
            "userid": self.user_id,
            "user_id": self.user_id,
            "fund_key": self.fund_key,
            "manual_id": "",
            "rzrq_fund_key": ""
        }
        response = self._send_request(url, data)
        return response