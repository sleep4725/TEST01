# <문화재 위치 정보>
import requests as req
from urllib.parse import urlencode
import xml.etree.ElementTree as Ett
# ==================================
class CHA:
    def __init__(self):
        self.target_url = "http://www.gis-heritage.go.kr/openapi/xmlService/spca.do"
        self.params = {
            "ccbaMnm1":"숭례문"
        }

    # Instance method (1)
    def reqURL(self):
        params = urlencode(self.params)
        url = self.target_url + "?" + params

        html = req.get(url)
        if html.status_code == 200:
            with open("cha_01.xml", "w", encoding='utf-8') as f:
                f.write(html.text)
                f.close()

def main():
    node = CHA() # 객체 생성
    # node.reqURL()
    node.xmlRead()
if __name__ == "__main__":
    main()

