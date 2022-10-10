# -*- coding=utf-8 -*-
# 识别验证码
# 测试地址1：https://passport2.chaoxing.com/login?loginType=4&newversion=true&fid=3012&newversion=true&refer=http://i.mooc.chaoxing.com
# 测试地址2：http://authserver.scitc.com.cn/authserver/getBackPasswordByEmail.do?service=http%3A%2F%2Fehall.scitc.com.cn%2Flogin%3Fservice%3Dhttp%3A%2F%2Fehall.scitc.com.cn%2Fnew%2Findex.html

#缺点：不能运算加减等复杂的验证，只能依靠打码平台；只针对对数字和字母识别率高。

import ddddocr
import requests


def read_p():
    # urls = 'https://passport2.chaoxing.com/num/code?1665215905599'
    urls = 'http://authserver.scitc.com.cn/authserver/captcha.html?ts=1665222343967'
    headers = {
        'User-Agent':
        'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Mobile Safari/537.36'
    }
    response = requests.get(urls, headers)
    fname = str('pic') + '.png'
    with open(fname, 'wb') as f:
        f.write(response.content)
    ocr = ddddocr.DdddOcr(use_gpu=True, device_id=0)
    with open("D:\\Python3_virtualenv\\python_network\\code\\pic.png",
              'rb') as f:
        img = f.read()
    res = ocr.classification(img)
    print(res)


if __name__ == '__main__':
    a = read_p()