import base64
import requests
from Crypto.Cipher import AES
from Crypto.Util.number import getPrime, isPrime
from Crypto.Util.strxor import strxor


# ACCOUNT = "13391234626bt"
# PWD = '0JUER4'

ACCOUNT = "baitai-350000"
PWD = 'Lx159357'

PROXIES = {
    'http': 'http://remake:123456@175.6.58.222:2021',
    'https': 'http://remake:123456@175.6.58.222:2021',
}

def get_login_secretkey():
    url = "https://gw.lingxingerp.com/newadmin/api/passport/getLoginSecretKey"

    payload = {}
    headers = {
        'authority': 'gw.lingxingerp.com',
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'auth-token': '',
        'cache-control': 'no-cache',
        'content-length': '0',
        'origin': 'https://erp.lingxing.com',
        'pragma': 'no-cache',
        'referer': 'https://erp.lingxing.com/',
        'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'cross-site',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
        'x-ak-company-id': '901217529031491584',
        'x-ak-request-source': 'erp',
    }

    response = requests.request("POST", url, headers=headers, data=payload).json()
    print(response)
    return response['data']


def utf8_parse(s):
    # 将字符串转换成字节流
    b = s.encode('utf-8')
    # 计算 WordArray 的长度（以 32 位整数为单位）
    length = (len(b) + 3) // 4
    # 初始化 WordArray 的内容为 0
    words = [0] * length
    # 将字节流中的每个字节存储到 WordArray 中
    for i in range(len(b)):
        word_index = i // 4
        byte_index = i % 4
        words[word_index] |= b[i] << (24 - byte_index * 8)
    # 创建 WordArray 对象并返回
    # print(words)
    return words


def get_key(old_key):
    key_list = utf8_parse(old_key)
    key_list_bate = [long_to_bytes(i) for i in key_list]
    # print(key_list_bate)
    new_key = bytearray()
    for i in key_list_bate:
        new_key.extend(i)
    # print(new_key)
    return new_key


def encrypt_aes(plaintext, key):
    cipher = AES.new(key.encode(), AES.MODE_ECB)
    padded_plaintext = plaintext + (AES.block_size - len(plaintext) % AES.block_size) * chr(AES.block_size - len(plaintext) % AES.block_size)
    ciphertext = cipher.encrypt(padded_plaintext.encode())
    return base64.b64encode(ciphertext).decode()


def login(pwd,secretId):
    headers = {
        'authority': 'gw.lingxingerp.com',
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'auth-token': '',
        'cache-control': 'no-cache',
        'content-type': 'application/json;charset=UTF-8',
        'origin': 'https://erp.lingxing.com',
        'pragma': 'no-cache',
        'referer': 'https://erp.lingxing.com/',
        'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'cross-site',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
        'x-ak-request-id': 'cf27b2fa-34b3-4955-aca4-8b79f8a8eff7',
        'x-ak-request-source': 'erp',
        'x-ak-version': '1.0.0.0.0.200',
        'x-ak-zid': '',
    }

    json_data = {
        'account': ACCOUNT,
        'pwd': pwd,
        'verify_code': '',
        'uuid': '254cd273-7e74-4f22-ba3e-ae199adbff19',
        'auto_login': 1,
        # 'randStr': '@ywL',
        # 'ticket': 't03hPBBq3jECZDNi0KLbw71udaHFI0LKOu8CMO75Jz4Dd16-8FlGsO6aDfXMLND3mixktpRm4DgpHAxeVIvvWY6dFcZycj0MRd_okhxHlslv_dFfPY6ShZD1w**',
        'sensorsAnonymousId': '18725ea3f9fd58-07aaf295e0ad0dc-26021151-1764000-18725ea3fa0804',
        'secretId': secretId,
    }

    response = requests.post('https://gw.lingxingerp.com/newadmin/api/passport/login', headers=headers, json=json_data).json()
    return response
    # Note: json_data will not be serialized by requests
    # exactly as it was in the original request.
    # data = '{"account":"13391234626bt","pwd":"7yyFicrO19pT3btePmIDCA==","verify_code":"","uuid":"254cd273-7e74-4f22-ba3e-ae199adbff19","auto_login":1,"randStr":"@ywL","ticket":"t03hPBBq3jECZDNi0KLbw71udaHFI0LKOu8CMO75Jz4Dd16-8FlGsO6aDfXMLND3mixktpRm4DgpHAxeVIvvWY6dFcZycj0MRd_okhxHlslv_dFfPY6ShZD1w**","sensorsAnonymousId":"18725ea3f9fd58-07aaf295e0ad0dc-26021151-1764000-18725ea3fa0804","secretId":"d860457fcea740018127082e56df1c4e"}'
    # response = requests.post('https://gw.lingxingerp.com/newadmin/api/passport/login', headers=headers, data=data)


def run():
    login_secretkey = get_login_secretkey()
    # print(login_secretkey)
    key = login_secretkey['secretKey']
    secretId = login_secretkey['secretId']
    # key = 'mVTZ8pXTQXsEBwFw'.encode('utf-8')
    data = PWD
    pwd = encrypt_aes(data,key)
    # print(pwd)
    res = login(pwd,secretId)

    # print(res)
    return res['token']


# if __name__ == '__main__':
#     print(run())