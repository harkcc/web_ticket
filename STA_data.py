from pdb import run
import requests
import time
import login

_token = None
_token_expiry = None
TOKEN_VALIDITY_PERIOD = 3600  # token有效期（秒）

def get_token():
    """
    获取token，如果token不存在或已过期，则重新获取
    """
    global _token, _token_expiry
    current_time = time.time()
    
    if _token is None or _token_expiry is None or current_time >= _token_expiry:
        _token = login.run()
        _token_expiry = current_time + TOKEN_VALIDITY_PERIOD
    
    return _token

def request_sta_data(sid, inboundPlanId):
    headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'zh-CN,zh;q=0.9',
        'ak-client-type': 'web',
        'ak-origin': 'https://erp.lingxing.com',
        'auth-token': get_token(),  # 使用get_token获取token
        'content-type': 'application/json;charset=UTF-8',
        'origin': 'https://erp.lingxing.com',
        'priority': 'u=1, i',
        'referer': 'https://erp.lingxing.com/',
        'sec-ch-ua': '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'cross-site',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        'x-ak-company-id': '901217529031491584',
        'x-ak-env-key': 'SAAS-101',
        'x-ak-platform': '1',
        'x-ak-request-id': 'c0092d02-4b05-49c0-93c3-c3d6f99214d8',
        'x-ak-request-source': 'erp',
        'x-ak-uid': '10431785',
        'x-ak-version': '3.5.1.3.1.104',
        'x-ak-zid': '10330128',
    }

    json_data = {
        'inboundPlanId': inboundPlanId,
        'sid': sid,
        'req_time_sequence': '/amz-sta-server/inbound-shipment/shipmentTrackPage$$1',
    }

    response = requests.post(
        'https://gw.lingxingerp.com/amz-sta-server/inbound-shipment/shipmentTrackPage',
        headers=headers,
        json=json_data,
    )

    result = response.json()
    if result['code'] == 1 and result['data']:
        # 获取第一个箱子的地址信息
        address = result['data'][0]['address']
        return {
            'addressLine1': address.get('addressLine1', ''),
            'addressLine2': address.get('addressLine2', ''),
            'city': address.get('city', ''),
            'companyName': address.get('companyName', ''),
            'countryCode': address.get('countryCode', ''),
            'name': address.get('name', ''),
            'postalCode': address.get('postalCode', ''),
            'stateOrProvinceCode': address.get('stateOrProvinceCode', ''),
            'phoneNumber': address.get('phoneNumber', ''),
            'email': address.get('email', '')
        }
    return None

def request_loacl_localTaskId(ticket_id):
    headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'zh-CN,zh;q=0.9',
        'ak-client-type': 'web',
        'ak-origin': 'https://erp.lingxing.com',
        'auth-token': get_token(),  # 使用get_token获取token
        'content-type': 'application/json;charset=UTF-8',
        'origin': 'https://erp.lingxing.com',
        'priority': 'u=1, i',
        'referer': 'https://erp.lingxing.com/',
        'sec-ch-ua': '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'cross-site',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        'x-ak-company-id': '901217529031491584',
        'x-ak-env-key': 'SAAS-101',
        'x-ak-platform': '1',
        'x-ak-request-id': '9304cca2-04e5-4d22-a260-22b0f03a9851',
        'x-ak-request-source': 'erp',
        'x-ak-uid': '10431785',
        'x-ak-version': '3.5.1.3.1.104',
        'x-ak-zid': '10330128',
    }

    json_data = {
        'dateType': '1',
        'transparentPlanItem': '',
        'transparentPlanSign': '',
        'shipmentId': ticket_id,
        'sidList': [],
        'countryCodeList': [],
        'statusList': [],
        'current': 1,
        'size': 20,
        'req_time_sequence': '/amz-sta-server/inbound-plan/page$$5',
    }

    response = requests.post('https://gw.lingxingerp.com/amz-sta-server/inbound-plan/page', headers=headers,
                             json=json_data)
    
    result = response.json()
    if result['code'] == 1 and result['data']['records']:
        record = result['data']['records'][0]  # 获取第一条记录
        return {
            'sid': record['sid'],
            'inboundPlanId': record['inboundPlanId'],
            'sellerName': record['sellerName']
        }
    return None

def get_address_info(ticket_id):
    """
    获取ticket_id对应的地址信息
    
    Args:
        ticket_id: ticket ID
        
    Returns:
        dict: 包含地址信息和seller信息的字典，如果获取失败返回None
    """
    # 获取sid和inboundPlanId
    basic_info = request_loacl_localTaskId(ticket_id)
    if not basic_info:
        return None
        
    # 获取地址信息
    address_info = request_sta_data(basic_info['sid'], basic_info['inboundPlanId'])
    if not address_info:
        return None
        
    # 合并信息
    return {
        'seller_info': {
            'sellerName': basic_info['sellerName'],
            'sid': basic_info['sid'],
            'inboundPlanId': basic_info['inboundPlanId']
        },
        'address_info': address_info
    }