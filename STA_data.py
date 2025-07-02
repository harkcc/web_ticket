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

def request_sta_data(sid, inboundPlanId,ticket_id,shipping_type):
    """
    根据inboundPlanId的前缀选择不同的请求方式获取数据
    
    Args:
        sid: 卖家ID
        inboundPlanId: 入库计划ID
        
    Returns:    
        dict: 包含地址信息的字典，如果获取失败返回None
    """
    # 检查inboundPlanId的前缀
    if shipping_type == 'amz':
        print("处理亚马逊的请求，ticket: {}".format(ticket_id))
        return request_sta_data_amz(sid, inboundPlanId)
    else:
        print("处理多平台的请求，inboundPlanId: {}".format(ticket_id))
        return request_sta_data_multi(ticket_id)

def request_sta_data_amz(sid, inboundPlanId):
    """
    原始的请求方式，用于处理FBAk开头的inboundPlanId
    """
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
        # print("DEBUG: 原始数据 ->", result['data'][0])  # 添加调试信息
        # print("DEBUG: shipmentName ->", result['data'][0].get('shipmentName'))
        # print("DEBUG: amazonReferenceId ->", result['data'][0].get('amazonReferenceId'))
        
        shipmentName = result['data'][0].get('shipmentName','')
        amazonReferenceId = result['data'][0].get('amazonReferenceId','')

        return {
            'type':'amz',
            'addressLine1': address.get('addressLine1', ''),
            'addressLine2': address.get('addressLine2', ''),
            'city': address.get('city', ''),
            'companyName': address.get('companyName', ''),
            'countryCode': address.get('countryCode', ''),
            'name': address.get('name', ''),
            'postalCode': address.get('postalCode', ''),
            'stateOrProvinceCode': address.get('stateOrProvinceCode', ''),
            'phoneNumber': address.get('phoneNumber', ''),
            'email': address.get('email', ''),
            'shipmentName':shipmentName,
            'amazonReferenceId':amazonReferenceId,
        }

    return None

def request_sta_data_multi(ticket_id):
    """
    新的请求方式，用于处理非FBAk开头的inboundPlanId（主要是沃尔玛等多平台）
    
    Args:
        sid: 卖家ID
        inboundPlanId: 入库计划ID
        
    Returns:
        dict: 包含地址信息的字典，如果获取失败返回None
    """
    try:
        print(f"处理多平台的请求，inboundPlanId: {ticket_id}")

        # 第一步：通过inboundPlanId查询获取详细信息的id
        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'zh-CN,zh;q=0.9',
            'ak-client-type': 'web',
            'ak-origin': 'https://erp.lingxing.com',
            'auth-token': get_token(),
            'content-type': 'application/json;charset=UTF-8',
            'origin': 'https://erp.lingxing.com',
            'priority': 'u=1, i',
            'referer': 'https://erp.lingxing.com/',
            'sec-ch-ua': '"Google Chrome";v="137", "Chromium";v="137", "Not/A)Brand";v="24"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"macOS"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36',
            'x-ak-company-id': '901217529031491584',
            'x-ak-env-key': 'SAAS-101',
            'x-ak-language': 'zh',
            'x-ak-platform': '2',
            'x-ak-request-source': 'erp',
            'x-ak-uid': '10431785',
            'x-ak-version': '3.6.3.3.0.075',
            'x-ak-zid': '10330128',
        }

        json_data = {
            'queryValues': [],
            'queryType': '5',
            'queryValue': ticket_id,  
            'storeIdList': [],
            'timeType': '1',
            'offset': 0,
            'length': 20,
            'req_time_sequence': '/mp-platform-warehouse-api/api/cargo/list$$3',
        }

        print(f"请求参数: {json_data}")
        response = requests.post('https://gw.lingxingerp.com/mp-platform-warehouse-api/api/cargo/list', headers=headers, json=json_data)
        print(f"响应状态码: {response.status_code}")
        print(f"响应内容前200字符: {response.text[:200]}")
        id_json = response.json()
        
        # 检查响应并提取ID
        if id_json['code'] != 1 or 'data' not in id_json or 'list' not in id_json['data'] or not id_json['data']['list']:
            print(f"获取ID失败，响应: {id_json}")
            return None
            
        cargo_id = id_json['data']['list'][0]['id']
        
        # 第二步：使用获取到的ID查询详细信息
        detail_headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'zh-CN,zh;q=0.9',
            'ak-client-type': 'web',
            'ak-origin': 'https://erp.lingxing.com',
            'auth-token': get_token(),
            'content-type': 'application/json;charset=UTF-8',
            'origin': 'https://erp.lingxing.com',
            'priority': 'u=1, i',
            'referer': 'https://erp.lingxing.com/',
            'sec-ch-ua': '"Google Chrome";v="137", "Chromium";v="137", "Not/A)Brand";v="24"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"macOS"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36',
            'x-ak-company-id': '901217529031491584',
            'x-ak-env-key': 'SAAS-101',
            'x-ak-language': 'zh',
            'x-ak-platform': '2',
            'x-ak-request-source': 'erp',
            'x-ak-uid': '10431785',
            'x-ak-version': '3.6.3.3.0.075',
            'x-ak-zid': '10330128',
        }

        detail_json_data = {
            'id': cargo_id,
            'req_time_sequence': '/mp-platform-warehouse-api/api/cargo/detail$$1',
        }

        detail_response = requests.post(
            'https://gw.lingxingerp.com/mp-platform-warehouse-api/api/cargo/detail',
            headers=detail_headers,
            json=detail_json_data,
        )
        detail_json = detail_response.json()
        
        if detail_json['code'] != 1 or 'data' not in detail_json:
            print(f"获取详细信息失败，响应: {detail_json}")
            return None
            
        # 获取shippingAddress信息
        shipping_address = detail_json['data'].get('shippingAddress', {})
        
        # 按照要求组合地址信息
        street_detail = shipping_address.get('streetDetail', '')
        city = shipping_address.get('city', '')
        province = shipping_address.get('province', '')
        country_name = shipping_address.get('receiveOrDeliverCountryName', '')
        postal_code = shipping_address.get('postalCode', '')
        countryCode = shipping_address.get('receiveOrDeliverCountry', '')
        
        # 组合地址行 - 按照要求的格式: streetDetail + city，province + receiveOrDeliverCountry + postalCode
        address_line2 = street_detail+','+city + ',' + province + ',' + countryCode + ','+ postal_code
 
        
        # 构建与request_sta_data_amz函数返回格式一致的结果
        return {
            'type':'multi',
            'addressLine1': street_detail,
            'addressLine2': address_line2,
            'city': city,
            'companyName': country_name,
            'countryCode': countryCode,
            'name': detail_json['data'].get('logisticsCode', ''),
            'postalCode': postal_code,
            'stateOrProvinceCode': province,
            'phoneNumber': detail_json['data'].get('phoneNumber', ''),
            'email': detail_json['data'].get('email', ''),
            'shipmentName': ticket_id,  
            'amazonReferenceId': detail_json['data'].get('amazonReferenceId', '')
        }
    
    except Exception as e:
        print(f"处理多平台请求时发生错误: {str(e)}")
        import traceback
        traceback.print_exc()
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
            'sellerName': record['sellerName'],          
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
    # 判断订单类型
    shipping_type = 'amz' if ticket_id.startswith('FBA') else 'multi'
    print(f"处理订单 ticket_id: {ticket_id}, shipping_type: {shipping_type}")
    
    # 对于多平台订单，直接处理
    if shipping_type == 'multi':
        print(f"多平台订单，直接获取地址信息")
        address_info = request_sta_data_multi(ticket_id)
        if not address_info:
            print(f"无法获取多平台地址信息，ticket_id: {ticket_id}")
            return None
            
        # 构建返回结构
        return {
            'seller_info': {
                'sellerName': address_info.get('companyName', ''),
                'country_name': address_info.get('countryCode', ''),
                'sid': '0',
                'inboundPlanId': ticket_id,
            },
            'address_info': address_info
        }
    
    # Amazon订单处理流程
    basic_info = request_loacl_localTaskId(ticket_id)
    if not basic_info:
        print(f"无法获取亚马逊基本信息，ticket_id: {ticket_id}")
        return None
        
    # 获取亚马逊订单地址信息
    address_info = request_sta_data_amz(basic_info['sid'], basic_info['inboundPlanId'])
    if not address_info:
        print(f"无法获取亚马逊地址信息，ticket_id: {ticket_id}")
        return None

    # 构建亚马逊订单返回结构
    country_dict = {
        "AC-BR": "巴西",
        "AC-CA": "加拿大",
        "AC-MX": "墨西哥",
        "AC-US": "美国",
        "CACNZ": "美国/加拿大",
        "HKCN": "香港",
        "CNHK": "中国香港",
        "CNJP": "日本",
        "FR": "法国",
        "DE": "德国",
        "IN": "印度",
        "IT": "意大利",
        "JP": "日本",
        "NL": "荷兰",
        "PL": "波兰",
        "SG": "新加坡",
        "ES": "西班牙",
        "SE": "瑞典",
        "UK": "英国",
        "GB": "英国",
        "UAE": "阿联酋",
        "AE": "阿联酋",
        "AU": "澳大利亚",
        "US": "美国",
        "CA": "加拿大",
        "MX": "墨西哥",
        "BR": "巴西",
        "SG": "新加坡"
    }

    return {
        'seller_info': {
            'sellerName': basic_info['sellerName'],
            'country_name': country_dict.get(basic_info['sellerName'], ''),
            'sid': basic_info['sid'],
            'inboundPlanId': basic_info['inboundPlanId'],
            # 'amazonReferenceId':basic_info['amazonReferenceId'],
        },
        'address_info': address_info
    }
