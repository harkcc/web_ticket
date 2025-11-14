#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试ERP数据获取和转换（不连接MongoDB）
"""

import requests
import login
import time
import json


def get_product_list(token, limit=5):
    """获取产品列表（筛选已配对的产品）"""
    headers = {
        'AK-Client-Type': 'web',
        'AK-Origin': 'https://erp.lingxing.com',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Content-Type': 'application/json;charset=UTF-8',
        'X-AK-Company-Id': '901217529031491584',
        'X-AK-ENV-KEY': 'SAAS-101',
        'X-AK-PLATFORM': '1',
        'X-AK-Zid': '10330128',
        'auth-token': token,
    }
    
    json_data = {
        'search_field_time': 'create_time',
        'sort_field': 'create_time',
        'sort_type': 'desc',
        'search_field': 'sku',
        'status': [1],  # 在售
        'is_matched_alibaba': '1',  # 已配对
        'offset': 0,
        'length': limit,
        'product_type': [1, 2],
        'req_time_sequence': '/api/product/lists$$17',
    }
    
    print(f"正在获取产品列表（前{limit}个）...")
    
    response = requests.post(
        'https://erp.lingxing.com/api/product/lists',
        headers=headers,
        json=json_data
    )
    
    if response.status_code == 200:
        data = response.json()
        if data.get('code') == 1 and 'list' in data:
            products = data['list']
            print(f"✓ 成功获取 {len(products)} 条产品")
            return products
        else:
            print(f"❌ 业务错误: {data.get('msg')}")
            return []
    else:
        print(f"❌ HTTP请求失败: {response.status_code}")
        return []


def get_product_detail(token, product_id):
    """获取产品详情"""
    headers = {
        'AK-Client-Type': 'web',
        'Accept': 'application/json, text/plain, */*',
        'X-AK-Company-Id': '901217529031491584',
        'X-AK-ENV-KEY': 'SAAS-101',
        'X-AK-PLATFORM': '1',
        'X-AK-Zid': '10330128',
        'auth-token': token,
    }
    
    response = requests.get(
        f'https://erp.lingxing.com/api/product/info?id={product_id}',
        headers=headers
    )
    
    if response.status_code == 200:
        return response.json()
    return None


def get_product_links(token, product_id):
    """获取产品链接（MSKU列表）"""
    headers = {
        'AK-Client-Type': 'web',
        'Accept': 'application/json, text/plain, */*',
        'X-AK-Company-Id': '901217529031491584',
        'X-AK-ENV-KEY': 'SAAS-101',
        'X-AK-PLATFORM': '1',
        'X-AK-Zid': '10330128',
        'auth-token': token,
    }
    
    response = requests.get(
        f'https://erp.lingxing.com/api/module/product/product.view/getProductListing?product_id={product_id}',
        headers=headers
    )
    
    if response.status_code == 200:
        return response.json()
    return None


def split_field(field_value, en_field_value=None):
    """
    拆分中英文字段，支持两种情况：
    1. 分开的字段：customs_clearance_material + customs_clearance_en_material
    2. 合并的字段：customs_clearance_material = "织布/Woven"
    
    Args:
        field_value: 中文字段值（可能包含英文）
        en_field_value: 英文字段值（如果分开存储）
    
    Returns:
        tuple: (中文, 英文)
    """
    # 情况1：如果有单独的英文字段
    if en_field_value and en_field_value.strip():
        zh = field_value.strip() if field_value else ""
        en = en_field_value.strip()
        return zh, en
    
    # 情况2：如果字段为空
    if not field_value or field_value.strip() == "":
        return "", ""
    
    # 情况3：如果包含斜杠分隔符（注意：JSON中的\/会被解析为/）
    if '/' in field_value:
        parts = field_value.split('/', 1)
        zh = parts[0].strip() if len(parts) > 0 else ""
        en = parts[1].strip() if len(parts) > 1 else ""
        return zh, en
    
    # 情况4：只有中文，没有英文
    return field_value.strip(), ""


def parse_special_attr(special_attr):
    """解析special_attr字段"""
    if not special_attr or len(special_attr) == 0:
        return "否", "否"
    
    attr_list = [str(x) for x in special_attr]
    has_electric = "是" if ("1" in attr_list or "2" in attr_list) else "否"
    has_magnetic = "是" if "6" in attr_list else "否"
    
    return has_electric, has_magnetic


def transform_product_data(product_info, link_info):
    """将ERP产品数据转换为MongoDB格式"""
    info = product_info.get('info', {})
    declaration = info.get('product_declaration_list', {})
    clearance = info.get('product_clearance_list', {})
    special_attr = info.get('special_attr', [])
    spec_info = info.get('spec_info', {})  # 规格信息
    
    # 解析电磁属性
    electrified, magnetic = parse_special_attr(special_attr)
    
    # 解析材质（支持两种格式）
    material_zh, material_en = split_field(
        clearance.get('customs_clearance_material', ''),
        clearance.get('customs_clearance_en_material', '')  # 可能有单独的英文字段
    )
    
    # 解析用途（支持两种格式）
    use_zh, use_en = split_field(
        clearance.get('customs_clearance_usage', ''),
        clearance.get('customs_clearance_en_usage', '')  # 可能有单独的英文字段
    )
    
    # 获取重量（从spec_info中获取净重）
    weight_str = spec_info.get('cg_product_net_weight', '0')
    try:
        weight = float(weight_str)
    except:
        weight = 0.0
    
    # 构建MongoDB文档
    document = {
        "msku": link_info.get('msku', ''),
        "productNameZh": declaration.get('customs_export_name', ''),
        "productNameEn": declaration.get('customs_import_name', ''),
        "price": declaration.get('customs_import_price', ''),
        "brand": info.get('brand_name', '') or "无",
        "model": info.get('model', '') or "无",
        "HS": declaration.get('customs_declaration_hs_code', ''),
        "image_url": "",
        "asin": link_info.get('asin', ''),
        "askPrice": "",
        "electrified": electrified,
        "magnetic": magnetic,
        "materialEn": material_en,
        "materialZh": material_zh,
        "outboundFee": "",
        "productLink": link_info.get('asin_url', ''),
        "putAwayFee": "",
        "useEn": use_en,
        "useZh": use_zh,
        "weight": weight,
        "X_ROW_K": "",
        "created_at": ""
    }
    
    return document


def main():
    """主函数"""
    print("=" * 70)
    print("ERP产品数据获取和转换测试")
    print("=" * 70)
    
    # 1. 登录
    print("\n[步骤1] 登录ERP系统...")
    token = login.run()
    print(f"✓ 登录成功")
    
    # 2. 获取产品列表
    print("\n[步骤2] 获取产品列表...")
    products = get_product_list(token, limit=3)  # 只获取3个产品测试
    
    if not products:
        print("❌ 没有获取到产品，退出")
        return
    
    # 3. 处理每个产品
    print(f"\n[步骤3] 处理产品数据...")
    all_documents = []
    
    for idx, product in enumerate(products, 1):
        product_id = product.get('id')
        sku = product.get('sku', 'Unknown')
        
        print(f"\n--- [{idx}/{len(products)}] 处理产品: {sku} (ID: {product_id}) ---")
        
        try:
            # 获取产品详情
            print(f"  → 获取产品详情...")
            detail = get_product_detail(token, product_id)
            if not detail or detail.get('code') != 1:
                print(f"  ❌ 获取产品详情失败")
                continue
            print(f"  ✓ 产品详情获取成功")
            
            # 获取产品链接
            print(f"  → 获取产品链接...")
            links = get_product_links(token, product_id)
            if not links or links.get('code') != 1:
                print(f"  ❌ 获取产品链接失败")
                continue
            
            link_data = links.get('data', [])
            if not link_data:
                print(f"  ⚠️  该产品没有MSKU链接")
                continue
            
            print(f"  ✓ 找到 {len(link_data)} 个MSKU")
            
            # 转换每个MSKU
            for link in link_data:
                msku = link.get('msku', '')
                print(f"    → 转换MSKU: {msku}")
                
                document = transform_product_data(detail, link)
                all_documents.append(document)
                print(f"    ✓ 转换完成")
            
            time.sleep(0.5)  # 避免请求过快
            
        except Exception as e:
            print(f"  ❌ 处理出错: {str(e)}")
            import traceback
            traceback.print_exc()
    
    # 4. 显示结果
    print("\n" + "=" * 70)
    print(f"转换完成！共生成 {len(all_documents)} 条记录")
    print("=" * 70)
    
    if all_documents:
        print("\n[示例] 第一条记录：")
        print(json.dumps(all_documents[0], indent=2, ensure_ascii=False))
        
        # 保存到JSON文件
        output_file = "erp_product_data_test.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(all_documents, f, indent=2, ensure_ascii=False)
        print(f"\n✓ 所有数据已保存到: {output_file}")
        
        # 显示字段统计
        print("\n[字段统计]")
        for key in all_documents[0].keys():
            non_empty = sum(1 for doc in all_documents if doc.get(key))
            print(f"  {key:20s}: {non_empty}/{len(all_documents)} 条有值")
    
    print("\n✓ 测试完成！")


if __name__ == "__main__":
    main()
