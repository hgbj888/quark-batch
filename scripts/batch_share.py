#!/usr/bin/env python3
"""
夸克网盘批量转存+分享工具
基于 quarkpan 库实现
"""

import argparse
import os
import sys
import re
from pathlib import Path

# 检查 quarkpan 是否安装
try:
    from quark_client import QuarkClient
except ImportError:
    print("错误: 请先安装 quarkpan 库: pip install quarkpan")
    sys.exit(1)

try:
    import pandas as pd
    import openpyxl
except ImportError:
    print("错误: 请先安装依赖: pip install openpyxl pandas")
    sys.exit(1)


def load_cookie_from_env():
    """从 .env 文件加载 Cookie"""
    env_path = Path(__file__).parent.parent / ".env"
    if not env_path.exists():
        return None
    
    with open(env_path, 'r') as f:
        for line in f:
            line = line.strip()
            if line.startswith('QUARK_COOKIE='):
                return line.split('=', 1)[1].strip()
    return None


def parse_share_url(url):
    """解析分享链接，提取 share_id 和提取码"""
    # 匹配 https://pan.quark.cn/s/xxxxx
    match = re.search(r'pan\.quark\.cn/s/([a-zA-Z0-9]+)', url)
    if not match:
        return None, None
    
    share_id = match.group(1)
    
    # 提取提取码
    pwd_match = re.search(r'(?:提取码|password|密码|码)[:：]?\s*([a-zA-Z0-9]{4,8})', url, re.IGNORECASE)
    password = pwd_match.group(1) if pwd_match else ''
    
    return share_id, password


def parse_input(input_data):
    """解析输入数据（自动去重）
    
    支持格式：
    1. 只有链接：https://pan.quark.cn/s/xxxx
    2. 名称+链接：XXX教程 https://pan.quark.cn/s/xxxx （空格或tab分隔）
    """
    links = []
    seen_urls = set()  # 用于去重
    MAX_URL_LENGTH = 200  # URL 最大长度限制
    
    # 判断是文件路径还是直接文本
    if Path(input_data).exists():
        file_path = Path(input_data).resolve()  # 解析绝对路径
        
        # 防止路径遍历：检查是否在允许的工作目录下
        allowed_dirs = [Path.cwd(), Path.home()]
        if not any(str(file_path).startswith(str(d)) for d in allowed_dirs):
            print(f"错误: 文件路径不允许超出工作目录")
            return []
        
        # 检查文件大小（最大 50MB）
        if file_path.stat().st_size > 50 * 1024 * 1024:
            print(f"错误: 文件过大（最大 50MB）")
            return []
        
        if file_path.suffix in ['.xlsx', '.xls']:
            # Excel 文件
            df = pd.read_excel(file_path)
            
            # 支持多种列名
            link_col = '链接' if '链接' in df.columns else '资源链接' if '资源链接' in df.columns else None
            name_col = '名称' if '名称' in df.columns else '资源名称' if '资源名称' in df.columns else None
            
            if not link_col:
                print("错误: Excel 文件中缺少'链接'列")
                return []
            
            for _, row in df.iterrows():
                name = row.get(name_col, '') if name_col else ''
                link = row.get(link_col, '')
                if link and isinstance(link, str):
                    link = link.strip()
                    if link and link not in seen_urls:
                        links.append({
                            'name': name.strip() if name else '',
                            'url': link
                        })
                        seen_urls.add(link)
        else:
            # 文本文件
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        parsed = _parse_line(line)
                        if parsed['url'] and parsed['url'] not in seen_urls:
                            links.append(parsed)
                            seen_urls.add(parsed['url'])
    else:
        # 直接文本
        for line in input_data.split('\n'):
            line = line.strip()
            if line:
                parsed = _parse_line(line)
                if parsed['url'] and parsed['url'] not in seen_urls:
                    links.append(parsed)
                    seen_urls.add(parsed['url'])
    
    return links


def _parse_line(line):
    """解析单行输入，提取名称和链接
    
    Returns:
        dict: {'name': '名称', 'url': '链接'}
    """
    MAX_URL_LENGTH = 200  # URL 最大长度限制
    
    # 匹配夸克网盘链接（带长度限制）
    url_match = re.search(r'https://pan\.quark\.cn/s/[a-zA-Z0-9]+', line)
    
    if url_match:
        url = url_match.group(0)
        
        # 长度检查
        if len(url) > MAX_URL_LENGTH:
            return {'name': '', 'url': ''}
        
        # 提取链接前的文本作为名称（限制名称长度）
        name_part = line[:url_match.start()].strip()
        # 去掉名称末尾的分隔符（空格、tab、-、:等）
        name_part = re.sub(r'[-:：\s\t]+$', '', name_part)
        # 限制名称最大长度
        name_part = name_part[:200] if name_part else ''
        
        return {'name': name_part, 'url': url}
    
    # 如果没有匹配到链接，返回空
    return {'name': '', 'url': ''}


def get_or_create_folder(client, folder_name="来自：分享"):
    """获取或创建指定名称的文件夹
    
    Args:
        client: QuarkClient 实例
        folder_name: 文件夹名称
        
    Returns:
        str: 文件夹 ID
    """
    # 先通过搜索查找是否存在
    try:
        search_result = client.files.search_files(keyword=folder_name, folder_id="0")
        if search_result and 'data' in search_result and 'list' in search_result['data']:
            for item in search_result['data']['list']:
                if item.get('file_name') == folder_name and item.get('file_type') == 0:
                    folder_id = item.get('fid')
                    print(f"找到已有文件夹: {folder_name} (ID: {folder_id})")
                    return folder_id
    except Exception as e:
        print(f"搜索文件夹时出错: {e}")
    
    # 不存在则创建
    try:
        result = client.create_folder(folder_name=folder_name, parent_id="0")
        if result.get('code') == 0:
            folder_id = result.get('data', {}).get('file_id', '')
            print(f"创建新文件夹: {folder_name} (ID: {folder_id})")
            return folder_id
        else:
            print(f"创建文件夹失败: {result.get('message')}")
            return "0"  # 失败则用根目录
    except Exception as e:
        print(f"创建文件夹时出错: {e}")
        return "0"


def process_share(client, share_url, target_folder_id="0", retry=2, delay=3):
    """处理单个分享链接：转存并创建分享
    
    Args:
        client: QuarkClient 实例
        share_url: 分享链接
        target_folder_id: 目标文件夹ID
        retry: 失败重试次数
        delay: 请求间隔（秒）
    """
    result = {
        'share_url': share_url,
        'file_id': '',
        'file_name': '',
        'new_share_url': '',
        'status': '等待中',
        'error': ''
    }
    
    for attempt in range(1, retry + 1):
        try:
            # 等待间隔（除第一次外）
            if attempt > 1:
                print(f"    第 {attempt} 次重试...")
                import time
                time.sleep(delay)
            
            # 使用 parse_and_save 转存分享
            save_result = client.shares.parse_and_save(
                share_url=share_url,
                target_folder_id=target_folder_id,
                save_all=True,
                wait_for_completion=True,
                timeout=60
            )
            
            if save_result.get('code') != 0:
                result['status'] = '失败'
                result['error'] = save_result.get('message', '转存失败')
                continue
            
            # 从返回结果中提取文件信息
            task_data = save_result.get('task_result', {}).get('data', {})
            save_as_data = task_data.get('save_as', {})
            
            # 获取转存后的文件ID
            file_ids = save_as_data.get('save_as_top_fids', [])
            if not file_ids:
                result['status'] = '失败'
                result['error'] = '无法获取转存后的文件ID'
                continue
            
            file_id = file_ids[0]
            result['file_id'] = file_id
            
            # 获取文件名
            share_info = save_result.get('share_info', {})
            files = share_info.get('files', [])
            if files:
                file_name = files[0].get('file_name', '')
                result['file_name'] = file_name
            else:
                # 从文件列表获取
                file_list = client.list_files(folder_id=target_folder_id, page=1, size=1)
                if file_list and 'data' in file_list and file_list['data'].get('list'):
                    file_name = file_list['data']['list'][0].get('file_name', '')
                    result['file_name'] = file_name
            
            # 创建分享
            share_result = client.shares.create_share(
                file_ids=[file_id],
                title=result['file_name'],
                expire_days=-1,  # 永久有效
                password=None  # 公开分享
            )
            
            # 从返回结果直接获取分享链接
            if share_result.get('share_url'):
                result['new_share_url'] = share_result['share_url']
                result['status'] = '成功'
                return result
            elif share_result.get('data') and share_result['data'].get('share_id'):
                share_id_new = share_result['data']['share_id']
                my_shares = client.shares.get_my_shares(page=1, size=50)
                
                if my_shares and 'data' in my_shares and 'list' in my_shares['data']:
                    for s in my_shares['data']['list']:
                        if s.get('share_id') == share_id_new:
                            result['new_share_url'] = s.get('share_url', '')
                            break
                
                result['status'] = '成功'
                return result
            else:
                result['status'] = '失败'
                result['error'] = '创建分享失败'
                continue
        
        except Exception as e:
            result['status'] = '失败'
            result['error'] = str(e)
            continue
    
    return result


def save_results(results, output_path):
    """保存结果到 Excel"""
    # 确保输出目录存在
    output_file = Path(output_path).resolve()
    
    # 防止路径遍历：检查是否在允许的工作目录下
    allowed_dirs = [Path.cwd(), Path.home()]
    if not any(str(output_file).startswith(str(d)) for d in allowed_dirs):
        print(f"错误: 输出路径不允许超出工作目录")
        return
    
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # 整理数据（统一4列：资源名称、网盘链接、状态、备注）
    data = []
    for r in results:
        status = r.get('status', '')
        
        # 状态为成功时显示链接，失败时留空
        if status == '成功':
            link = r.get('new_share_url', '')
        else:
            link = ''
        
        # 备注处理
        error = r.get('error', '')
        if status == '失败':
            if '转存' in error:
                remark = '转存失败'
            elif '分享' in error:
                remark = '分享失败'
            else:
                remark = error if error else '处理失败'
        else:
            remark = ''
        
        data.append({
            '资源名称': r.get('file_name', ''),
            '网盘链接': link,
            '状态': status,
            '备注': remark
        })
    
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    # 统计
    success = len([r for r in results if r.get('status') == '成功'])
    failed = len(results) - success
    
    print(f"\n处理完成!")
    print(f"成功: {success} 个, 失败: {failed} 个")
    print(f"结果已保存到: {output_file}")


def main():
    parser = argparse.ArgumentParser(description="夸克网盘批量转存+分享工具")
    parser.add_argument("--input", required=True, help="输入文件路径或链接文本")
    parser.add_argument("--output", default="outputs/tables/result.xlsx", help="输出Excel文件路径")
    parser.add_argument("--folder", default="0", help="转存目标文件夹ID")
    parser.add_argument("--cookie", help="夸克网盘 Cookie（可选）")
    parser.add_argument("--retry", type=int, default=2, help="失败重试次数（默认2次）")
    parser.add_argument("--delay", type=int, default=1, help="请求间隔秒数（默认1秒）")
    
    args = parser.parse_args()
    
    # 获取 Cookie
    cookie = args.cookie or load_cookie_from_env()
    if not cookie:
        print("错误: 请通过 --cookie 参数提供 Cookie 或创建 .env 文件")
        sys.exit(1)
    
    # 初始化客户端
    print("初始化夸克网盘客户端...")
    try:
        # 使用 Cookie 初始化客户端
        client = QuarkClient(cookies=cookie, auto_login=False)
        
        # 测试连接
        storage = client.get_storage_info()
        print(f"登录成功! 用户: {storage.get('data', {}).get('nick_name', '未知')}")
    except Exception as e:
        print(f"初始化失败: {e}")
        sys.exit(1)
    
    # 获取目标文件夹（固定：来自：分享）
    target_folder = get_or_create_folder(client, "来自：分享")
    
    # 解析输入
    print(f"解析输入: {args.input}")
    links = parse_input(args.input)
    print(f"找到 {len(links)} 个链接（已去重）")
    
    # 处理
    results = []
    for i, link in enumerate(links, 1):
        url = link['url']
        name = link.get('name', '')
        print(f"[{i}/{len(links)}] 处理: {name or url[:50]}...")
        
        result = process_share(client, url, target_folder, args.retry, args.delay)
        
        if result['status'] == '成功':
            print(f"  ✓ 成功: {result['new_share_url']}")
        else:
            print(f"  ✗ 失败: {result['error']}")
        
        results.append(result)
        
        # 处理间隔
        if i < len(links):
            import time
            time.sleep(args.delay)
    
    # 保存结果
    save_results(results, args.output)


if __name__ == "__main__":
    main()