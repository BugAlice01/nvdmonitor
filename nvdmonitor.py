import requests
from bs4 import BeautifulSoup
import time
from datetime import datetime, timedelta, timezone
from dateutil import parser
import xlwt
import xlrd
from xlutils.copy import copy
import yaml
from typing import List, Dict, Optional
import os
import json

# 文件夹路径常量
RESULT_FOLDER = "result"
JSON_FOLDER = "json"
# 颜色常量定义
COLOR_RED = "\033[1;31m"
COLOR_GREEN = "\033[1;32m"
COLOR_YELLOW = "\033[1;33m"
COLOR_BLUE = "\033[1;34m"
COLOR_CYAN = "\033[1;36m"
COLOR_RESET = "\033[0m"

# 常量定义
BASE_URL = "https://nvd.nist.gov"
DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'en-GB,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
}
MAX_RETRIES = 3
RETRY_DELAY = 5  # 秒



def log_message(message: str, log_file: str = 'log.txt', level: str = "info") -> None:
    """记录日志信息，支持不同级别"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    level_info = {
        "error": (COLOR_RED, "[ERROR]"),
        "warning": (COLOR_YELLOW, "[WARN]"),
        "success": (COLOR_GREEN, "[SUCCESS]"),
        "info": (COLOR_CYAN, "[INFO]"),
    }.get(level.lower(), (COLOR_RESET, "[LOG]"))

    color, prefix = level_info
    log_entry = f"{color}[{timestamp}] {prefix} {message}{COLOR_RESET}"
    print(log_entry)


def get_excel_filename(target_name: str, target_date: datetime) -> str:
    """获取Excel文件名，使用目标日期而不是当前日期"""
    return os.path.join(RESULT_FOLDER, f"{target_name}漏洞监控_{target_date.strftime('%Y%m%d')}.xls")


def get_json_filename(target_name: str, target_date: datetime) -> str:
    """获取JSON文件名"""
    return os.path.join(JSON_FOLDER, f"{target_name}_vulns_{target_date.strftime('%Y%m%d')}.json")


def load_existing_vulnerabilities(target_name: str, target_date: datetime) -> Dict[str, Dict]:
    """从目标日期的历史文件(JSON和Excel)加载已存在的漏洞"""
    existing_vulns = {}

    # 构造目标日期的文件名模式
    date_str = target_date.strftime('%Y%m%d')
    json_filename = f"{target_name}_vulns_{date_str}.json"
    excel_filename = f"{target_name}漏洞监控_{date_str}.xls"

    # 从JSON文件加载
    json_path = os.path.join(JSON_FOLDER, json_filename)
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                vulns = json.load(f)
                if isinstance(vulns, dict):
                    existing_vulns.update(vulns)
        except Exception as e:
            log_message(f"读取JSON文件{json_filename}失败: {e}", level="warning")

    # 从Excel文件加载
    excel_path = os.path.join(RESULT_FOLDER, excel_filename)
    if os.path.exists(excel_path):
        try:
            workbook = xlrd.open_workbook(excel_path)
            sheet = workbook.sheet_by_index(0)

            current_vuln = {}
            for row_idx in range(sheet.nrows):
                row = sheet.row_values(row_idx)
                if not row:
                    continue

                cell_value = str(row[0]).strip()
                if cell_value.startswith("对于") and "任务" in cell_value:
                    # 新漏洞开始
                    current_vuln = {}
                elif cell_value.startswith("链接:"):
                    vuln_link = cell_value.split(":")[1].strip()
                    current_vuln['link'] = vuln_link
                    vuln_id = vuln_link.split('/')[-1]
                    current_vuln['id'] = vuln_id
                elif cell_value.startswith("CVSS评分:"):
                    current_vuln['cvss'] = cell_value.split(":")[1].strip()
                elif cell_value.startswith("摘要:"):
                    current_vuln['summary'] = cell_value.split(":")[1].strip()
                    if current_vuln.get('id'):
                        existing_vulns[current_vuln['id']] = current_vuln.copy()
        except Exception as e:
            log_message(f"读取Excel文件{excel_filename}失败: {e}", level="warning")

    return existing_vulns


def filter_new_vulnerabilities(vulnerabilities: List[Dict], existing_vulns: Dict[str, Dict]) -> List[Dict]:
    """过滤掉已经存在的漏洞"""
    return [vuln for vuln in vulnerabilities if vuln['id'] not in existing_vulns]


def write_to_excel(target_name: str, vulnerabilities: List[Dict], existing_vulns: Dict[str, Dict],
                   target_date: datetime) -> bool:
    """将漏洞信息写入Excel文件"""
    try:
        filename = get_excel_filename(target_name, target_date)

        # 初始化Excel工作簿和工作表
        if os.path.exists(filename):
            # 如果文件已存在，打开并复制
            workbook = xlrd.open_workbook(filename)
            wb = copy(workbook)
            sheet = wb.get_sheet(0)  # 获取第一个工作表
        else:
            # 如果文件不存在，创建新的
            wb = xlwt.Workbook(encoding='utf-8')
            sheet = wb.add_sheet('漏洞列表')  # 创建工作表
            # 初始化行索引
            row = 0

        # 设置样式
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = '微软雅黑'
        font.height = 220  # 11pt
        style.font = font

        # 只写入新漏洞
        new_vulns = filter_new_vulnerabilities(vulnerabilities, existing_vulns)
        for vuln in new_vulns:
            # 解析发布时间并转换为中国时区
            published_time = parser.parse(vuln['published'])
            china_time = published_time.astimezone(timezone(timedelta(hours=8)))
            china_time_str = china_time.strftime('%Y年%m月%d日 %H:%M:%S')

            sheet.write(row, 0, f"对于{target_name}任务，最新的cve漏洞公开时间是: {china_time_str}", style)
            sheet.write(row + 1, 0, f"链接: {vuln['link']}", style)
            sheet.write(row + 2, 0, f"CVSS评分: {vuln['cvss']}", style)
            sheet.write(row + 3, 0, f"摘要: {vuln['summary']}", style)
            sheet.write(row + 4, 0, "-" * 80, style)
            row += 5

        wb.save(filename)
        log_message(f"结果已写入Excel文件: {filename} (新增{len(new_vulns)}个漏洞)")
        return True
    except Exception as e:
        log_message(f"写入Excel文件失败: {e}")
        return False


def load_config() -> Dict:
    """加载并验证配置文件"""
    try:
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f) or {}

            # 验证必填字段
            for field in ['target', 'querry']:
                if not config.get(field):
                    raise ValueError(f"config.yaml必须包含'{field}'字段")

            # 设置默认值
            config['day_ago'] = int(config.get('day_ago', 0))
            config['page'] = config.get('page', '1')
            config['wechat_bot'] = config.get('wechat_bot', {})

            # 解析页面范围
            if '-' in config['page']:
                start, end = map(int, config['page'].split('-'))
                if start < 1 or end < start:
                    raise ValueError("页面范围无效，必须满足1 ≤ 开始页 ≤ 结束页")
                config['start_page'], config['end_page'] = start, end
            else:
                page = int(config['page'])
                if page < 1:
                    raise ValueError("页面必须大于等于1")
                config['start_page'] = config['end_page'] = page

            return config

    except FileNotFoundError:
        raise FileNotFoundError("未找到config.yaml文件")
    except ValueError as e:
        raise ValueError(f"配置文件错误: {e}")
    except Exception as e:
        raise Exception(f"读取配置文件时出错: {e}")


def make_request(url: str, headers: Dict, retries: int = MAX_RETRIES) -> Optional[requests.Response]:
    """带有重试机制的请求函数"""
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=15)
            response.raise_for_status()
            return response
        except requests.exceptions.RequestException as e:
            if attempt < retries - 1:
                log_message(f"请求失败 (尝试 {attempt + 1}/{retries}): {e}, {retries}秒后重试...")
                time.sleep(RETRY_DELAY)
            else:
                raise
    return None


def parse_vulnerability_row(row) -> Optional[Dict]:
    """解析单个漏洞行"""
    try:
        published_tag = row.find('span', attrs={'data-testid': lambda x: x and x.startswith('vuln-published-on-')})
        vuln_id_tag = row.find('a', attrs={'data-testid': lambda x: x and x.startswith('vuln-detail-link-')})
        if not published_tag or not vuln_id_tag:
            return None

        published_date = parser.parse(published_tag.text.strip())
        summary_tag = row.find('p', attrs={'data-testid': lambda x: x and x.startswith('vuln-summary-')})

        # 修改CVSS评分提取逻辑
        cvss_info = ""
        # 检查CVSS v3.x
        cvss3_tag = row.find('a', attrs={'data-testid': lambda x: x and x.startswith('vuln-cvss3-link-')})
        if cvss3_tag:
            cvss_score = cvss3_tag.text.strip()
            severity_class = cvss3_tag.get('class', '')
            if 'label-danger' in severity_class:
                severity = "HIGH"
            elif 'label-warning' in severity_class:
                severity = "MEDIUM"
            elif 'label-low' in severity_class:
                severity = "LOW"
            else:
                severity = ""
            cvss_info = f"V3.x: {cvss_score} {severity}"
        else:
            # 检查CVSS v2.0
            cvss2_tag = row.find('span', attrs={'data-testid': lambda x: x and x.startswith('vuln-cvss2-na-')})
            if cvss2_tag and "(not available)" not in cvss2_tag.text:
                cvss_info = "V2.0: " + cvss2_tag.text.split(":")[1].strip()

        return {
            'id': vuln_id_tag.text.strip(),
            'link': f"{BASE_URL}{vuln_id_tag['href']}",
            'published': published_date.strftime("%B %d, %Y; %I:%M:%S %p %z"),
            'cvss': cvss_info if cvss_info else "N/A",
            'summary': summary_tag.text.strip() if summary_tag else "N/A"
        }
    except Exception as e:
        log_message(f"解析漏洞行时出错: {e}")
        return None


def send_wechat_notification(webhook_url: str, target_name: str, vulnerabilities: List[Dict], day_ago: int) -> bool:
    """发送企业微信通知"""
    if not webhook_url:
        return False

    try:
        day = datetime.now() - timedelta(days=day_ago)
        content = [
            f"**{target_name}NVD漏洞监控报告**",
            f"扫描时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}",
            f"在{day.year}年{day.month}月{day.day}日后共发现 {len(vulnerabilities)} 个新漏洞:\n"
        ]

        for idx, vuln in enumerate(vulnerabilities, 1):
            content.extend([
                f"{idx}. **{vuln['id']}** (CVSS: {vuln['cvss']})",
                f"发布时间: {vuln['published']}",
                f"[查看详情]({vuln['link']})\n"
            ])

        response = requests.post(
            webhook_url,
            json={
                "msgtype": "markdown",
                "markdown": {"content": "\n".join(content)}
            },
            timeout=10
        )

        if response.status_code == 200:
            log_message("企业微信机器人推送成功!")
            return True
        log_message(f"企业微信机器人推送失败: {response.text}")
        return False
    except Exception as e:
        log_message(f"企业微信机器人推送异常: {e}")
        return False


def save_vulnerabilities_to_json(target_name: str, vulnerabilities: List[Dict], target_date: datetime) -> None:
    """将漏洞信息保存为JSON文件"""
    try:
        filename = get_json_filename(target_name, target_date)
        vulns_dict = {vuln['id']: vuln for vuln in vulnerabilities}

        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(vulns_dict, f, ensure_ascii=False, indent=2)
        log_message(f"漏洞数据已保存为JSON文件: {filename}")
    except Exception as e:
        log_message(f"保存JSON文件失败: {e}", level="warning")


def main():
    try:
        config = load_config()
        target_name = config['target']
        query_param = config['querry']
        day_ago = config['day_ago']
        start_page = config['start_page']
        end_page = config['end_page']
        webhook_url = config['wechat_bot'].get('webhook_url', '')

        start_time = datetime.now()
        log_message(
            f"开始扫描: 目标={target_name}, 查询参数={query_param}, 天数={day_ago}, 页面范围={start_page}-{end_page}")

        day = datetime.now() - timedelta(days=day_ago)
        target_date = parser.parse(day.strftime("%B %d, %Y; 0:00:00 am +0800"))
        all_vulns = []

        # 修改这里：传递target_date参数
        existing_vulns = load_existing_vulnerabilities(target_name, target_date)

        for page in range(start_page, end_page + 1):
            start_index = (page - 1) * 20
            url = f"{BASE_URL}/vuln/search/results?query={query_param}&startIndex={start_index}"
            log_message(f"正在扫描第 {page} 页 (startIndex={start_index})...")
            time.sleep(1)

            try:
                response = make_request(url, DEFAULT_HEADERS)
                if not response:
                    continue

                soup = BeautifulSoup(response.text, 'html.parser')
                vuln_rows = soup.find_all('tr', attrs={'data-testid': lambda x: x and x.startswith('vuln-row-')})

                if not vuln_rows:
                    log_message(f"第 {page} 页未找到任何漏洞数据")
                    continue

                for row in vuln_rows:
                    vuln_data = parse_vulnerability_row(row)
                    if vuln_data and parser.parse(vuln_data['published']) > target_date:
                        all_vulns.append(vuln_data)

            except Exception as e:
                log_message(f"处理第 {page} 页时出错: {e}")
                continue

        # 过滤掉已经存在的漏洞
        new_vulns = filter_new_vulnerabilities(all_vulns, existing_vulns)

        chinese_target_date = f"{day.year}年{day.month}月{day.day}日"
        if all_vulns:
            log_message(
                f"\n在{chinese_target_date}后发布的漏洞共有 {len(all_vulns)} 个，其中 {len(new_vulns)} 个是新发现的:",
                level="success")
            if new_vulns:
                for vuln in new_vulns:
                    log_message(f"""
        {COLOR_BLUE}漏洞ID: {vuln['id']}
        链接: {vuln['link']}
        发布日期: {vuln['published']}
        CVSS评分: {vuln['cvss']}
        摘要: {vuln['summary'][:500]}...
        {"-" * 80}{COLOR_RESET}""")

            # 确保文件夹存在
            os.makedirs(RESULT_FOLDER, exist_ok=True)
            os.makedirs(JSON_FOLDER, exist_ok=True)
            if all_vulns:
                save_vulnerabilities_to_json(target_name, all_vulns, target_date)
            # 写入Excel文件（使用目标日期生成文件名）
            write_to_excel(target_name, new_vulns, existing_vulns, day)
            # 只有新漏洞才推送
            if webhook_url and new_vulns:
                send_wechat_notification(webhook_url, target_name, new_vulns, day_ago)
        else:
            log_message(f"\n在{chinese_target_date}后没有发现任何漏洞", level="info")

        duration = (datetime.now() - start_time).total_seconds()
        log_message(f"扫描完成, 耗时: {duration:.2f}秒", level="info")

    except Exception as e:
        log_message(f"程序运行出错: {e}", level="error")
        exit(1)


if __name__ == '__main__':
    main()