import time
import schedule
import subprocess
import yaml
from datetime import datetime
import threading
import re

# 颜色常量定义
COLOR_RED = "\033[1;31m"
COLOR_GREEN = "\033[1;32m"
COLOR_YELLOW = "\033[1;33m"
COLOR_BLUE = "\033[1;34m"
COLOR_MAGENTA = "\033[1;35m"
COLOR_CYAN = "\033[1;36m"
COLOR_RESET = "\033[0m"

def load_config():
    """加载配置文件"""
    try:
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            return config
    except Exception as e:
        print(f"加载配置文件失败: {e}")
        return None


def detect_errors(output):
    """检测主程序输出中的各种错误"""
    # 定义错误模式的正则表达式
    error_patterns = [
        r'Error:',  # 通用错误
        r'Error in config\.yaml:',  # 配置文件错误
        r'Error reading config\.yaml:',  # 配置文件读取错误
        r'config\.yaml必须包含',  # 必填字段缺失
        r'config\.yaml中的页数范围无效',  # 页数配置错误
        r'HTTP错误:',  # 请求HTTP错误
        r'请求错误:',  # 请求错误
        r'其他错误:',  # 其他未分类错误
        r'写入Excel文件失败:',  # Excel写入失败
        r'企业微信机器人推送失败:',  # 微信推送失败
        r'企业微信机器人推送异常:',  # 微信推送异常
    ]

    # 检查是否有任何错误模式匹配
    for pattern in error_patterns:
        if re.search(pattern, output):
            return True

    # 检查是否有异常退出代码
    if "exit(1)" in output:
        return True

    return False


def run_monitor():
    """执行监控脚本"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"{COLOR_CYAN}[{timestamp}] [INFO] 开始执行监控任务...{COLOR_RESET}")
    try:
        # 使用subprocess调用原脚本
        result = subprocess.run(['python3', 'nvdmonitor.py'],
                               capture_output=True,
                               text=True)

        # 合并输出和错误信息用于检查
        full_output = result.stdout + "\n" + result.stderr

        # 判断执行是否正常
        if result.returncode == 0:
            # 检查主程序是否有任何错误输出
            if detect_errors(full_output):
                print(f"{COLOR_RED}[{timestamp}] [ERROR] 监控任务执行完成，但检测到异常输出{COLOR_RESET}")
                print(f"{COLOR_YELLOW}[DETAIL] 输出内容:{COLOR_RESET}")
                print(full_output)
                return False
            else:
                print(f"{COLOR_GREEN}[{timestamp}] [SUCCESS] 监控任务执行成功{COLOR_RESET}")
                if result.stdout.strip():
                    print(f"{COLOR_BLUE}[DETAIL] 输出内容:{COLOR_RESET}")
                    print(result.stdout)
                return True
        else:
            print(f"{COLOR_RED}[{timestamp}] [ERROR] 监控任务执行失败{COLOR_RESET}")
            if full_output.strip():
                print(f"{COLOR_YELLOW}[DETAIL] 输出内容:{COLOR_RESET}")
                print(full_output)
            return False

    except Exception as e:
        print(f"{COLOR_RED}[{timestamp}] [ERROR] 监控任务执行异常: {e}{COLOR_RESET}")
        return False



def setup_scheduler():
    """设置调度任务"""
    config = load_config()
    if not config:
        return

    scheduler_config = config.get('scheduler', {})
    mode = scheduler_config.get('mode', 'fixed')

    if mode == 'fixed':
        # 固定时间模式
        times = scheduler_config.get('fixed_times', [])
        for time_str in times:
            schedule.every().day.at(time_str).do(run_monitor_with_heartbeat)
            print(f"已设置每天 {time_str} 执行监控任务")
    else:
        # 间隔时间模式
        interval = scheduler_config.get('interval_hours', 6)
        schedule.every(interval).hours.do(run_monitor_with_heartbeat)
        print(f"已设置每 {interval} 小时执行一次监控任务")


def run_monitor_with_heartbeat():
    """执行监控任务并立即进行心跳检测"""
    # 执行监控任务
    status = run_monitor()

    # 立即进行心跳检测
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if status:
        print(f"{COLOR_GREEN}[{timestamp}] [HEARTBEAT] 监控任务完成 - 状态正常{COLOR_RESET}")
    else:
        print(f"{COLOR_RED}[{timestamp}] [HEARTBEAT] 监控任务完成 - 状态异常{COLOR_RESET}")

    # 显示下次执行时间
    next_run = schedule.next_run()
    if next_run:
        print(f"{COLOR_CYAN}[{timestamp}] [INFO] 下次执行时间: {next_run.strftime('%Y-%m-%d %H:%M:%S')}{COLOR_RESET}")



def run_scheduler():
    """运行调度器"""
    while True:
        try:
            schedule.run_pending()
            time.sleep(1)  # 减少等待时间，提高响应速度
        except Exception as e:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print(f"{COLOR_RED}[{timestamp}] [ERROR] 定时任务异常: {e}{COLOR_RESET}")
            time.sleep(60)  # 发生异常后等待60秒

if __name__ == '__main__':
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"{COLOR_GREEN}[{timestamp}] [INFO] 启动漏洞监控调度系统...{COLOR_RESET}")
    setup_scheduler()  # 设置定时任务

    # 初始心跳检测
    print(f"{COLOR_GREEN}[{timestamp}] [HEARTBEAT] 调度系统启动成功{COLOR_RESET}")

    try:
        run_scheduler()  # 进入定时循环
    except KeyboardInterrupt:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"{COLOR_YELLOW}[{timestamp}] [INFO] 监控调度系统已停止{COLOR_RESET}")