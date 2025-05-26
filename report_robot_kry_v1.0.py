import tkinter as tk
from tkinter import messagebox, ttk
import requests
import hashlib
import time
import json
import sys
import random
import threading
import os
import datetime
from datetime import timedelta

# 固定参数
FIXED_TOKEN = "TOKEN" # 请自行修改
FIXED_SHOP_ID = "00000000" # 请自行修改
FIXED_SHOP_NAME = "某某餐饮店" # 请自行修改
FIXED_APP_KEY = "APP_KEY"  # 请自行修改
FIXED_APP_SECRET = "APP_SECRET"  # 请自行修改
FIXED_FEISHU_WEBHOOK = "FEISHU_WEBHOOK"  # 请自行修改,以https开头的飞书机器 人地址。

class OrderReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{FIXED_SHOP_NAME}实时营业报告")
        self.root.geometry("650x500")  # 调整窗口大小
        self.root.resizable(True, True)  # 允许窗口调整大小
        
        # 确保中文显示正常
        self.configure_fonts()
        
        self.token = FIXED_TOKEN
        self.shop_id = FIXED_SHOP_ID
        self.shop_name = FIXED_SHOP_NAME
        self.app_key = FIXED_APP_KEY
        self.app_secret = FIXED_APP_SECRET
        self.feishu_webhook = FIXED_FEISHU_WEBHOOK
        self.last_report = ""
        self.running = False
        self.thread = None
        self.last_report_data = None
        
        self.create_widgets()
        self.root.after(3600000, self.clear_result_text)  # 每小时清空一次窗口显示内容

    def configure_fonts(self):
        # 设置支持中文的字体
        default_font = ('Microsoft YaHei UI', 10)
        text_font = ('Microsoft YaHei UI', 10)
        
        self.root.option_add('*Font', default_font)
        self.root.option_add('*Text.Font', text_font)

    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 状态框架
        status_frame = ttk.LabelFrame(main_frame, text="程序状态", padding="5")
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        self.time_var = tk.StringVar(value="")
        self.time_label = ttk.Label(status_frame, textvariable=self.time_var)
        self.time_label.pack(side=tk.RIGHT, padx=5)
        
        # 控制按钮框架
        control_frame = ttk.Frame(main_frame, padding="5")
        control_frame.pack(fill=tk.X, pady=5)
        
        self.btn_start = ttk.Button(control_frame, text="开始获取数据", command=self.start_application)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        
        self.btn_stop = ttk.Button(control_frame, text="停止", command=self.stop_application, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=5)
        
        self.btn_exit = ttk.Button(control_frame, text="退出", command=self.confirm_exit)
        self.btn_exit.pack(side=tk.RIGHT, padx=5)
        
        # 结果显示框架
        result_frame = ttk.LabelFrame(main_frame, text="营业数据", padding="5")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, font=('Microsoft YaHei UI', 11))
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(self.result_text, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
        
        # 日志框架
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="5")
        log_frame.pack(fill=tk.X, pady=5, ipady=5)
        
        self.log_text = tk.Text(log_frame, height=3, wrap=tk.WORD, font=('Microsoft YaHei UI', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 底部信息
        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        ttk.Label(footer_frame, text="版本: 1.7").pack(side=tk.RIGHT, padx=10, pady=5)

    def start_application(self):
        if self.running:
            return
        
        self.start_scheduled_checks()

    def start_scheduled_checks(self):
        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        
        self.log("开始获取营业数据...")
        self.status_var.set("运行中")
        
        # 在后台线程运行定时任务
        self.thread = threading.Thread(target=self.run_scheduled_task, daemon=True)
        self.thread.start()

    def run_scheduled_task(self):
        try:
            # 初始检查
            self.perform_check()
            
            # 主循环
            while self.running:
                now = datetime.datetime.now()
                current_time = now.strftime("%H:%M")
                
                # 根据时间段调整刷新频率
                if ("10:30" <= current_time <= "14:00") or ("17:00" <= current_time <= "23:00"):
                    interval = 60  # 1分钟
                else:
                    interval = 180  # 3分钟
                
                self.time_var.set(f"下次刷新: {datetime.datetime.now() + datetime.timedelta(seconds=interval)}")
                
                # 等待下一次检查
                time.sleep(interval)
                
                if self.running:
                    self.perform_check()
        except Exception as e:
            self.log(f"后台任务异常: {str(e)}")
            self.status_var.set("异常")
            self.running = False
            self.root.after(0, self.update_ui_state)

    def perform_check(self):
        self.log("正在获取最新数据...")
        try:
            delay = random.randint(2, 6)
            time.sleep(delay)  # 随机延迟避免API请求过于频繁
            self.get_order_data()
            self.log("数据更新成功")
        except Exception as e:
            self.log(f"获取数据失败: {str(e)}")
            self.status_var.set("错误")
            self.root.after(0, lambda: messagebox.showerror("请求失败", f"错误原因：{str(e)}"))

    def get_order_data(self):
        url = "https://openapi.keruyun.com/open/standard/order/queryList"
        timestamp = int(time.time())
        version = "2.0"
        common_params = {
            "appKey": self.app_key,
            "shopIdenty": self.shop_id,
            "version": version,
            "timestamp": timestamp
        }
        
        # 处理时区，使用北京时间（假设API服务端使用北京时间）
        now = datetime.datetime.now()
        start_date = now.strftime("%Y-%m-%d 00:00:00")
        end_date = now.strftime("%Y-%m-%d %H:%M:%S")
        
        all_orders = []
        page_num = 1
        page_size = 50  # 设定每页最大数量，根据API文档建议调整
        
        while True:
            body = {
                "dateType": "OPEN_TIME",  # 下单时间
                "startDate": start_date,
                "endDate": end_date,
                "orderTypeList": ["FOR_HERE"],  # 仅堂食订单
                "orderStatusList": ["WAIT_SETTLED", "SETTLED"],  # 包含未结账和已结账状态
                "pageBean": {
                    "pageNum": page_num,
                    "pageSize": page_size
                }
            }
            
            # 生成签名
            sorted_params = ''.join([f"{k}{v}" for k, v in sorted(common_params.items())])
            sign_str = f"{sorted_params}body{json.dumps(body)}{self.token}"
            sign = hashlib.sha256(sign_str.encode()).hexdigest()
            
            params = {**common_params, "sign": sign}
            headers = {"Content-Type": "application/json"}
            
            # 发送请求
            response = requests.post(url, headers=headers, params=params, json=body, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            if data.get("code") != 0:
                raise Exception(f"API返回错误：{data.get('message', '未知错误')}")
            
            result_data = data["result"]["data"]
            all_orders.extend(result_data["list"])
            
            total_count = result_data["totalCount"]
            # 修复类型错误：确保比较的是整数
            if (page_num * page_size) >= int(total_count):
                break  # 所有分页数据已获取
            page_num += 1
        
        self.process_response({"result": {"data": {"list": all_orders, "totalCount": total_count}}})

    def process_response(self, data):
        try:
            data_list = data["result"]["data"]["list"]
            
            # 过滤出堂食订单
            all_dine_in_orders = [order for order in data_list if order["orderType"] == "FOR_HERE"]
            
            # 计算总订单数（使用API返回的totalCount，处理分页情况）
            total_orders = data["result"]["data"]["totalCount"]
            
            # 已结账订单
            settled_orders = [order for order in all_dine_in_orders if order["orderStatus"] == "SETTLED"]
            
            # 未结账订单（仅保留C版本，正确逻辑）
            unsettled_orders = [order for order in all_dine_in_orders if order["orderStatus"] == "WAIT_SETTLED"]
            
            # 安全获取金额（从分转换为元）
            def safe_get_amount(order, key):
                try:
                    return float(order.get(key, 0)) / 100
                except (ValueError, TypeError):
                    return 0.0
            
            # 计算各项金额
            settled_turnover = sum(safe_get_amount(order, "orderAmt") for order in settled_orders)
            settled_income = sum(safe_get_amount(order, "orderReceivedAmt") for order in settled_orders)
            unsettled_income = sum(safe_get_amount(order, "orderAmt") for order in unsettled_orders)
            
            # 生成报告
            report_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            report = (
                f"【{self.shop_name}】实时营业报告\n"
                f"更新时间: {report_time}\n\n"
                f"今日截止当前订单总数（堂食）：{total_orders}单\n"
                f"已结账营业额：{settled_turnover:.2f}元\n"
                f"已结账营业收入：{settled_income:.2f}元\n"
                f"未结账订单数：{len(unsettled_orders)}单\n"
                f"未结账收入：{unsettled_income:.2f}元\n"
            )
            
            current_report_data = {
                "total_orders": total_orders,
                "settled_turnover": settled_turnover,
                "settled_income": settled_income,
                "unsettled_orders_count": len(unsettled_orders),
                "unsettled_income": unsettled_income
            }
            
            if self.last_report_data is None or current_report_data != self.last_report_data:
                # 更新UI
                self.root.after(0, self.update_report, report)
                self.last_report_data = current_report_data
            else:
                # 在程序窗口中显示提示信息
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, "数据无更新")
        except Exception as e:
            raise Exception(f"数据解析失败：{str(e)}")

    def update_report(self, report):
        if report != self.last_report:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, report)
            self.send_to_feishu(report)
            self.last_report = report

    def send_to_feishu(self, message):
        headers = {"Content-Type": "application/json"}
        data = {
            "msg_type": "text",
            "content": {"text": message}
        }
        try:
            response = requests.post(self.feishu_webhook, headers=headers, json=data, timeout=5)
            response.raise_for_status()
            self.log("已发送到飞书机器人")
        except requests.RequestException as e:
            self.log(f"飞书发送失败: {str(e)}")

    def stop_application(self):
        self.running = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.status_var.set("已停止")
        self.log("程序已停止")

    def confirm_exit(self):
        if self.running:
            if messagebox.askyesno("确认退出", "程序正在运行中，是否确认退出？"):
                self.running = False
                self.root.destroy()
        else:
            self.root.destroy()

    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}")

    def clear_result_text(self):
        self.result_text.delete(1.0, tk.END)
        self.root.after(3600000, self.clear_result_text)  # 每小时清空一次窗口显示内容

if __name__ == "__main__":
    # 确保中文显示正常
    if sys.platform.startswith('win'):
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        # 隐藏命令行窗口
        import win32gui, win32con
        hwnd = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    
    root = tk.Tk()
    app = OrderReportApp(root)
    
    root.mainloop()