import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import json
import requests
import threading
import subprocess
import os
import queue
import re
from datetime import datetime
import time # _close_and_delete_browser_via_bitapi 中需要

# --- 核心修改：使用 playsound3 替代 playsound ---
try:
    # 使用 playsound3 库
    from playsound3 import playsound
    PLAYSOUND_AVAILABLE = True
except ImportError:
    # 如果 playsound3 未安装，则尝试使用内置的 winsound 作为备用
    try:
        import winsound
        PLAYSOUND_AVAILABLE = 'winsound'
    except ImportError:
        playsound = None
        PLAYSOUND_AVAILABLE = False
        print("警告: playsound3 和 winsound 库均未安装。将无法播放任何音效。请运行 'pip install playsound3'。")


class AwsAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AWS自动化控制面板 V8.3 (最终修复版)") # 版本号更新
        self.root.geometry("1400x850")

        self.BIT_API_BASE_URL = "http://127.0.0.1:54345"

        sounds_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'shared', 'sounds')
        self.sound_paths = {
            "ready": os.path.join(sounds_dir, 'ready.wav'),
            "boom": os.path.join(sounds_dir, 'boom.wav'),
            "success": os.path.join(sounds_dir, 'ohhyeah.wav'),
            "failure": os.path.join(sounds_dir, 'meiqian.wav'),
            "batch_complete": os.path.join(sounds_dir, 'jieshan.wav')
        }

        self.node_processes = {}
        self.log_queue = queue.Queue()
        self.pause_states = {}
        self.log_tabs = {}
        self.original_data_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'Original_data.json')
        self.combined_country_data_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'shared', 'combined_country_data.json')
        self.not_used_cards_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'Not used cards.txt')
        self.workline_stats_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'workline_stats.json')
        self.gui_settings_file = "gui_settings.json"
        
        self.country_data_map = {}
        self.dealer_thread_running = False
        self.dealer_thread_stop_event = threading.Event()
        self.available_worklines_queue = queue.Queue()
        self.workline_ports = {}
        self.next_available_port = 45000
        self.workline_stats = self._load_workline_stats()
        self.active_workline_data = {}
        self.total_worklines_started = 0
        self.batch_complete_sound_played = False

        self.dealer_pause_event = threading.Event()

        # --- 第四部分: 原始日志输出 ---
        frame4 = ttk.LabelFrame(self.root, text="第四部分: 原始日志输出", padding=(10, 5))
        frame4.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.notebook = ttk.Notebook(frame4)
        self.notebook.pack(fill="both", expand=True)

        self.log_all_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_all_frame, text='全部日志')
        self.log_all_text = scrolledtext.ScrolledText(self.log_all_frame, wrap=tk.WORD, height=10)
        self.log_all_text.pack(fill="both", expand=True)

        self.clear_log_button = ttk.Button(frame4, text="清除所有日志", command=self.clear_logs)
        self.clear_log_button.pack(fill="x", padx=0, pady=5, side="bottom")

        self.load_country_data()
        self._schedule_log_cleanup()

        # --- 顶部控制区 ---
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=5)

        # --- 第一部分: 原数据输入与处理 ---
        frame1 = ttk.LabelFrame(top_frame, text="第一步: 原数据输入与处理", padding=(10, 5))
        frame1.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        ttk.Label(frame1, text="原始注册数据 (格式: mailbox\\t...\\tcountry_code\\tproxy_country_code):").pack(anchor='w', padx=5, pady=2)
        self.raw_data_input = scrolledtext.ScrolledText(frame1, wrap=tk.WORD, height=8)
        self.raw_data_input.pack(fill="both", expand=True, padx=5, pady=2)

        self.process_data_button = ttk.Button(frame1, text="预处理数据", command=self.process_raw_data_and_distribute)
        self.process_data_button.pack(pady=5, fill='x', padx=5)

        self.data_remaining_label = ttk.Label(frame1, text="原始数据剩余: 0 条")
        self.data_remaining_label.pack(anchor='w', padx=5, pady=2)

        self.total_success_label = ttk.Label(frame1, text="总成功数: 0")
        self.total_success_label.pack(anchor='w', padx=5, pady=2)
        self.total_failure_label = ttk.Label(frame1, text="总失败数: 0")
        self.total_failure_label.pack(anchor='w', padx=5, pady=2)
        self._update_total_stats_display()

        self.clear_total_stats_button = ttk.Button(frame1, text="清零总战绩", command=self.clear_all_workline_stats)
        self.clear_total_stats_button.pack(pady=5, fill='x', padx=5)


        # --- 第二部分: 启动自动化任务 ---
        frame2 = ttk.LabelFrame(top_frame, text="第二步: 启动与控制", padding=(10, 5))
        frame2.pack(side="right", fill="y", padx=(5, 0))

        ttk.Label(frame2, text="工作线数量:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.num_worklines_entry = ttk.Entry(frame2, width=15)
        self.num_worklines_entry.insert(0, "1")
        self.num_worklines_entry.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        self.headless_var = tk.BooleanVar()
        self.headless_check = ttk.Checkbutton(frame2, text="无头模式", variable=self.headless_var)
        self.headless_check.grid(row=1, column=0, columnspan=2, padx=5, pady=2, sticky='w')
        
        ttk.Label(frame2, text="邮箱接口:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.email_api_provider_var = tk.StringVar()
        self.email_api_provider_combo = ttk.Combobox(frame2, textvariable=self.email_api_provider_var, 
                                                     values=['Domain', 'Outlook'], state='readonly', width=12)
        self.email_api_provider_combo.grid(row=2, column=1, padx=5, pady=2, sticky='w')
        self.email_api_provider_combo.set('Domain')

        button_frame = ttk.Frame(frame2)
        button_frame.grid(row=3, column=0, columnspan=2, pady=5)

        self.start_button = ttk.Button(button_frame, text="!! 开始运行 !!", command=self.start_automation_orchestrator)
        self.start_button.pack(side="top", fill='x', pady=2)

        self.stop_button = ttk.Button(button_frame, text="!! 停止所有脚本 !!", command=self.stop_all_automation, state="disabled")
        self.stop_button.pack(side="top", fill='x', pady=2)

        self.close_all_button = ttk.Button(button_frame, text="关闭并删除所有窗口", command=self.show_close_all_dialog, state="disabled")
        self.close_all_button.pack(side="top", fill='x', pady=2)
        
        self.toggle_dealer_button = ttk.Button(button_frame, text="停止发牌", command=self.toggle_dealer_pause, state="disabled")
        self.toggle_dealer_button.pack(side="top", fill='x', pady=2)

        # --- 第三部分: 实时状态监控 ---
        frame3 = ttk.LabelFrame(self.root, text="第三部分: 窗口实时状态监控", padding=(10, 5))
        frame3.pack(fill="x", padx=10, pady=5)
        
        self.tree = ttk.Treeview(frame3, columns=("window", "email", "status", "details", "success", "failure", "action", "manage", "screenshot"), show="headings", height=8)
        self.tree.heading("window", text="窗口名")
        self.tree.heading("email", text="邮箱")
        self.tree.heading("status", text="状态")
        self.tree.heading("details", text="详情")
        self.tree.heading("success", text="成功")
        self.tree.heading("failure", text="失败")
        self.tree.heading("action", text="操作")
        self.tree.heading("manage", text="管理")
        self.tree.heading("screenshot", text="截图")

        self.tree.column("window", width=80, anchor='center')
        self.tree.column("email", width=180)
        self.tree.column("status", width=100, anchor='center')
        self.tree.column("details", width=200)
        self.tree.column("success", width=60, anchor='center')
        self.tree.column("failure", width=60, anchor='center')
        self.tree.column("action", width=80, anchor='center')
        self.tree.column("manage", width=80, anchor='center')
        self.tree.column("screenshot", width=80, anchor='center')
        
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Button-1>", self.on_tree_click)

        self.root.after(100, self.process_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.load_settings()

    def save_settings(self):
        settings = {
            "num_worklines": self.num_worklines_entry.get(),
            "headless_mode": self.headless_var.get(),
            "email_api_provider": self.email_api_provider_var.get()
        }
        try:
            with open(self.gui_settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            self.log(f"[GUI-ERROR] 保存GUI设置失败: {e}")

    def load_settings(self):
        try:
            if os.path.exists(self.gui_settings_file):
                with open(self.gui_settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.num_worklines_entry.delete(0, tk.END)
                    self.num_worklines_entry.insert(0, settings.get("num_worklines", "1"))
                    self.headless_var.set(settings.get("headless_mode", False))
                    self.email_api_provider_var.set(settings.get("email_api_provider", "Domain"))
        except (IOError, json.JSONDecodeError) as e:
            self.log(f"[GUI-WARN] 加载GUI设置失败: {e}，将使用默认值。")
            self.num_worklines_entry.insert(0, "1")
            self.headless_var.set(False)
            self.email_api_provider_var.set("Domain")

    def toggle_dealer_pause(self):
        if self.dealer_pause_event.is_set():
            self.dealer_pause_event.clear() 
            self.toggle_dealer_button.config(text="恢复发牌")
            self.log("[GUI-发牌师] 发牌师已暂停 (将完成当前任务后生效)。")
        else:
            self.dealer_pause_event.set()
            self.toggle_dealer_button.config(text="停止发牌")
            self.log("[GUI-发牌师] 发牌师已恢复。")

    def _play_sound_non_blocking(self, sound_name):
        if not PLAYSOUND_AVAILABLE: return
        sound_path = self.sound_paths.get(sound_name)
        if sound_path and os.path.exists(sound_path):
            try:
                if PLAYSOUND_AVAILABLE == 'winsound':
                    winsound.PlaySound(sound_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                else:
                    threading.Thread(target=playsound, args=(sound_path,), daemon=True).start()
            except Exception as e:
                self.log(f"[GUI-AUDIO-ERROR] 播放音效 '{sound_path}' 失败: {e}")
        else:
            self.log(f"[GUI-AUDIO-WARN] 音效文件未找到: {sound_path}")

    def _schedule_log_cleanup(self):
        self.clear_logs()
        self.root.after(600000, self._schedule_log_cleanup)

    def _load_workline_stats(self):
        if os.path.exists(self.workline_stats_path):
            try:
                with open(self.workline_stats_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                self.log(f"[GUI-WARN] 战绩文件 {self.workline_stats_path} 损坏，将重新创建。")
                return {}
        return {}

    def _save_workline_stats(self):
        os.makedirs(os.path.dirname(self.workline_stats_path), exist_ok=True)
        with open(self.workline_stats_path, 'w', encoding='utf-8') as f:
            json.dump(self.workline_stats, f, indent=4, ensure_ascii=False)
        self._update_total_stats_display()

    def clear_all_workline_stats(self):
        if messagebox.askokcancel("确认", "确定要清零所有工作线的成功和失败总数吗？此操作不可恢复。"):
            for workline_id in self.workline_stats:
                self.workline_stats[workline_id]['success'] = 0
                self.workline_stats[workline_id]['failure'] = 0
            
            self._save_workline_stats()

            for item_id in self.tree.get_children():
                current_values = list(self.tree.item(item_id, 'values'))
                if current_values:
                    current_values[4] = 0
                    current_values[5] = 0
                    self.tree.item(item_id, values=tuple(current_values))
            
            self.log("[GUI-管理] 所有工作线的战绩已清零。")

    def _update_total_stats_display(self):
        total_success = sum(stats.get('success', 0) for stats in self.workline_stats.values())
        total_failure = sum(stats.get('failure', 0) for stats in self.workline_stats.values())
        self.total_success_label.config(text=f"总成功数: {total_success}")
        self.total_failure_label.config(text=f"总失败数: {total_failure}")

    def load_country_data(self):
        try:
            with open(self.combined_country_data_path, 'r', encoding='utf-8') as f:
                raw_data = json.load(f)
                for country_full_name, details in raw_data.items():
                    if 'numeric_id' in details and 'dialing_code' in details:
                        self.country_data_map[details['country_code'].upper()] = {
                            'numeric_id': details['numeric_id'],
                            'full_name': country_full_name,
                            'dialing_code': details['dialing_code']
                        }
            self.log(f"[GUI] 成功加载国家数据，共 {len(self.country_data_map)} 条。")
        except Exception as e:
            self.log(f"[GUI-ERROR] 加载国家数据文件失败: {e}")
            messagebox.showerror("错误", f"加载国家数据文件失败: {e}")

    def log(self, message, instance_id=None):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_all_text.insert(tk.END, log_entry)
        self.log_all_text.see(tk.END)
        
        if instance_id and instance_id in self.log_tabs:
            log_widget = self.log_tabs[instance_id]
            log_widget.insert(tk.END, log_entry)
            log_widget.see(tk.END)

    def clear_logs(self):
        self.log_all_text.delete('1.0', tk.END)
        for log_widget in self.log_tabs.values():
            log_widget.delete('1.0', tk.END)
        self.log("所有日志已清除。")

    def process_raw_data_and_distribute(self):
        raw_text = self.raw_data_input.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning("警告", "原数据输入框不能为空！")
            return

        lines = raw_text.split('\n')
        newly_processed_data = []
        
        for line_num, line in enumerate(lines):
            line = line.strip()
            if not line: continue
            
            parts = line.split('\t')
            if len(parts) < 6:
                self.log(f"[GUI-WARN] 原始数据行 {line_num+1} 格式不正确 (字段少于6)，跳过: {line}")
                continue

            try:
                mailbox_full, account, password = parts[0].strip(), parts[1].strip(), parts[2].strip()
                step1_info_parts = parts[3].strip().split('|')
                if len(step1_info_parts) != 4:
                    self.log(f"[GUI-WARN] 原始数据行 {line_num+1} 的1step_info格式不正确，跳过: {line}")
                    continue
                step1_number, step1_month_year_raw, step1_code, real_name = [p.strip() for p in step1_info_parts]
                step1_month, step1_year_raw = step1_month_year_raw.split('/')
                step1_year = f"20{step1_year_raw.strip()}" if len(step1_year_raw.strip()) == 2 else step1_year_raw.strip()
                
                step2_parts = re.split(r'\s{2,}', parts[4].strip())
                if len(step2_parts) != 3:
                     self.log(f"[GUI-WARN] 原始数据行 {line_num+1} 的2step_info格式不正确，跳过: {line}")
                     continue
                step2_number, step2_year_month, step2_code = [p.strip() for p in step2_parts]
                step2_year, step2_month = step2_year_month[:2], step2_year_month[2:]
                
                # --- FINAL BUG FIX ---
                # 注册国家代码
                country_code = parts[5].strip().upper()
                
                # 代理国家代码，无需验证，直接使用
                if len(parts) >= 7 and parts[6].strip():
                    proxy_country_code = parts[6].strip().upper()
                else:
                    proxy_country_code = country_code # 默认等于注册国家

                # 只验证注册国家代码是否存在于我们的数据映射中
                country_info = self.country_data_map.get(country_code)
                if not country_info:
                    self.log(f"[GUI-WARN] 原始数据行 {line_num+1} 的注册国家代码 '{country_code}' 未在数据文件中找到，跳过。")
                    continue
                
                newly_processed_data.append({
                    "mailbox": mailbox_full, "account": account, "password": password,
                    "1step_number": step1_number, "1step_month": step1_month.strip(), "1step_year": step1_year, "1step_code": step1_code,
                    "2step_number": step2_number, "2step_year": step2_year, "2step_month": step2_month, "2step_code": step2_code,
                    "real_name": real_name, "street": "", "city": "", "state": "", "postcode": "",
                    "country_full_name": country_info['full_name'], 
                    "country_code": country_code,
                    "proxy_country_code": proxy_country_code,
                    "phone_number_id": None, "phone_number": None, "phone_number_url": None,
                    "numeric_id": country_info['numeric_id'], "dialing_code": country_info['dialing_code']
                })
            except Exception as e:
                self.log(f"[GUI-ERROR] 处理原始数据行 {line_num+1} 时出错: {e} -> 原始行: '{line}'")
                continue

        if not newly_processed_data:
            messagebox.showerror("错误", "未能成功解析任何数据！请检查格式。")
            return

        try:
            existing_data = []
            if os.path.exists(self.original_data_path):
                try:
                    with open(self.original_data_path, 'r', encoding='utf-8') as f:
                        file_content = f.read().strip()
                        if file_content: existing_data = json.loads(file_content)
                except json.JSONDecodeError:
                    existing_data = []
            
            combined_data = existing_data + newly_processed_data
            os.makedirs(os.path.dirname(self.original_data_path), exist_ok=True)
            with open(self.original_data_path, 'w', encoding='utf-8') as f:
                json.dump(combined_data, f, indent=4, ensure_ascii=False)
            
            self.update_data_remaining_label()
            self.log(f"[GUI] ✅ 成功追加 {len(newly_processed_data)} 条数据，当前共 {len(combined_data)} 条。")
            messagebox.showinfo("成功", f"成功追加 {len(newly_processed_data)} 条数据。")
            self._play_sound_non_blocking("ready")
        except Exception as e:
            self.log(f"[GUI] ❌ 保存 Original_data.json 时出错: {e}")
            messagebox.showerror("错误", f"保存 Original_data.json 时出错: {e}")

    def update_data_remaining_label(self):
        try:
            with open(self.original_data_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.data_remaining_label.config(text=f"原始数据剩余: {len(data)} 条")
        except (FileNotFoundError, json.JSONDecodeError):
            self.data_remaining_label.config(text="原始数据剩余: 0 条 (文件不存在或损坏)")

    def fetch_phone_number_from_api(self, numeric_id, phone_number_dialing_code):
        api_url = f"https://api.small5.co/hub/en/proxy.php?action=getNumber&service=am&country={numeric_id}&platform=sms"
        try:
            response = requests.get(api_url, timeout=30)
            response.raise_for_status()
            raw_data = response.text
            self.log(f"[GUI-API] 获取手机号原始响应: {raw_data}")

            if raw_data.startswith('ACCESS_NUMBER:'):
                parts = raw_data.split(':')
                if len(parts) >= 3:
                    phone_number_id = parts[1]
                    raw_phone_number = parts[2]
                    
                    if raw_phone_number.startswith(phone_number_dialing_code):
                        cleaned_phone_number = raw_phone_number[len(phone_number_dialing_code):]
                        self.log(f"[GUI-API] 已去除区号：{raw_phone_number} -> {cleaned_phone_number}")
                    else:
                        cleaned_phone_number = raw_phone_number
                        self.log(f"[GUI-API-WARN] 获取的手机号 '{raw_phone_number}' 未以区号 '{phone_number_dialing_code}' 开头，未去除区号。")

                    phone_number_url = f"https://api.small5.co/hub/en/proxy.php?action=getStatus&id={phone_number_id}&platform=sms"
                    return {
                        "phone_number_id": phone_number_id,
                        "phone_number": cleaned_phone_number,
                        "phone_number_url": phone_number_url
                    }
            raise ValueError(f"API响应格式不正确: {raw_data}")
        except requests.exceptions.RequestException as e:
            raise Exception(f"请求手机号API失败: {e}")
        except ValueError as e:
            raise Exception(f"解析手机号API响应失败: {e}")

    def fetch_address_from_api(self, country_code):
        address_api_url = f"https://rd.32v.us/api/random-address?country={country_code}"
        try:
            response = requests.get(address_api_url, timeout=30)
            response.raise_for_status()
            address_data = response.json()
            self.log(f"[GUI-API] 获取地址原始响应: {json.dumps(address_data)}")
            
            return {
                "street": address_data.get("address_line", ""),
                "city": address_data.get("city", ""),
                "state": address_data.get("state", ""),
                "postcode": address_data.get("postcode", "")
            }
        except requests.exceptions.RequestException as e:
            self.log(f"[GUI-API-ERROR] 请求地址API失败: {e}. 国家代码: {country_code}")
            raise Exception(f"请求地址API失败: {e}")
        except json.JSONDecodeError as e:
            self.log(f"[GUI-API-ERROR] 解析地址API响应失败: {e}. 国家代码: {country_code}")
            raise Exception(f"解析地址API响应失败: {e}")

    def update_proxy_ip(self, port, country_code):
        proxy_api_url = 'http://localhost:8080/api/proxy/start'
        payload = {
            "line": "Line A (AS Route)", "country_code": country_code, "start_port": port, "count": 1, "time": 30
        }
        try:
            self.log(f"[GUI-PROXY] 正在请求更新端口 {port} 的IP (国家: {country_code})...")
            response = requests.post(proxy_api_url, json=payload, timeout=25)
            response.raise_for_status()
            response_data = response.json()
            if response_data.get('success'):
                self.log(f"[GUI-PROXY] 端口 {port} IP更新成功: {response_data.get('msg')}")
                return True
            else:
                self.log(f"[GUI-PROXY-ERROR] 端口 {port} IP更新失败: {response_data.get('msg')}")
                raise Exception(f"IP更新API返回失败: {response_data.get('msg')}")
        except requests.exceptions.RequestException as e:
            self.log(f"[GUI-PROXY-ERROR] 请求IP更新API失败 (端口 {port}): {e}")
            raise Exception(f"请求IP更新API失败: {e}")

    def dealer_thread(self):
        self.dealer_thread_running = True
        self.log("[GUI-发牌师] 发牌师线程已启动，开始分配任务...")
        
        num_worklines = int(self.num_worklines_entry.get())
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        if not self.workline_ports:
            current_port = self.next_available_port
            for i in range(num_worklines):
                workline_id = f"W{i + 1}"
                self.workline_ports[workline_id] = current_port
                current_port += 1
                self.log(f"[GUI-发牌师] 工作线 {workline_id} 分配端口: {self.workline_ports[workline_id]}")
            self.next_available_port = current_port

        for i in range(num_worklines):
            workline_id = f"W{i + 1}"
            if workline_id in self.workline_ports:
                workline_data_dir = os.path.join(project_root, 'data', workline_id)
                os.makedirs(workline_data_dir, exist_ok=True)
                if workline_id not in list(self.available_worklines_queue.queue):
                    self.available_worklines_queue.put(workline_id) 

        while not self.dealer_thread_stop_event.is_set():
            self.dealer_pause_event.wait() 

            workline_id = None
            current_data_from_original = None
            
            try:
                workline_id = self.available_worklines_queue.get(timeout=1)

                with open(self.original_data_path, 'r', encoding='utf-8') as f: 
                    file_content = f.read().strip()
                    original_data_in_mem = json.loads(file_content) if file_content else []

                if not original_data_in_mem:
                    self.log(f"[GUI-发牌师] Original_data.json 中无更多数据。工作线 {workline_id} 将保持空闲。")
                    self.available_worklines_queue.put(workline_id)
                    self.dealer_thread_stop_event.wait(5)
                    continue
                
                current_data_from_original = original_data_in_mem.pop(0) 

                try:
                    self.active_workline_data[workline_id] = current_data_from_original.copy()
                    
                    email_api_provider = self.email_api_provider_var.get()
                    current_data_from_original['email_api_provider'] = email_api_provider
                    self.log(f"[GUI-发牌师] 为 {workline_id} 设置邮箱接口为: {email_api_provider}", instance_id=workline_id)

                    self.log(f"[GUI-发牌师] 为工作线 {workline_id} 获取地址信息 (国家: {current_data_from_original['country_code']})...", instance_id=workline_id)
                    address_info = self.fetch_address_from_api(current_data_from_original['country_code'])
                    current_data_from_original.update(address_info)

                    self.log(f"[GUI-发牌师] 为工作线 {workline_id} 获取手机号 (国家: {current_data_from_original['country_code']})...", instance_id=workline_id)
                    phone_info = self.fetch_phone_number_from_api(current_data_from_original['numeric_id'], current_data_from_original['dialing_code'])
                    current_data_from_original.update(phone_info)

                    proxy_port = self.workline_ports[workline_id]
                    proxy_country = current_data_from_original['proxy_country_code']
                    self.log(f"[GUI-发牌师] 为工作线 {workline_id} 更新代理IP (代理国家: {proxy_country})...", instance_id=workline_id)
                    ip_updated = self.update_proxy_ip(proxy_port, proxy_country)
                    if not ip_updated:
                        self.log(f"[GUI-发牌师-WARN] 工作线 {workline_id} 的IP更新失败。", instance_id=workline_id)

                    workline_data_path = os.path.join(project_root, 'data', workline_id, 'data.json')
                    with open(workline_data_path, 'w', encoding='utf-8') as f:
                        json.dump([current_data_from_original], f, indent=4, ensure_ascii=False)
                    
                    self.log(f"[GUI-发牌师] 数据已分配至 {workline_id}/data.json。即将启动脚本。", instance_id=workline_id)
                    self.start_single_node_process(workline_id, proxy_port)

                    with open(self.original_data_path, 'w', encoding='utf-8') as f:
                        json.dump(original_data_in_mem, f, indent=4, ensure_ascii=False)
                    self.update_data_remaining_label()

                except Exception as prep_error:
                    self.log(f"[GUI-发牌师-错误] 工作线 {workline_id} 任务准备阶段失败: {prep_error}", instance_id=workline_id)
                    original_data_in_mem.insert(0, current_data_from_original)
                    with open(self.original_data_path, 'w', encoding='utf-8') as f:
                        json.dump(original_data_in_mem, f, indent=4, ensure_ascii=False)
                    self.update_data_remaining_label()
                    self.log(f"[GUI-发牌师] 原始数据已回滚。", instance_id=workline_id)
                    
                    if workline_id in self.active_workline_data:
                        del self.active_workline_data[workline_id]

                    self.available_worklines_queue.put(workline_id)
                    continue

            except queue.Empty:
                pass
            except FileNotFoundError:
                self.log(f"[GUI-发牌师] Original_data.json 文件不存在或已被清空。")
                self.dealer_thread_stop_event.set()
            except json.JSONDecodeError:
                self.log(f"[GUI-发牌师] Original_data.json 文件内容格式错误。")
                self.dealer_thread_stop_event.set()
            except Exception as e:
                self.log(f"[GUI-发牌师-致命错误] 主循环发生意外错误: {e}", instance_id=workline_id if workline_id else "未知")
                self.dealer_thread_stop_event.set()

            self.dealer_thread_stop_event.wait(0.1)
        
        self.log("[GUI-发牌师] 发牌师线程已停止。")
        self.dealer_thread_running = False

    def start_automation_orchestrator(self):
        self.save_settings()
        self.tree.delete(*self.tree.get_children())
        self.pause_states = {}
        for i in list(self.notebook.tabs()):
            if self.notebook.tab(i, "text") != "全部日志":
                self.notebook.forget(i)
        self.log_tabs.clear()
        self.active_workline_data.clear()
        self.available_worklines_queue = queue.Queue()

        num_worklines_str = self.num_worklines_entry.get()
        try:
            num_worklines = int(num_worklines_str)
            if num_worklines <= 0: raise ValueError
        except ValueError:
            messagebox.showerror("错误", "工作线数量必须是一个正整数！")
            return
        
        try:
            with open(self.original_data_path, 'r', encoding='utf-8') as f:
                initial_data = json.load(f)
                if not initial_data:
                    messagebox.showwarning("警告", "Original_data.json 中没有数据！")
                    return
        except (FileNotFoundError, json.JSONDecodeError):
            messagebox.showwarning("警告", "Original_data.json 文件不存在或损坏！")
            return

        self.total_worklines_started = num_worklines
        self.batch_complete_sound_played = False

        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.close_all_button.config(state="normal")
        self.toggle_dealer_button.config(state="normal", text="停止发牌")

        self.dealer_thread_stop_event.clear()
        self.dealer_pause_event.set()
        threading.Thread(target=self.dealer_thread, daemon=True).start()
        
        self.log("[GUI] 自动化调度器已启动。")
        self._play_sound_non_blocking("boom")

    def start_single_node_process(self, instance_id, proxy_port):
        if instance_id in self.node_processes and self.node_processes[instance_id].poll() is None:
            self.log(f"[GUI-调度器] 警告: 工作线 {instance_id} 仍在运行。", instance_id)
            return

        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        main_script_path = os.path.join(project_root, 'main_controller.js')
        
        command = ['node', main_script_path, f'--window={instance_id}', f'--proxy_port={proxy_port}']
        if self.headless_var.get():
            command.append('--headless=new')
        
        self.log(f"[GUI-调度器] 启动命令: {' '.join(command)}", instance_id=instance_id)
        
        creationflags = 0x08000000 if os.name == 'nt' else 0 # CREATE_NO_WINDOW
        preexec_fn = os.setsid if os.name != 'nt' else None

        try:
            process = subprocess.Popen(
                command, stdout=subprocess.PIPE, stdin=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding='utf-8', errors='replace', bufsize=1, 
                creationflags=creationflags, cwd=project_root, preexec_fn=preexec_fn
            )
            self.node_processes[instance_id] = process
            threading.Thread(target=self.enqueue_output, args=(process.stdout, self.log_queue, instance_id), daemon=True).start()
            self.log(f"[GUI-调度器] 工作线 {instance_id} 脚本已启动。", instance_id=instance_id)
            
            if not self.tree.exists(instance_id):
                self.workline_stats.setdefault(instance_id, {'success': 0, 'failure': 0})
                self.tree.insert("", "end", iid=instance_id, 
                                 values=(instance_id, "N/A", "启动中", "", 
                                         self.workline_stats[instance_id]['success'], 
                                         self.workline_stats[instance_id]['failure'], 
                                         "暂停", "管理", "截图"))
            else:
                current_values = list(self.tree.item(instance_id, 'values'))
                current_values[2] = "启动中"
                current_values[3] = "正在执行任务"
                self.tree.item(instance_id, values=tuple(current_values))

            self.draw_buttons_for_item(instance_id)

        except Exception as e:
            self.log(f"[GUI-调度器-错误] 启动工作线 {instance_id} 失败: {e}", instance_id=instance_id)
            if instance_id in self.node_processes:
                del self.node_processes[instance_id]
            
            if instance_id in self.active_workline_data:
                failed_data = self.active_workline_data.pop(instance_id)
                try:
                    with open(self.original_data_path, 'r+', encoding='utf-8') as f_orig:
                        content = f_orig.read().strip()
                        original_data_list = json.loads(content) if content else []
                        original_data_list.insert(0, failed_data)
                        f_orig.seek(0)
                        f_orig.truncate()
                        json.dump(original_data_list, f_orig, indent=4, ensure_ascii=False)
                    self.update_data_remaining_label()
                    self.log(f"[GUI-调度器] 启动失败，数据已回滚。", instance_id=instance_id)
                except Exception as e_rollback:
                    self.log(f"[GUI-调度器-ERROR] 回滚数据时出错: {e_rollback}", instance_id=instance_id)

            self.available_worklines_queue.put(instance_id)

    def enqueue_output(self, out, queue_obj, instance_id):
        try:
            for line in iter(out.readline, ''):
                queue_obj.put({"type": "LOG", "payload": line, "instance_id": instance_id})
            out.close()
        except Exception as e:
            self.log(f"[GUI-ERROR] 读取 {instance_id} 输出时出错: {e}", instance_id)
        finally:
            queue_obj.put({"type": "PROCESS_EXIT", "payload": instance_id})

    def stop_all_automation(self):
        self.dealer_thread_stop_event.set()
        self.dealer_pause_event.set()
        self.log("[GUI] 正在终止所有Node.js脚本...")
        
        for instance_id, process in list(self.node_processes.items()):
            if process.poll() is None:
                try:
                    process.stdin.write(f"TERMINATE::{instance_id}\n") 
                    process.stdin.flush()
                    process.terminate()
                    try:
                        process.wait(timeout=2)
                    except subprocess.TimeoutExpired:
                        process.kill()
                    self.log(f"[GUI] 工作线 {instance_id} 已终止。", instance_id)
                except Exception as e:
                    self.log(f"[GUI-ERROR] 终止工作线 {instance_id} 失败: {e}", instance_id)
            if instance_id in self.node_processes:
                del self.node_processes[instance_id] 

        self.node_processes.clear()
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.close_all_button.config(state="disabled")
        self.toggle_dealer_button.config(state="disabled")
        self.log("[GUI] 所有自动化脚本已终止。")
        while not self.available_worklines_queue.empty():
            try:
                self.available_worklines_queue.get_nowait()
            except queue.Empty:
                break
        self._save_workline_stats()

    def process_log_queue(self):
        try:
            while True:
                message_obj = self.log_queue.get_nowait()
                msg_type = message_obj.get("type")
                payload = message_obj.get("payload")
                instance_id = message_obj.get("instance_id")

                if msg_type == "PROCESS_EXIT":
                    self.handle_process_exit(payload)
                    continue

                log_message = payload.strip()
                
                if not instance_id:
                    match = re.search(r'\[(W\d+)\]', log_message)
                    if match:
                        instance_id = match.group(1)
                
                if instance_id and instance_id not in self.log_tabs:
                    new_frame = ttk.Frame(self.notebook)
                    self.notebook.add(new_frame, text=instance_id)
                    new_log_text = scrolledtext.ScrolledText(new_frame, wrap=tk.WORD, height=10)
                    new_log_text.pack(fill="both", expand=True)
                    self.log_tabs[instance_id] = new_log_text
                    self.notebook.select(new_frame)
                
                if log_message.startswith("STATUS_UPDATE::"):
                    try:
                        status_data = json.loads(log_message.replace("STATUS_UPDATE::", ""))
                        instance_id = status_data.get("instanceId")
                        if not instance_id: continue

                        email = status_data.get("account", "")
                        status = status_data.get("status", "")
                        details = status_data.get("details", "")

                        current_values = list(self.tree.item(instance_id, 'values')) if self.tree.exists(instance_id) else [""] * 9
                        
                        self.workline_stats.setdefault(instance_id, {'success': 0, 'failure': 0})
                        if status == "成功":
                            self.workline_stats[instance_id]['success'] += 1
                            self._play_sound_non_blocking("success")
                        elif status == "失败":
                            self.workline_stats[instance_id]['failure'] += 1
                            self._play_sound_non_blocking("failure")
                        
                        self._save_workline_stats()

                        current_values[0] = instance_id
                        current_values[1] = email
                        current_values[2] = status
                        current_values[3] = details
                        current_values[4] = self.workline_stats[instance_id]['success']
                        current_values[5] = self.workline_stats[instance_id]['failure']
                        
                        if self.tree.exists(instance_id):
                            self.tree.item(instance_id, values=tuple(current_values))
                        else:
                            self.tree.insert("", "end", iid=instance_id, values=tuple(current_values))
                        
                        self.draw_buttons_for_item(instance_id)

                    except json.JSONDecodeError:
                        self.log(f"[GUI-ERROR] 无法解析状态消息: {log_message}")
                
                if log_message:
                    self.log(log_message, instance_id)
        except queue.Empty:
            pass
        finally:
            if not self.node_processes and not self.dealer_thread_running and self.stop_button['state'] == 'normal':
                 self.stop_all_automation() 
            self.root.after(100, self.process_log_queue)
    
    def handle_process_exit(self, exited_instance_id):
        self.log(f"[GUI-调度器] 检测到工作线 {exited_instance_id} 进程退出。", exited_instance_id)
        if exited_instance_id in self.node_processes:
            del self.node_processes[exited_instance_id]
        
        if exited_instance_id in self.active_workline_data:
            del self.active_workline_data[exited_instance_id]

        if self.tree.exists(exited_instance_id):
            current_values = list(self.tree.item(exited_instance_id, 'values'))
            if current_values[2] not in ["成功", "失败"]:
                current_values[2] = "已停止"
                current_values[3] = "进程退出"
            self.tree.item(exited_instance_id, values=tuple(current_values))
            self.draw_buttons_for_item(exited_instance_id)
        
        if exited_instance_id not in list(self.available_worklines_queue.queue):
            self.available_worklines_queue.put(exited_instance_id)
        
        self.check_if_batch_is_complete()

    def check_if_batch_is_complete(self):
        if self.batch_complete_sound_played:
            return

        try:
            with open(self.original_data_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                remaining_data = len(data)
        except (FileNotFoundError, json.JSONDecodeError):
            remaining_data = 0
        
        idle_worklines = self.available_worklines_queue.qsize()

        self.log(f"[GUI-检查] 剩余数据: {remaining_data}, 空闲工作线: {idle_worklines}, 总工作线: {self.total_worklines_started}")

        if remaining_data == 0 and idle_worklines == self.total_worklines_started and self.total_worklines_started > 0:
            self.log("[GUI-批次完成] 所有任务已完成。")
            self._play_sound_non_blocking("batch_complete")
            self.batch_complete_sound_played = True

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column_id_str = self.tree.identify_column(event.x)
            column_id = int(column_id_str.replace("#", ""))
            item_id = self.tree.identify_row(event.y)

            if not item_id: return
            
            if column_id == 7:
                self.toggle_pause(item_id)
            elif column_id == 8:
                self.show_manage_browser_dialog(item_id)
            elif column_id == 9:
                self.request_screenshot(item_id)

    def draw_buttons_for_item(self, item_id):
        current_values = list(self.tree.item(item_id, 'values'))
        if not current_values or not current_values[0]: return

        status = current_values[2]

        if status in ["启动中", "运行中", "暂停中", "排队中"]:
            if item_id in self.pause_states and self.pause_states[item_id] == '暂停中':
                current_values[6] = "恢复"
            else:
                current_values[6] = "暂停"
        elif status in ["成功", "失败", "已完成", "已停止", "进程退出"]:
            current_values[6] = "重新启动"
        else:
            current_values[6] = "暂停"

        current_values[7] = "管理"
        current_values[8] = "截图"

        self.tree.item(item_id, values=tuple(current_values))
    
    def request_screenshot(self, instance_id):
        process = self.node_processes.get(instance_id)
        if not process or process.poll() is not None:
            messagebox.showwarning("警告", f"工作线 {instance_id} 未在运行中，无法截图。")
            return

        command = f"SCREENSHOT::{instance_id}\n"
        try:
            process.stdin.write(command)
            process.stdin.flush()
            self.log(f"[GUI] 已向 {instance_id} 发送截图命令。", instance_id=instance_id)
            messagebox.showinfo("截图", f"已向 {instance_id} 发送截图命令。\n截图将保存在 'screenshot' 文件夹中。")
        except Exception as e:
            self.log(f"[GUI-ERROR] 发送截图命令失败: {e}", instance_id=instance_id)
            messagebox.showerror("错误", f"向 {instance_id} 发送截图命令失败: {e}")

    def toggle_pause(self, instance_id):
        if self.tree.set(instance_id, "action") == "重新启动":
            self.log(f"[GUI] 收到命令: 重新启动 {instance_id}...", instance_id)
            if instance_id not in list(self.available_worklines_queue.queue):
                self.available_worklines_queue.put(instance_id)
            
            current_values = list(self.tree.item(instance_id, 'values'))
            current_values[2] = "排队中"
            current_values[3] = "等待任务分配"
            self.tree.item(instance_id, values=tuple(current_values))
            self.draw_buttons_for_item(instance_id)
            return

        process = self.node_processes.get(instance_id)
        if not process or process.poll() is not None:
            return messagebox.showwarning("警告", f"工作线 {instance_id} 未在运行中。")

        current_state = self.pause_states.get(instance_id, '运行中')
        new_state = '暂停中' if current_state == '运行中' else '运行中'
        command_prefix = "PAUSE" if new_state == '暂停中' else "RESUME"
        command = f"{command_prefix}::{instance_id}\n"
        try:
            process.stdin.write(command)
            process.stdin.flush()
            self.pause_states[instance_id] = new_state
            
            current_values = list(self.tree.item(instance_id, 'values'))
            current_values[2] = new_state
            current_values[3] = "用户手动操作" if new_state == '暂停中' else "已恢复运行"
            self.tree.item(instance_id, values=tuple(current_values))

            self.log(f"[GUI] 已发送命令: {command.strip()}", instance_id=instance_id)
            self.draw_buttons_for_item(instance_id)
        except Exception as e:
            self.log(f"[GUI-ERROR] 发送命令失败: {e}", instance_id=instance_id)

    def _center_dialog(self, dialog):
        dialog.update_idletasks()
        main_width, main_height = self.root.winfo_width(), self.root.winfo_height()
        main_x, main_y = self.root.winfo_x(), self.root.winfo_y()
        dialog_width, dialog_height = dialog.winfo_width(), dialog.winfo_height()
        x = main_x + (main_width // 2) - (dialog_width // 2)
        y = main_y + (main_height // 2) - (dialog_height // 2)
        dialog.geometry(f"+{x}+{y}")

    def show_manage_browser_dialog(self, instance_id):
        dialog = tk.Toplevel(self.root)
        dialog.title(f"管理工作线 {instance_id} 浏览器")
        dialog.geometry("400x180")
        dialog.transient(self.root)
        dialog.grab_set()

        self._center_dialog(dialog)

        ttk.Label(dialog, text=f"确定要关闭并删除工作线 {instance_id} 的浏览器吗？", wraplength=350).pack(pady=15)
        
        btn_yes_return_card = ttk.Button(dialog, text="是的，并返回卡片信息到txt", 
                                         command=lambda: self._handle_browser_management(instance_id, True, dialog))
        btn_yes_return_card.pack(pady=5)
        
        btn_yes_no_return = ttk.Button(dialog, text="是的，但不需要返回卡片信息到txt", 
                                       command=lambda: self._handle_browser_management(instance_id, False, dialog))
        btn_yes_no_return.pack(pady=5)
        
        btn_no_cancel = ttk.Button(dialog, text="不，我点错了", command=dialog.destroy)
        btn_no_cancel.pack(pady=5)
        
        self.root.wait_window(dialog)

    def _save_card_info_to_txt(self, task_data, instance_id):
        if task_data:
            try:
                card_info = [
                    task_data.get('1step_number', ''),
                    f"{task_data.get('1step_month', '')}/{task_data.get('1step_year', '')}",
                    task_data.get('1step_code', ''),
                    task_data.get('real_name', '')
                ]
                info_line = '|'.join(card_info) + '\n'
                os.makedirs(os.path.dirname(self.not_used_cards_path), exist_ok=True)
                with open(self.not_used_cards_path, 'a', encoding='utf-8') as f:
                    f.write(info_line)
                self.log(f"[GUI-管理] 已将卡片信息保存到 {os.path.basename(self.not_used_cards_path)}。", instance_id)
            except Exception as e:
                self.log(f"[GUI-管理-错误] 保存卡片信息失败: {e}", instance_id)
        else:
            self.log(f"[GUI-管理-WARN] 未找到 {instance_id} 的活跃任务数据。", instance_id)

    def _get_browser_id_by_port(self, port):
        list_browsers_api_url = f"{self.BIT_API_BASE_URL}/browser/list"
        page = 0
        while True:
            try:
                payload = {"page": page, "pageSize": 100}
                response = requests.post(list_browsers_api_url, json=payload, timeout=10)
                response.raise_for_status()
                list_data = response.json()
                
                if list_data.get('success') and list_data.get('data') and list_data['data'].get('list'):
                    browser_list = list_data['data']['list']
                    for browser_info in browser_list:
                        if browser_info.get('proxyMethod') == 2 and \
                           browser_info.get('host') == '127.0.0.1' and \
                           str(browser_info.get('port')) == str(port):
                            self.log(f"[GUI-管理] 在端口 {port} 找到浏览器ID: {browser_info.get('id')}")
                            return browser_info.get('id')
                    
                    if len(browser_list) == 100:
                        page += 1
                        continue
                    else:
                        break
                else:
                    self.log(f"[GUI-管理-WARN] 调用列表API未返回成功或数据为空。")
                    break
            except requests.RequestException as e:
                self.log(f"[GUI-管理-ERROR] 调用BitBrowser列表API失败: {e}")
                return None
        
        self.log(f"[GUI-管理-WARN] 遍历完所有窗口，未找到端口 {port} 对应的浏览器ID。")
        return None


    def _close_and_delete_browser_via_bitapi(self, browser_id, port, instance_id=None):
        if not browser_id:
            self.log(f"[GUI-管理-WARN] 浏览器ID缺失 (端口: {port})，无法执行删除。", instance_id)
            return False

        close_url = f"{self.BIT_API_BASE_URL}/browser/close"
        delete_url = f"{self.BIT_API_BASE_URL}/browser/delete"

        try:
            self.log(f"[GUI-管理] 正在关闭浏览器 (ID: {browser_id})...", instance_id)
            requests.post(close_url, json={"id": browser_id}, timeout=10)
        except requests.RequestException as e:
            self.log(f"[GUI-管理-WARN] 调用关闭API时出错 (可能窗口已关闭): {e}", instance_id)
        
        time.sleep(1)

        try:
            self.log(f"[GUI-管理] 正在删除浏览器 (ID: {browser_id})...", instance_id)
            delete_response = requests.post(delete_url, json={"id": browser_id}, timeout=20)
            delete_response.raise_for_status()
            if delete_response.json().get('success'):
                self.log(f"[GUI-管理] 浏览器 (ID: {browser_id}) 删除成功。", instance_id)
                return True
            else:
                self.log(f"[GUI-管理-ERROR] 删除API返回失败: {delete_response.text}", instance_id)
                return False
        except requests.RequestException as e:
            self.log(f"[GUI-管理-ERROR] 调用删除API失败: {e}", instance_id)
            return False

    def _handle_browser_management(self, instance_id, return_card_info, dialog):
        dialog.destroy()
        self.log(f"[GUI-管理] 正在处理 {instance_id} 浏览器删除请求...")
        
        proxy_port = self.workline_ports.get(instance_id)
        if not proxy_port:
            self.log(f"[GUI-管理-WARN] 未找到 {instance_id} 的端口信息。", instance_id)
            return

        if instance_id in self.node_processes and self.node_processes[instance_id].poll() is None:
            process = self.node_processes[instance_id]
            try:
                process.terminate()
                process.wait(timeout=5)
            except Exception:
                process.kill()
            del self.node_processes[instance_id]
        
        if return_card_info:
            task_data = self.active_workline_data.get(instance_id)
            self._save_card_info_to_txt(task_data, instance_id)
        
        browser_id = self._get_browser_id_by_port(proxy_port)
        self._close_and_delete_browser_via_bitapi(browser_id, proxy_port, instance_id)

        if instance_id in self.active_workline_data:
            del self.active_workline_data[instance_id]
            
        if instance_id not in list(self.available_worklines_queue.queue):
            self.available_worklines_queue.put(instance_id)
        self.log(f"[GUI-管理] 工作线 {instance_id} 已恢复空闲。", instance_id)

        if self.tree.exists(instance_id):
            current_values = list(self.tree.item(instance_id, 'values'))
            current_values[2] = "空闲"
            current_values[3] = "已手动清理"
            self.tree.item(instance_id, values=tuple(current_values))
            self.draw_buttons_for_item(instance_id)

    def show_close_all_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("关闭并删除所有窗口")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        self._center_dialog(dialog)

        ttk.Label(dialog, text="警告: 这将关闭并删除所有运行中的浏览器。确定继续吗？", wraplength=350).pack(pady=15)
        
        btn_yes_return_card = ttk.Button(dialog, text="是的，并返回所有卡片信息到txt", 
                                         command=lambda: self._handle_close_all_browsers(True, dialog))
        btn_yes_return_card.pack(pady=5)
        
        btn_yes_no_return = ttk.Button(dialog, text="是的，但不需要返回卡片信息到txt", 
                                       command=lambda: self._handle_close_all_browsers(False, dialog))
        btn_yes_no_return.pack(pady=5)
        
        btn_no_cancel = ttk.Button(dialog, text="不，我点错了", command=dialog.destroy)
        btn_no_cancel.pack(pady=5)

        self.root.wait_window(dialog)

    def _handle_close_all_browsers(self, return_card_info, dialog):
        dialog.destroy()
        self.log("[GUI-管理] 正在关闭并删除所有窗口...")

        self.dealer_thread_stop_event.set()
        
        for instance_id, process in list(self.node_processes.items()):
            if process.poll() is None:
                try:
                    process.terminate()
                    process.wait(timeout=2)
                except Exception:
                    process.kill()
        
        if return_card_info:
            for instance_id, task_data in list(self.active_workline_data.items()):
                self._save_card_info_to_txt(task_data, instance_id)
        
        for instance_id_iter, port in list(self.workline_ports.items()):
            browser_id = self._get_browser_id_by_port(port)
            if browser_id:
                self._close_and_delete_browser_via_bitapi(browser_id, port, instance_id_iter)

        self.node_processes.clear()
        self.active_workline_data.clear()
        self.available_worklines_queue = queue.Queue()
        self.tree.delete(*self.tree.get_children())
        
        for i in list(self.notebook.tabs()):
            if self.notebook.tab(i, "text") != "全部日志":
                self.notebook.forget(i)
        self.log_tabs.clear()

        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.close_all_button.config(state="disabled")
        
        self.log("[GUI] 所有窗口和进程已清理。", instance_id="全部")
        self._save_workline_stats()

    def on_closing(self):
        self.save_settings()
        if messagebox.askokcancel("退出", "确定要退出并停止所有自动化脚本吗？"):
            self.stop_all_automation()
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AwsAutomationApp(root)
    root.mainloop()
