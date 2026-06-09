import json
import os
import re
import socket
import threading
import tkinter as tk
import gc
import sys
import qrcode
from PIL import Image, ImageTk
from datetime import datetime
from tkinter import (
    Tk, ttk, messagebox, scrolledtext,
    Toplevel, simpledialog, Listbox,
    StringVar, BooleanVar, IntVar, END
)
import pandas as pd
from flask import Flask, render_template, request, redirect
from waitress import serve

# ==============================================
# 兼容脚本/EXE的根目录（打包专用）
# ==============================================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path):
    base_path = os.path.dirname(os.path.abspath(__file__))
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    return os.path.join(base_path, relative_path)

DATA_PATH = os.path.join(BASE_DIR, "vote_backup.json")
HOLDING_FILE_PATH = os.path.join(BASE_DIR, "持股明细文件.xlsx")

# ==============================================
# Flask 服务常量配置
# ==============================================
FLASK_HOST = "0.0.0.0"
FLASK_PORT = 5000
FLASK_THREADS_BASE = 10 # 线程数基础值，实际线程数 = 选举人数 + 该值

# ==============================================
# 全局常量
# ==============================================
UI_STYLE = {
    "font": ("微软雅黑", 9),
    "title_font": ("微软雅黑", 10, "bold"),
    "accent": "#42C6A5",
    "warning": "#FF6B35",
    "info": "#42C6A5",
    "window_size": "1000x550",
}

DEFAULT_CONFIG = {
    "title": "现场投票系统",
    "default_candidates": [],
    "voters": [],
    "show_extra_candidates": False,
    "vote_mode": "percent",
    "server_running": False,
    "allow_abstain": False      # 是否允许弃票（一人一票：允许不足额；百分比：允许总分<100%）
}

MSG = {
    "no_data": "暂无投票数据！",
    "no_candidate": "暂无候选人！",
    "reset_confirm": "确定要清空投票数据吗？\n复位后可进行第二次投票！",
    "reset_success": "投票数据已全部复位！\n候选人/选举人配置保留可直接二次投票",
    "input_error": "请输入正整数！",
    "input_range": "当选人数必须在 1 ~ {} 之间！",
    "reset_need_stop_server": "复位投票数据前需要将服务器关闭",
    "input_threshold": "请输入有效票最低分数(%)",
    "threshold_error": "有效票分数必须为大于0的数字",
    "extra_candidate_zero": "额外候选人分数不能为0%，请删除或修改为大于0的数值！",
}

# ==============================================
# 数据管理层
# ==============================================
class VoteDataManager:
    def __init__(self):
        self.config = DEFAULT_CONFIG.copy()
        self.vote_records = []
        self.voted_users = set()
        self._file_lock = threading.Lock()
        self.shareholder_shares = {}
        self.total_shares = 0
        self.load_share_data() 
        self.load_from_file()
        try:
            self.iconbitmap(resource_path("logo.ico"))
        except Exception:
            pass

    def load_from_file(self) -> None:
        if not os.path.exists(DATA_PATH):
            return
        try:
            with open(DATA_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.vote_records = data.get("vote_data", [])
                self.voted_users = set(data.get("submit_cache", []))
                self.config.update(data.get("config", {}))
        except Exception:
            pass

    def save_to_file(self) -> None:
        try:
            with self._file_lock:   # 加锁
                with open(DATA_PATH, "w", encoding="utf-8") as f:
                    json.dump({
                        "vote_data": self.vote_records,
                        "submit_cache": list(self.voted_users),
                        "config": self.config
                    }, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def reset_vote_data(self) -> None:
        self.vote_records.clear()
        self.voted_users.clear()
        self.save_to_file()

    def load_share_data(self):
        """加载持股文件数据到缓存"""
        if not os.path.exists(HOLDING_FILE_PATH):
            print("持股文件不存在")
            return
        try:
            df = pd.read_excel(HOLDING_FILE_PATH)
            df.columns = [c.strip() for c in df.columns]
            name_col = None
            share_col = None
            for col in df.columns:
                if "股东" in col or "名称" in col:
                    name_col = col
                if "持股" in col or "股份" in col or "数" in col:
                    share_col = col
            if name_col is None or share_col is None:
                print("持股文件列名错误，找不到股东名称或持股数列")
                return
            df["股东名称"] = df[name_col].astype(str).str.strip()
            df["持股数"] = df[share_col].astype(str).apply(
                lambda x: float(re.sub(r"[^-0-9.]", "", x)) if re.sub(r"[^-0-9.]", "", x) else 0.0
            )
            self.shareholder_shares = dict(zip(df["股东名称"], df["持股数"]))
            self.total_shares = sum(self.shareholder_shares.values())
        except Exception as e:
            print(f"加载持股文件失败：{e}")

data_manager = VoteDataManager()

# ==============================================
# Flask 投票服务层
# ==============================================
class VoteServer:
    app = Flask(__name__, template_folder=os.path.join(os.path.dirname(__file__), 'templates'))
    app.config["ASSETS_DEBUG"] = False
    server_running = False
    elected_limit = 0
    valid_threshold = 50

    @staticmethod
    def start_server():
        voter_count = len(data_manager.config.get("voters", []))
        thread_pool_size = voter_count + FLASK_THREADS_BASE
        if thread_pool_size < 1:
            thread_pool_size = FLASK_THREADS_BASE
        VoteServer.server_running = True
        serve(VoteServer.app, host=FLASK_HOST, port=FLASK_PORT, threads=thread_pool_size)

    @staticmethod
    def stop():
        VoteServer.server_running = False

    @staticmethod
    @app.route("/signin", methods=["GET", "POST"])
    def signin():
        signin_file = os.path.join(BASE_DIR, "signin_records.json")
        signin_data = {}
        if os.path.exists(signin_file):
            try:
                with open(signin_file, "r", encoding="utf-8") as f:
                    signin_data = json.load(f)
            except:
                pass

        if request.method == "POST":
            name = request.form.get("name", "").strip()
            id_card = request.form.get("id_card", "").strip().upper()
            # 身份证号简单校验（18位，最后可为数字或X）
            if not name or not re.match(r'^[1-9]\d{16}[\dX]$', id_card):
                return render_template("signin.html", error="请输入正确的姓名和18位身份证号")
            key = f"{name}{id_card}"
            if key in signin_data:
                return render_template("signin.html", error="您已完成签到，请勿重复操作！")
            signin_data[key] = {
                "name": name,
                "id_card": id_card,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            with data_manager._file_lock:
                with open(signin_file, "w", encoding="utf-8") as f:
                    json.dump(signin_data, f, ensure_ascii=False, indent=2)
            return render_template("signin_success.html", name=name)

        return render_template("signin.html", error=None)
    
    @staticmethod
    @app.route("/")
    def index():
        return render_template("index.html", title=data_manager.config["title"])

    @staticmethod
    @app.route("/vote")
    def vote_page():
        voter_name = request.args.get("voter_name", "").strip()
        if not voter_name:
            return redirect("/")
        
        config = data_manager.config
        if voter_name not in config["voters"]:
            return render_template("no_permission.html")
        
        if voter_name in data_manager.voted_users:
            return render_template("duplicate_vote.html")
        
        # 读取持股文件，获取当前用户的持股数
        user_share = int(data_manager.shareholder_shares.get(voter_name, 0))   
        elected_limit = VoteServer.elected_limit
        vote_mode = config["vote_mode"]
        allow_abstain = config.get("allow_abstain", False)
        show_extra = config["show_extra_candidates"]
        candidates = config["default_candidates"]
        extra_candidate_zero = MSG["extra_candidate_zero"]

        return render_template("vote.html",
            voter_name=voter_name,
            elected_limit=elected_limit,
            vote_mode=vote_mode,
            allow_abstain=allow_abstain,
            candidates=candidates,
            show_extra=show_extra,
            extra_candidate_zero=extra_candidate_zero,
            user_share=user_share)
    
    @staticmethod
    @app.route("/submit", methods=["POST"])
    def submit_vote():
        
        try:
            voter_name = request.form.get("voter_name", "").strip()
            vote_mode = request.form.get("vote_mode", "").strip()
            allow_abstain = request.form.get("allow_abstain", "false").lower() == "true"
            client_ip = request.remote_addr
            total_share = 0
            if voter_name not in data_manager.config["voters"] or voter_name in data_manager.voted_users:
                return "<h2>系统异常</h2><br><a href='/'>返回首页</a>"
            record = {"voter": voter_name, "votes": {}, "valid_votes": {}, "ip": client_ip, "mode": vote_mode}
            user_share = int(data_manager.shareholder_shares.get(voter_name, 0))
            if vote_mode == "percent":
                voter_name = request.form.get("voter_name", "").strip()
                def safe_int(s):
                    try:
                        return int(s) if s.strip() != '' else 0
                    except (ValueError, AttributeError):
                        return 0

                default_shares = [safe_int(s) for s in request.form.getlist("default_share")]
                extra_shares = [safe_int(s) for s in request.form.getlist("extra_share")]
                extra_names = [n.strip() for n in request.form.getlist("extra_name") if n.strip()]
                # 获取持股数（直接从缓存获取）
                user_share = data_manager.shareholder_shares.get(voter_name, 0)                
                # 检查额外候选人股数不能为0
                for idx, name in enumerate(extra_names):
                    share = extra_shares[idx] if idx < len(extra_shares) else 0
                    if share == 0:
                        return f"<h2>投票失败：{MSG['extra_candidate_zero']}</h2><br><a href='/'>返回首页</a>"
                
                total_share = sum(default_shares) + sum(extra_shares)
                
                # 总股数校验
                if total_share > user_share:
                    return f"<h2>投票失败：分配股数总和({total_share})超过您的持股数({user_share})！</h2><br><a href='/'>返回首页</a>"
                if not allow_abstain and total_share != user_share:
                    return f"<h2>投票失败：不允许弃票时，被分配股数的候选人必须等于当选人是({user_share})，当前总和{total_share}</h2><br><a href='/'>返回首页</a>"
                
                # 非零人数校验
                elected_limit = VoteServer.elected_limit
                all_shares = []
                for idx, name in enumerate(data_manager.config["default_candidates"]):
                    share = default_shares[idx] if idx < len(default_shares) else 0
                    all_shares.append(share)
                for idx, name in enumerate(extra_names):
                    share = extra_shares[idx] if idx < len(extra_shares) else 0
                    all_shares.append(share)
                non_zero_count = sum(1 for s in all_shares if s > 0)
                if non_zero_count > elected_limit:
                    return f"<h2>投票失败：获得非零股数的候选人数量({non_zero_count})超过允许的最大值({elected_limit})</h2><br><a href='/'>返回首页</a>"
                if not allow_abstain and non_zero_count != elected_limit:
                    return f"<h2>投票失败：不允许弃票时，必须恰好有{elected_limit}位候选人获得非零股数（当前{non_zero_count}位）</h2><br><a href='/'>返回首页</a>"
                
                # 记录投票
                for idx, name in enumerate(data_manager.config["default_candidates"]):
                    record["votes"][name] = default_shares[idx] if idx < len(default_shares) else 0
                for idx, name in enumerate(extra_names):
                    record["votes"][name] = extra_shares[idx] if idx < len(extra_shares) else 0
            else:
                # 一人一票模式
                elected_limit = VoteServer.elected_limit
                selected_count = 0
                for name in data_manager.config["default_candidates"]:
                    if name in request.form.getlist("vote_target"):
                        selected_count += 1
                extra_names = [n.strip() for n in request.form.getlist("extra_name") if n.strip()]
                extra_targets = request.form.getlist("extra_target")
                for idx, name in enumerate(extra_names):
                    if idx < len(extra_targets) and extra_targets[idx] == '1':
                        selected_count += 1
                if not allow_abstain and selected_count != elected_limit:
                    return f"<h2>投票失败：不允许弃票时，必须恰好选择{elected_limit}人！当前选择{selected_count}人。</h2><br><a href='/'>返回首页</a>"
                if selected_count > elected_limit:
                    return f"<h2>投票失败：选择人数({selected_count})超过当选人数({elected_limit})</h2><br><a href='/'>返回首页</a>"
                for name in data_manager.config["default_candidates"]:
                    record["votes"][name] = 1 if name in request.form.getlist("vote_target") else 0
                    record["valid_votes"][name] = "赞成票" if record["votes"][name] == 1 else "未选择"
                for idx, name in enumerate(extra_names):
                    val = 1 if idx < len(extra_targets) and extra_targets[idx] == '1' else 0
                    record["votes"][name] = val
                    record["valid_votes"][name] = "赞成票" if val == 1 else "未选择"
            data_manager.vote_records.append(record)
            data_manager.voted_users.add(voter_name)
            data_manager.save_to_file()
            return render_template("vote_success.html")
        except Exception as e:
            return f"<h2>投票失败：{str(e)}</h2><br><a href='/'>返回首页</a>"

# ==============================================
# GUI 界面层
# ==============================================
class VoteSystemGUI:
    LOG_REFRESH_INTERVAL = 2000   # 2秒刷新一次日志
    def __init__(self, root: Tk):
        self.root = root
        self.server_thread = None
        self.running = True
        self.after_id = None
        self.log_paused = False
        self.elected_count = None
        self.icon_path = self.get_icon_path()
        self.setup_window()
        self.root.iconbitmap(resource_path("logo.ico"))
        self.setup_ui()
        self.refresh_ui()
        self.start_log_loop()

    def get_icon_path(self):
        try:
            ico_files = [f for f in os.listdir(BASE_DIR) if f.lower().endswith(".ico")]
            if ico_files:
                return os.path.join(BASE_DIR, ico_files[0])
        except:
            pass
        return None

    def set_window_icon(self, window):
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                window.iconbitmap(self.icon_path)
            except:
                pass
    def show_signin_pool(self):
        """显示签到池，支持自动滚动播放（可调速、启停）"""
        import tkinter as tk
        from tkinter import ttk

        signin_file = os.path.join(BASE_DIR, "signin_records.json")
        current_font_size = 12
        scroll_interval = 500      # 初始滚动间隔（毫秒）
        scroll_job = None          # 用于存储 after 任务的ID
        is_scrolling = False       # 滚动状态

        top = Toplevel(self.root)
        top.title("签到池 - 自动滚动模式 (管理投屏)")
        top.geometry("1000x650")
        top.resizable(True, True)
        self.set_window_icon(top)

        # 主框架 grid 布局
        main_frame = ttk.Frame(top)
        main_frame.pack(fill="both", expand=True)
        main_frame.grid_rowconfigure(0, weight=0)  # 标题
        main_frame.grid_rowconfigure(1, weight=0)  # 统计
        main_frame.grid_rowconfigure(2, weight=1)  # 表格区
        main_frame.grid_rowconfigure(3, weight=0)  # 按钮控制栏
        main_frame.grid_columnconfigure(0, weight=1)

        # 标题
        title_label = ttk.Label(main_frame, text="📋 签到池 (自动滚动)", font=("微软雅黑", 26, "bold"))
        title_label.grid(row=0, column=0, pady=10)

        # 统计信息
        stat_label = ttk.Label(main_frame, text="", font=("微软雅黑", 14))
        stat_label.grid(row=1, column=0, pady=5)

        # 表格区域
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=10)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = ("姓名", "身份证号(部分隐藏)", "签到时间")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        for col in columns:
            tree.heading(col, text=col)
            if col == "签到时间":
                width = 250
            elif col == "身份证号(部分隐藏)":
                width = 220   # 适当加宽以容纳掩码字符串
            else:
                width = 200
            tree.column(col, width=width, anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        tree.tag_configure("oddrow", background="#F8F8F8")
        tree.tag_configure("evenrow", background="#FFFFFF")

        # ========== 滚动控制函数 ==========
        def start_scroll():
            nonlocal is_scrolling, scroll_job
            if is_scrolling:
                return
            is_scrolling = True
            # 禁用开始按钮，启用停止按钮
            btn_start.config(state="disabled")
            btn_stop.config(state="normal")
            # 开始滚动任务
            scroll_step()

        def stop_scroll():
            nonlocal is_scrolling, scroll_job
            is_scrolling = False
            if scroll_job:
                top.after_cancel(scroll_job)
                scroll_job = None
            btn_start.config(state="normal")
            btn_stop.config(state="disabled")

        def scroll_step():
            nonlocal scroll_job
            if not is_scrolling:
                return
            # 滚动一行（如果已到底部，则跳回顶部）
            if tree.yview()[1] >= 1.0:
                # 到达底部，回到顶部
                tree.yview_moveto(0)
            else:
                tree.yview_scroll(1, "units")
            # 继续下一次滚动
            scroll_job = top.after(scroll_interval, scroll_step)

        def set_speed(val):
            nonlocal scroll_interval
            scroll_interval = int(float(val))
            # 如果正在滚动，重新启动以应用新间隔
            if is_scrolling:
                stop_scroll()
                start_scroll()

        # ========== 字体缩放函数 ==========
        def apply_content_font():
            tree.tag_configure("oddrow", font=("微软雅黑", current_font_size))
            tree.tag_configure("evenrow", font=("微软雅黑", current_font_size))

        def zoom_in():
            nonlocal current_font_size
            if current_font_size < 30:
                current_font_size += 2
                apply_content_font()
                ttk.Style().configure("Treeview", rowheight=current_font_size + 15)

        def zoom_out():
            nonlocal current_font_size
            if current_font_size > 10:
                current_font_size -= 2
                apply_content_font()
                ttk.Style().configure("Treeview", rowheight=current_font_size + 15)

        # ========== 数据刷新 ==========
        def refresh_data():
            nonlocal current_font_size
            if not os.path.exists(signin_file):
                stat_label.config(text="暂无签到记录")
                for row in tree.get_children():
                    tree.delete(row)
                return
            try:
                with open(signin_file, "r", encoding="utf-8") as f:
                    signin_data = json.load(f)
            except Exception as e:
                messagebox.showerror("错误", f"读取签到记录失败：{e}")
                return
            for row in tree.get_children():
                tree.delete(row)
            if not signin_data:
                stat_label.config(text="暂无签到人员")
                return
            records = list(signin_data.values())
            records.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
            for i, info in enumerate(records):
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                # 身份证号掩码（假设已存储为 id_card）
                id_card = info.get("id_card", "")
                if len(id_card) == 18:
                    masked_id = id_card[:3] + "***********" + id_card[-4:]
                else:
                    masked_id = "无效身份证号"
                tree.insert("", "end", values=(info["name"], masked_id, info["timestamp"]), tags=(tag,))
            stat_label.config(text=f"总签到人数：{len(records)}")
            apply_content_font()
            ttk.Style().configure("Treeview", rowheight=current_font_size + 15)
            stop_scroll()

        # ========== 底部按钮栏 ==========
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, sticky="ew", pady=10)
        btn_frame.grid_columnconfigure(0, weight=1)

        # 字体缩放
        zoom_frame = ttk.Frame(btn_frame)
        zoom_frame.pack(side="left", padx=5)
        ttk.Button(zoom_frame, text="A+", command=zoom_in, width=3).pack(side="left", padx=2)
        ttk.Button(zoom_frame, text="A-", command=zoom_out, width=3).pack(side="left", padx=2)

        # 滚动控制按钮
        btn_start = ttk.Button(btn_frame, text="▶ 开始滚动", command=start_scroll, width=12)
        btn_start.pack(side="left", padx=5)
        btn_stop = ttk.Button(btn_frame, text="⏸ 停止滚动", command=stop_scroll, width=12, state="disabled")
        btn_stop.pack(side="left", padx=5)

        # 速度调节滑块
        speed_frame = ttk.Frame(btn_frame)
        speed_frame.pack(side="left", padx=10)
        ttk.Label(speed_frame, text="滚动速度(ms):").pack(side="left")
        speed_scale = ttk.Scale(speed_frame, from_=200, to=2000, orient="horizontal", length=150, command=set_speed)
        speed_scale.set(scroll_interval)
        speed_scale.pack(side="left", padx=5)
        speed_label = ttk.Label(speed_frame, text=f"{scroll_interval}")
        speed_label.pack(side="left")
        # 动态显示当前速度
        def on_speed_change(val):
            speed_label.config(text=str(int(float(val))))
            set_speed(val)
        speed_scale.configure(command=on_speed_change)

        # 刷新按钮
        refresh_btn = ttk.Button(btn_frame, text="🔄 刷新数据", command=refresh_data, style="Accent.TButton")
        refresh_btn.pack(side="left", padx=5)

        # 关闭按钮
        close_btn = ttk.Button(btn_frame, text="关闭窗口", command=top.destroy, width=15)
        close_btn.pack(side="right", padx=5)

        # 初始化数据
        refresh_data()

    def setup_window(self) -> None:
        self.root.title(f'{data_manager.config["title"]} - 投票管理系统')
        self.root.geometry(UI_STYLE["window_size"])
        self.root.resizable(True, True)
        self.set_window_icon(self.root)
        self.root.protocol("WM_DELETE_WINDOW", self.safe_close)
        style = ttk.Style()
        style.configure(".", font=UI_STYLE["font"])
        style.configure("Accent.TButton", font=UI_STYLE["title_font"], foreground=UI_STYLE["accent"])
        style.configure("Warning.TButton", font=UI_STYLE["title_font"], foreground=UI_STYLE["warning"])

    def setup_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)
        left_frame = ttk.LabelFrame(main_frame, text="📋 投票基础配置", padding=10)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        ttk.Label(left_frame, text="投票标题：", font=UI_STYLE["title_font"]).pack(anchor="w", pady=(0, 5))
        self.title_entry = ttk.Entry(left_frame, width=30)
        self.title_entry.insert(0, data_manager.config["title"])
        self.title_entry.pack(fill="x", pady=(0, 10))
        ttk.Button(left_frame, text="更新标题", command=self.update_title).pack(fill="x", pady=(0, 10))
        ttk.Label(left_frame, text="候选人列表：", font=UI_STYLE["title_font"]).pack(anchor="w", pady=(0, 5))
        self.candidate_list = Listbox(left_frame, width=30, height=8)
        self.candidate_list.pack(fill="both", expand=True, pady=(0, 5))
        candidate_btn_frame = ttk.Frame(left_frame)
        candidate_btn_frame.pack(fill="x")
        ttk.Button(candidate_btn_frame, text="批量添加", command=self.add_candidates).pack(side="left", expand=True, fill="x")
        ttk.Button(candidate_btn_frame, text="删除选中", command=self.delete_candidate).pack(side="left", expand=True, fill="x")
        ttk.Button(candidate_btn_frame, text="清空", command=self.clear_candidates).pack(side="left", expand=True, fill="x")
        ttk.Label(left_frame, text="投票模式：", font=UI_STYLE["title_font"]).pack(anchor="w", pady=(15, 5))
        mode_frame = ttk.Frame(left_frame)
        mode_frame.pack(fill="x")
        self.mode_var = StringVar(value=data_manager.config["vote_mode"])
        ttk.Radiobutton(mode_frame, text="百分比投票", variable=self.mode_var, value="percent", command=self.update_mode).pack(side="left", expand=True)
        ttk.Radiobutton(mode_frame, text="一人一票", variable=self.mode_var, value="one_vote", command=self.update_mode).pack(side="left", expand=True)
        
        # 允许弃票复选框（样式参照允许添加额外候选人）
        self.allow_abstain_var = BooleanVar(value=data_manager.config.get("allow_abstain", False))
        ttk.Checkbutton(left_frame, text="允许弃票", variable=self.allow_abstain_var, command=self.update_allow_abstain).pack(anchor="w", pady=(5, 5))
        
        self.extra_var = BooleanVar(value=data_manager.config["show_extra_candidates"])
        ttk.Checkbutton(left_frame, text="允许添加额外候选人", variable=self.extra_var, command=self.update_config).pack(anchor="w", pady=5)

        center_frame = ttk.LabelFrame(main_frame, text="⚙️ 服务器与数据操作", padding=10)
        center_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        self.server_btn = ttk.Button(center_frame, text="启动服务器", command=self.toggle_server)
        self.server_btn.pack(fill="x", pady=(0, 10))
        ttk.Label(center_frame, text="服务器访问地址（可复制）：").pack(anchor="w")
        self.server_addr_entry = ttk.Entry(center_frame, font=("微软雅黑", 10), foreground="green")
        self.server_addr_entry.pack(fill="x", pady=(0, 15))
        self.server_addr_entry.insert(0, "未启动服务器")
        self.server_addr_entry.config(state="readonly")
        # 在 center_frame 中，原有 ttk.Button 附近添加
        ttk.Button(center_frame, text="📝 签到QR码", command=self.show_signin_qrcode, style="Accent.TButton").pack(fill="x", pady=5)
        ttk.Button(center_frame, text="📷 服务器地址QR码", command=self.show_server_qrcode, style="Accent.TButton").pack(fill="x", pady=5)
        ttk.Button(center_frame, text="🧮 计算投票结果", command=self.calculate_result, style="Accent.TButton").pack(fill="x", pady=5)
        ttk.Button(center_frame, text="🔄 复位投票数据", command=self.reset_vote_data, style="Warning.TButton").pack(fill="x", pady=5)

        right_frame = ttk.LabelFrame(main_frame, text="👥 选举人管理与日志", padding=10)
        right_frame.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        # 选举人管理与日志区域内的 voter_top_frame
        voter_top_frame = ttk.Frame(right_frame)
        voter_top_frame.pack(fill="both", expand=True, pady=(0, 10))

        # 第0行：标签（第0列和第3列不需要标签）
        ttk.Label(voter_top_frame, text="选举人列表", font=UI_STYLE["title_font"]).grid(row=0, column=1, padx=5)
        ttk.Label(voter_top_frame, text="未投票人员", font=UI_STYLE["title_font"]).grid(row=0, column=2, padx=5)

        # 第1列：垂直按钮框架（位于第0列）
        left_btn_frame = ttk.Frame(voter_top_frame)
        left_btn_frame.grid(row=1, column=0, rowspan=2, sticky="ns", padx=5)
        ttk.Button(left_btn_frame, text="批量添加", command=self.add_voters).pack(fill="x", pady=2)
        ttk.Button(left_btn_frame, text="删除选中", command=self.delete_voter).pack(fill="x", pady=2)
        ttk.Button(left_btn_frame, text="复制全部", command=self.copy_voters_list).pack(fill="x", pady=2)
        ttk.Button(left_btn_frame, text="清空", command=self.clear_voters).pack(fill="x", pady=2)
        ttk.Button(left_btn_frame, text="导入签到用户", command=self.import_signed_users).pack(fill="x", pady=2)
        ttk.Button(left_btn_frame, text="查看签到池", command=self.show_signin_pool).pack(fill="x", pady=2)
        # 第2列：选举人列表
        self.voter_list = Listbox(voter_top_frame, width=20, height=10)
        self.voter_list.grid(row=1, column=1, padx=5, sticky="nsew")

        # 第3列：未投票人员列表
        self.unfinished_list = Listbox(voter_top_frame, width=20, height=10)
        self.unfinished_list.grid(row=1, column=2, padx=5, sticky="nsew")

        # 第4列：刷新状态按钮（垂直居中）
        refresh_btn = ttk.Button(voter_top_frame, text="刷新状态", command=self.check_vote_status)
        refresh_btn.grid(row=1, column=3, padx=5, sticky="n")  # sticky="n" 靠上，也可用 "ns" 拉伸

        # 设置列权重和行权重
        voter_top_frame.grid_columnconfigure(1, weight=1)   # 选举人列表列可拉伸
        voter_top_frame.grid_columnconfigure(2, weight=1)   # 未投票列表列可拉伸
        voter_top_frame.grid_columnconfigure(3, weight=0)   # 刷新按钮列不拉伸
        voter_top_frame.grid_rowconfigure(1, weight=1)      # 第二行可拉伸，使列表高度自适应

        # 如果希望刷新按钮垂直居中，可以创建一个额外的Frame包装，简单起见保持顶部对齐
        ttk.Label(right_frame, text="实时投票记录", font=UI_STYLE["title_font"]).pack(anchor="w", pady=(0, 5))

        # 日志显示开关框架
        log_switch_frame = ttk.Frame(right_frame)
        log_switch_frame.pack(fill="x")
        self.log_var = IntVar(value=1)
        ttk.Radiobutton(log_switch_frame, text="显示日志", variable=self.log_var, value=1, command=self.toggle_log_display).pack(side="left")
        ttk.Radiobutton(log_switch_frame, text="隐藏日志", variable=self.log_var, value=0, command=self.toggle_log_display).pack(side="left")

        # 按钮框架
        self.log_button_frame = ttk.Frame(right_frame)
        self.pause_btn = ttk.Button(self.log_button_frame, text="暂停日志", command=self.toggle_pause_log)
        self.pause_btn.pack(side="left", padx=2)
        self.copy_btn = ttk.Button(self.log_button_frame, text="复制日志", command=self.copy_log)
        self.copy_btn.pack(side="left", padx=2)
        self.log_button_frame.pack(fill="x", pady=(5, 0))

        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(right_frame, width=50, height=15)
        self.log_text.pack(fill="both", expand=True, pady=(5, 0))
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_columnconfigure(2, weight=2)
        main_frame.grid_rowconfigure(0, weight=1)

    def update_allow_abstain(self):
        data_manager.config["allow_abstain"] = self.allow_abstain_var.get()
        data_manager.save_to_file()

    def toggle_log_display(self):
        if self.log_var.get() == 1:
            # 显示日志：立即刷新完整日志
            self.update_log_display()
            # 恢复按钮文字（根据暂停状态）
            if self.log_paused:
                self.pause_btn.config(text="恢复日志")
            else:
                self.pause_btn.config(text="暂停日志")
        else:
            # 隐藏日志：只显示简化的“第几票已完成”
            self.update_log_display_hidden()

    def toggle_pause_log(self):
        self.log_paused = not self.log_paused
        if self.log_paused:
            self.pause_btn.config(text="恢复日志")
        else:
            self.pause_btn.config(text="暂停日志")
            # 恢复后立即刷新一次（根据当前显示/隐藏状态）
            if self.log_var.get() == 1:
                self.update_log_display()
            else:
                self.update_log_display_hidden()

    def copy_log(self):
        """复制日志内容到剪贴板"""
        try:
            # 获取当前日志文本（不论是否隐藏，只要有内容）
            log_content = self.log_text.get(1.0, END).strip()
            if not log_content:
                messagebox.showinfo("提示", "没有可复制的日志内容")
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(log_content)
            messagebox.showinfo("成功", f"已复制 {len(log_content)} 字符到剪贴板")
        except Exception as e:
            messagebox.showerror("错误", f"复制失败：{str(e)}")
    def update_log_display_hidden(self):
        """隐藏日志模式：只显示每票的序号和投票人，不显示候选人明细"""
        self.log_text.delete(1.0, END)
        if not data_manager.vote_records:
            self.log_text.insert(END, "暂无投票数据")
        else:
            for i, record in enumerate(data_manager.vote_records):
                voter = record.get('voter', '未知')
                votes = record.get('votes', {})
                mode = record.get('mode', 'percent')
                self.log_text.insert(END, f"第{i+1}票 | \n")
                if mode == 'percent':
                    for candidate, value in votes.items():
                        self.log_text.insert(END, f"    {candidate}: {value}票\n")
                else:  # one_vote
                    for candidate, value in votes.items():
                        status = "选中" if value == 1 else "未选"
                        self.log_text.insert(END, f"    {candidate}: {status}\n")
                self.log_text.insert(END, "\n")
            ##self.log_text.insert(END, f"\n总投票人数：{len(data_manager.vote_records)}")
    
    def update_log_display(self):
        self.log_text.delete(1.0, END)
        if not data_manager.vote_records:
            self.log_text.insert(END, "暂无投票数据")
        else:
            for i, record in enumerate(data_manager.vote_records):
                voter = record.get('voter', '未知')
                votes = record.get('votes', {})
                mode = record.get('mode', 'percent')  # 默认为百分比模式（兼容旧数据）
                self.log_text.insert(END, f"第{i+1}票 | {voter}\n")
                if mode == 'percent':
                    for candidate, value in votes.items():
                        self.log_text.insert(END, f"    {candidate}: {value}票\n")
                else:  # one_vote
                    for candidate, value in votes.items():
                        status = "选中" if value == 1 else "未选"
                        self.log_text.insert(END, f"    {candidate}: {status}\n")
                self.log_text.insert(END, "\n")

    def start_log_loop(self) -> None:
        if self.running:
            if self.log_var.get() == 1 and not self.log_paused:
                self.update_log_display()
            elif self.log_var.get() == 0 and not self.log_paused:
                self.update_log_display_hidden()
            self.after_id = self.root.after(self.LOG_REFRESH_INTERVAL, self.start_log_loop)

    def safe_close(self):
        self.running = False
        if self.after_id:
            self.root.after_cancel(self.after_id)
        if VoteServer.server_running:
            VoteServer.stop()
            if self.server_thread and self.server_thread.is_alive():
                self.server_thread.join(timeout=2)
        data_manager.save_to_file()
        self.root.destroy()
        gc.collect()

    def toggle_server(self) -> None:
        config = data_manager.config
        if not config["server_running"]:
            if not config["default_candidates"]:
                messagebox.showwarning("提示", MSG["no_candidate"])
                return
            
            if config["vote_mode"] == "percent":
                threshold = simpledialog.askstring("设置有效票最低分数", MSG["input_threshold"])
                if threshold is None:
                    return
                if not threshold.isdigit() or int(threshold) <= 0:
                    messagebox.showerror("错误", MSG["threshold_error"])
                    return
                VoteServer.valid_threshold = int(threshold)
                messagebox.showinfo("成功", f"已设置有效票最低分数：{threshold}%")
                
                total_candidates = len(config["default_candidates"])
                elected = self.get_elected_number(total_candidates)
                if elected is None:
                    return
                VoteServer.elected_limit = elected
                messagebox.showinfo("成功", f"已设置当选人数：{elected}人")
            else:
                total_candidates = len(config["default_candidates"])
                elected = self.get_elected_number(total_candidates)
                if elected is None:
                    return
                VoteServer.elected_limit = elected
                VoteServer.valid_threshold = 50

            config["server_running"] = True
            self.server_btn.config(text="停止服务器")
            ip = socket.gethostbyname(socket.gethostname())
            addr = f"http://{ip}:{FLASK_PORT}"
            
            self.server_addr_entry.config(state="normal")
            self.server_addr_entry.delete(0, END)
            self.server_addr_entry.insert(0, addr)
            self.server_addr_entry.config(state="readonly")
            
            self.server_thread = threading.Thread(target=VoteServer.start_server)
            self.server_thread.start()
            
            data_manager.config["server_running"] = True
            self.update_log_display()
            
        else:
            config["server_running"] = False
            VoteServer.stop()
            self.server_btn.config(text="启动服务器")
            self.server_addr_entry.config(state="normal")
            self.server_addr_entry.delete(0, END)
            self.server_addr_entry.insert(0, "未启动服务器")
            self.server_addr_entry.config(state="readonly")

    def get_elected_number(self, max_num: int) -> int | None:
        while True:
            res = simpledialog.askstring("设置当选人数", f"请输入当选人数\n(总候选人：{max_num}人，范围：1~{max_num})")
            if res is None:
                return None
            if not res.isdigit():
                messagebox.showerror("错误", MSG["input_error"])
                continue
            num = int(res)
            if 1 <= num <= max_num:
                return num
            messagebox.showerror("错误", MSG["input_range"].format(max_num))
    
    def show_server_qrcode(self):
        """生成服务器地址二维码并显示在弹窗中，窗口可缩放，二维码随窗口等比缩放居中"""
        addr = self.server_addr_entry.get().strip()
        if not addr or addr == "未启动服务器":
            messagebox.showwarning("提示", "请先启动服务器，再生成二维码！")
            return
        # 生成原始二维码图像（固定大小，用于缩放）
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
        qr.add_data(addr)
        qr.make(fit=True)
        original_img = qr.make_image(fill_color="black", back_color="white")

        top = Toplevel(self.root)
        top.title("服务器地址二维码")
        top.geometry("400x450")
        top.resizable(True, True)          # 允许窗口缩放
        self.set_window_icon(top)

        # 顶部显示地址文本
        addr_label = ttk.Label(top, text=addr, font=("微软雅黑", 10))
        addr_label.pack(pady=10)

        # 中间区域：用于放置二维码并自适应
        img_frame = ttk.Frame(top)
        img_frame.pack(fill="both", expand=True, padx=10, pady=10)

        img_label = ttk.Label(img_frame)
        img_label.pack(fill="both", expand=True)

        def resize_qrcode(event=None):
            """窗口大小改变时，等比缩放二维码"""
            frame_width = img_frame.winfo_width()
            frame_height = img_frame.winfo_height()
            if frame_width < 10 or frame_height < 10:
                return
            orig_w, orig_h = original_img.size
            scale = min(frame_width / orig_w, frame_height / orig_h)
            new_w = int(orig_w * scale)
            new_h = int(orig_h * scale)
            # Pillow 版本兼容
            try:
                resample = Image.Resampling.LANCZOS
            except AttributeError:
                resample = Image.ANTIALIAS
            resized = original_img.resize((new_w, new_h), resample)
            img_tk = ImageTk.PhotoImage(resized)
            img_label.config(image=img_tk)
            img_label.image = img_tk          # 保持引用，防止被垃圾回收

        # 绑定窗口大小改变事件
        top.bind("<Configure>", resize_qrcode)
        # 延迟一次初始化缩放（确保窗口布局完成）
        top.after(100, resize_qrcode)
    
    def show_signin_qrcode(self):
        """生成签到页面二维码并显示在弹窗中，窗口可缩放，二维码随窗口等比缩放居中"""
        addr = self.server_addr_entry.get().strip()
        if not addr or addr == "未启动服务器":
            messagebox.showwarning("提示", "请先启动服务器，再生成签到二维码！")
            return
        signin_url = addr + "/signin"
        # 生成原始二维码图像
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
        qr.add_data(signin_url)
        qr.make(fit=True)
        original_img = qr.make_image(fill_color="black", back_color="white")

        top = Toplevel(self.root)
        top.title("签到页面二维码")
        top.geometry("400x450")
        top.resizable(True, True)
        self.set_window_icon(top)

        # 顶部显示URL文本
        addr_label = ttk.Label(top, text=signin_url, font=("微软雅黑", 10))
        addr_label.pack(pady=10)

        # 中间区域：用于放置二维码并自适应
        img_frame = ttk.Frame(top)
        img_frame.pack(fill="both", expand=True, padx=10, pady=10)

        img_label = ttk.Label(img_frame)
        img_label.pack(fill="both", expand=True)

        def resize_qrcode(event=None):
            """窗口大小改变时，等比缩放二维码"""
            frame_width = img_frame.winfo_width()
            frame_height = img_frame.winfo_height()
            if frame_width < 10 or frame_height < 10:
                return
            orig_w, orig_h = original_img.size
            scale = min(frame_width / orig_w, frame_height / orig_h)
            new_w = int(orig_w * scale)
            new_h = int(orig_h * scale)
            try:
                resample = Image.Resampling.LANCZOS
            except AttributeError:
                resample = Image.ANTIALIAS
            resized = original_img.resize((new_w, new_h), resample)
            img_tk = ImageTk.PhotoImage(resized)
            img_label.config(image=img_tk)
            img_label.image = img_tk

        top.bind("<Configure>", resize_qrcode)
        top.after(100, resize_qrcode)    
    def calculate_result(self) -> None:
        records_snapshot = list(data_manager.vote_records)
        if not records_snapshot:
            messagebox.showwarning("提示", MSG["no_data"])
            return
        
        vote_mode = data_manager.config["vote_mode"]
        if vote_mode == "percent":
            if VoteServer.elected_limit == 0:
                messagebox.showwarning("提示", "请先启动服务器并设置有效票最低分数和当选人数！")
                return
            if not os.path.exists(HOLDING_FILE_PATH):
                messagebox.showerror("错误", "未找到持股明细文件！")
                return
            try:
                df = pd.read_excel(HOLDING_FILE_PATH)
                df.columns = [c.strip() for c in df.columns]
                name_col = None
                share_col = None
                for col in df.columns:
                    if "股东" in col or "名称" in col:
                        name_col = col
                    if "持股" in col or "股份" in col or "数" in col:
                        share_col = col
                if name_col is None or share_col is None:
                    messagebox.showerror("错误", "持股文件需包含‘股东名称’列和‘持股数’列")
                    return
                df["股东名称"] = df[name_col].astype(str).str.strip()
                df["持股数"] = df[share_col].astype(str).apply(
                    lambda x: float(re.sub(r"[^-0-9.]", "", x)) if re.sub(r"[^-0-9.]", "", x) else 0.0
                )
                shareholder_shares = dict(zip(df["股东名称"], df["持股数"]))
                total_shares = sum(shareholder_shares.values())
                if total_shares == 0:
                    messagebox.showerror("错误", "总持股数为0，无法计算")
                    return
                valid_threshold = VoteServer.valid_threshold
                baseline = total_shares * valid_threshold / 100.0
                elected_limit = VoteServer.elected_limit
                
                candidates = set()
                for rec in records_snapshot:
                    candidates.update(rec["votes"].keys())
                candidates = list(candidates)
                
                candidate_weighted_score = {c: 0.0 for c in candidates}
                for rec in records_snapshot:
                    voter_name = rec["voter"]
                    share = shareholder_shares.get(voter_name, 0)
                    if share == 0:
                        continue
                    votes = rec["votes"]
                    for cand, score in votes.items():
                        candidate_weighted_score[cand] += share * (score / 100.0)
                
                result = []
                for cand, score in candidate_weighted_score.items():
                    is_valid = score > baseline
                    result.append({
                        "候选人": cand,
                        "总得分": score,
                        "基准线": baseline,
                        "是否有效票": "是" if is_valid else "否",
                    })
                result.sort(key=lambda x: x["总得分"], reverse=True)
                rank = 1
                prev_score = None
                for i, item in enumerate(result):
                    if i > 0 and item["总得分"] != prev_score:
                        rank = i + 1
                    item["排名"] = rank
                    item["状态"] = "当选" if (i <= elected_limit) and (item["是否有效票"] == "是") else "未当选"
                    prev_score = item["总得分"]
                
                save_df = pd.DataFrame(result)
                save_df = save_df[["排名", "候选人", "总得分", "基准线", "是否有效票", "状态"]]
                clean_title = re.sub(r'[\\/:*?"<>|]', "", data_manager.config["title"].strip())
                time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                save_path = os.path.join(BASE_DIR, f"{clean_title}_结果_{time_str}.xlsx")
                save_df.to_excel(save_path, index=False)
                
                self.show_rank_window_percent(result, baseline, valid_threshold, elected_limit)
                messagebox.showinfo("成功", f"结果已保存：\n{save_path}")
                
            except Exception as e:
                messagebox.showerror("错误", f"计算失败：{str(e)}")
        else:
            if VoteServer.elected_limit == 0:
                messagebox.showwarning("提示", "请先启动服务器并设置当选人数！")
                return
            records_snapshot = list(data_manager.vote_records)
            candidates = list({k for r in records_snapshot for k in r["votes"].keys()})
            candidate_votes = {c: 0 for c in candidates}
            for rec in records_snapshot:
                for cand, share in rec["votes"].items():
                    candidate_votes[cand] += share
            sorted_items = sorted(candidate_votes.items(), key=lambda x: x[1], reverse=True)
            result = []
            rank = 1
            prev_votes = None
            for i, (name, vote_count) in enumerate(sorted_items):
                if i > 0 and vote_count != prev_votes:
                    rank = i + 1
                result.append({
                    "排名": rank, "候选人": name, "总得票数": vote_count,
                    "状态": "当选" if rank <= VoteServer.elected_limit else "未当选"
                })
                prev_votes = vote_count
            save_path = self.save_result_to_excel_one_vote(result)
            messagebox.showinfo("成功", f"结果已保存：\n{save_path}")
            self.show_rank_window_one_vote(result, VoteServer.elected_limit)

    def show_rank_window_one_vote(self, result, elected):
        top = Toplevel(self.root)
        top.title("🏆 投票结果排行榜 - 一人一票模式")
        top.geometry("900x600")
        top.resizable(True, True)
        self.set_window_icon(top)

        total_votes = len(data_manager.vote_records)
        total_cast = 0
        for rec in data_manager.vote_records:
            total_cast += sum(rec["votes"].values())
        avg_cast = total_cast / total_votes if total_votes > 0 else 0

        info_frame = ttk.Frame(top, padding=5)
        info_frame.pack(fill='x', pady=5)
        info_text = f"总投票人数：{total_votes}   |   总票数（人次）：{total_cast}   |   平均每人投出：{avg_cast:.2f}票   |   当选人数：{elected}"
        ttk.Label(info_frame, text=info_text, font=("微软雅黑", 10)).pack()

        tree_frame = ttk.Frame(top)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ("排名", "候选人", "总得票数", "得票率", "状态")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        col_widths = {"排名": 80, "候选人": 200, "总得票数": 120, "得票率": 120, "状态": 100}
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths[col], minwidth=60, anchor="center")

        # 计算得票率（占总投票人数的百分比）
        for i, item in enumerate(result):
            vote_count = item["总得票数"]
            rate = (vote_count / total_votes * 100) if total_votes > 0 else 0
            tags = "elected" if item["状态"] == "当选" else "normal"
            if i % 2 == 0:
                tags = (tags, "evenrow")
            else:
                tags = (tags, "oddrow")
            tree.insert("", "end", values=(
                item["排名"], item["候选人"], vote_count, f"{rate:.2f}%", item["状态"]
            ), tags=tags)

        tree.tag_configure("elected", background="#C6EFCE")
        tree.tag_configure("evenrow", background="#F8F8F8")
        tree.tag_configure("oddrow", background="#FFFFFF")

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        ttk.Button(top, text="确定关闭", command=top.destroy, width=15).pack(pady=10)

    def save_result_to_excel_one_vote(self, result: list) -> str:
        clean_title = re.sub(r'[\\/:*?"<>|]', "", data_manager.config["title"].strip())
        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(BASE_DIR, f"{clean_title}_{time_str}.xlsx")
        pd.DataFrame(result, columns=["排名","候选人","总得票数","状态"]).to_excel(path, index=False)
        return path


    def show_rank_window_percent(self, result, baseline, threshold, elected_limit):
        top = Toplevel(self.root)
        top.title("🏆 投票结果排行榜 - 百分比模式")
        top.geometry("1000x650")
        top.resizable(True, True)
        self.set_window_icon(top)

        # 统计信息
        total_votes = len(data_manager.vote_records)
        # 统计至少有一个候选人评分达到阈值的投票人数
        valid_voters = sum(1 for item in result if item["是否有效票"] == "是")
        info_frame = ttk.Frame(top, padding=5)
        info_frame.pack(fill='x', pady=5)
        info_text = f"总投票人数：{total_votes}   |   评分≥{threshold}%的候选人人数：{valid_voters}   |   基准线=每位投票股东总持股×阈值（{threshold}%）={baseline:.2f}   |   当选人数：{elected_limit}"
        ttk.Label(info_frame, text=info_text, font=("微软雅黑", 10)).pack()

        # 创建 Treeview 和滚动条
        tree_frame = ttk.Frame(top)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ("排名", "候选人", "总得分", "基准线", "是否有效票", "状态")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        col_widths = {"排名": 80, "候选人": 200, "总得分": 120, "基准线": 120, "是否有效票": 100, "状态": 100}
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths[col], minwidth=60, anchor="center")

        # 插入数据并设置交替行颜色
        for i, item in enumerate(result):
            tags = "elected" if item["状态"] == "当选" else "normal"
            if i % 2 == 0:
                tags = (tags, "evenrow")
            else:
                tags = (tags, "oddrow")
            tree.insert("", "end", values=(
                item["排名"], item["候选人"], f"{item['总得分']:.2f}",
                f"{item['基准线']:.2f}", item["是否有效票"], item["状态"]
            ), tags=tags)

        tree.tag_configure("elected", background="#A8D08D")
        tree.tag_configure("evenrow", background="#F8F8F8")
        tree.tag_configure("oddrow", background="#FFFFFF")
        # 当选行覆盖背景色（优先）
        tree.tag_configure("elected", background="#C6EFCE")

        # 布局
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        ttk.Button(top, text="确定关闭", command=top.destroy, width=15).pack(pady=10)

    def reset_vote_data(self):
        if data_manager.config["server_running"]:
            messagebox.showerror("操作失败", MSG["reset_need_stop_server"])
            return
        if messagebox.askyesno("确认", MSG["reset_confirm"]):
            data_manager.reset_vote_data()
            self.unfinished_list.delete(0,END)
            messagebox.showinfo("成功", MSG["reset_success"])
            self.refresh_ui()

    def refresh_ui(self):
        self.candidate_list.delete(0,END)
        for i,n in enumerate(data_manager.config["default_candidates"]):
            self.candidate_list.insert(END,f"{i+1}. {n}")
        self.voter_list.delete(0,END)
        for i,n in enumerate(data_manager.config["voters"]):
            self.voter_list.insert(END,f"{i+1}. {n}")

    def update_title(self):
        data_manager.config["title"] = self.title_entry.get().strip()
        self.root.title(f'{data_manager.config["title"]} - 投票管理系统')
        messagebox.showinfo('成功', '标题已更新！')
        data_manager.save_to_file()

    def update_mode(self):
        data_manager.config["vote_mode"] = self.mode_var.get()
        data_manager.save_to_file()

    def update_config(self):
        data_manager.config["show_extra_candidates"] = self.extra_var.get()
        data_manager.save_to_file()

    def add_candidates(self):
        top=Toplevel(self.root)
        top.title("批量添加候选人")
        text=scrolledtext.ScrolledText(top,width=40,height=10)
        text.pack(pady=5,padx=10)
        def save():
            ls=[x.strip() for x in text.get(1.0,END).split() if x.strip()]
            for l in ls:
                if l not in data_manager.config["default_candidates"]:
                    data_manager.config["default_candidates"].append(l)
            self.refresh_ui()
            data_manager.save_to_file()
            top.destroy()
        ttk.Button(top,text="保存",command=save).pack(pady=5)

    def delete_candidate(self):
        s=self.candidate_list.curselection()
        if s:
            data_manager.config["default_candidates"].pop(s[0])
            self.refresh_ui()
            data_manager.save_to_file()

    def clear_candidates(self):
        if messagebox.askyesno("确认","清空所有候选人？"):
            data_manager.config["default_candidates"].clear()
            self.refresh_ui()
            data_manager.save_to_file()

    def add_voters(self):
        top = Toplevel(self.root)
        top.title("批量添加选举人")
        ttk.Label(top, text="每行一个，格式：姓名+身份证号（例如：张三11010119900307663X）").pack(pady=5)
        text = scrolledtext.ScrolledText(top, width=50, height=10)
        text.pack(pady=5, padx=10)

        def save():
            lines = text.get(1.0, END).splitlines()
            new_voters = []
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # 使用更宽松的正则：匹配任意非数字字符开头（姓名），后跟15或18位数字或末尾X
                # 正则解释：([^\d]+) 匹配1个或多个非数字字符（姓名），然后 (\d{15}[\dXx]|\d{17}[\dXx]) 匹配15位数字+数字/X 或 17位数字+数字/X
                match = re.match(r'^([^\d]+)(\d{15}[\dXx]|\d{17}[\dXx])$', line)
                if not match:
                    messagebox.showerror("错误", f"格式错误：{line}\n需要姓名+18位身份证号（例如：张三11010119900307663X）")
                    return
                name_part = match.group(1).strip()
                id_card = match.group(2).upper()
                if not name_part or len(id_card) not in (15, 18):
                    messagebox.showerror("错误", f"无效姓名或身份证号：{line}")
                    return
                full_name = f"{name_part}{id_card}"
                if full_name not in data_manager.config["voters"]:
                    new_voters.append(full_name)
            if new_voters:
                data_manager.config["voters"].extend(new_voters)
                self.refresh_ui()
                data_manager.save_to_file()
                messagebox.showinfo("成功", f"已添加 {len(new_voters)} 人")
            top.destroy()

        ttk.Button(top, text="保存", command=save).pack(pady=5)
    
    def delete_voter(self):
        s=self.voter_list.curselection()
        if s:
            data_manager.config["voters"].pop(s[0])
            self.refresh_ui()
            data_manager.save_to_file()

    def clear_voters(self):
        if messagebox.askyesno("确认","清空所有选举人？"):
            data_manager.config["voters"].clear()
            self.unfinished_list.delete(0, END)
            self.refresh_ui()
            data_manager.save_to_file()

    def check_vote_status(self):
        voted={r["voter"] for r in data_manager.vote_records}
        un=set(data_manager.config["voters"])-voted
        self.unfinished_list.delete(0,END)
        for i,n in enumerate(sorted(un)):
            self.unfinished_list.insert(END,f"{i+1}. {n}")
        if not un:
            messagebox.showinfo('完成', '✅ 所有人已完成投票！')
        else:
            messagebox.showwarning('提示', f'未投票人数：{len(un)}')
    def copy_voters_list(self):
        """复制当前选举人列表（每行一个姓名）到剪贴板"""
        voters = data_manager.config.get("voters", [])
        if not voters:
            messagebox.showinfo("提示", "选举人列表为空")
            return
        text = "\n".join(voters)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("成功", f"已复制 {len(voters)} 条选举人名单到剪贴板")

    def import_signed_users(self):
        """从签到记录中导入未在选举人列表中的用户"""
        signin_file = os.path.join(BASE_DIR, "signin_records.json")
        if not os.path.exists(signin_file):
            messagebox.showinfo("提示", "暂无签到记录")
            return
        try:
            with open(signin_file, "r", encoding="utf-8") as f:
                signin_data = json.load(f)
        except Exception as e:
            messagebox.showerror("错误", f"读取签到记录失败：{e}")
            return

        existing_voters = set(data_manager.config["voters"])
        new_voters = []
        for key, info in signin_data.items():
            full_name = f"{info['name']}{info['id_card']}"
            if full_name not in existing_voters:
                new_voters.append(full_name)
        if not new_voters:
            messagebox.showinfo("提示", "所有签到用户均已存在于选举人列表中")
            return

        data_manager.config["voters"].extend(new_voters)
        data_manager.save_to_file()
        self.refresh_ui()
        messagebox.showinfo("成功", f"已导入 {len(new_voters)} 位新用户\n\n" + "\n".join(new_voters[:10]) + ("\n..." if len(new_voters) > 10 else ""))
# ==============================================
# 程序入口
# ==============================================
if __name__ == "__main__":
    main_root = Tk()
    app_gui = VoteSystemGUI(main_root)
    main_root.mainloop()