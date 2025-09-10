#!/usr/bin/env python3

# -*- coding: utf-8 -*-
"""
自動PDF監視・分類ツール（配布版）
複数フォルダを同時監視、フォルダごとの分類設定対応
"""

import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext, ttk
from tkinter import font as tkfont
import os
import time
import threading
from datetime import datetime
import json
import sys
import winreg
import socket
import platform
import ctypes
import ctypes.wintypes as wintypes
import random
import statistics as stats

# ファイル監視用
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# システムトレイ用
import pystray
from PIL import Image, ImageDraw

# PDF処理用
import fitz  # PyMuPDF
import io
import base64
import anthropic
import re

class PDFWatcherHandler(FileSystemEventHandler):
    """PDFファイル監視イベントハンドラー"""
    
    def __init__(self, classifier):
        self.classifier = classifier
        
    def on_created(self, event):
        """新しいファイルが作成されたときの処理"""
        if not event.is_directory and event.src_path.lower().endswith('.pdf'):
            self.classifier.log_message(f"新しいPDFを検出: {os.path.basename(event.src_path)}")
            threading.Timer(2.0, self.classifier.process_new_file, args=[event.src_path]).start()

class FolderSettingsDialog:
    """フォルダ別設定ダイアログ"""
    
    def __init__(self, parent, folder_info):
        self.parent = parent
        self.folder_info = folder_info.copy()
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("フォルダ別設定")
        self.dialog.geometry("900x720")
        try:
            self.dialog.minsize(800, 600)
            self.dialog.resizable(True, True)
        except Exception:
            pass
        self.dialog.configure(bg="white")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._is_fullscreen = False
        self._prev_geometry = None
        self.setup_dialog()
        
        # ダイアログを中央に配置
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))
    
    def setup_dialog(self):
        """ダイアログのGUI構築"""
        # スクロール可能なコンテンツ領域（ボタンは下部固定）
        self._canvas = tk.Canvas(self.dialog, bg="white", highlightthickness=0)
        self._vscroll = ttk.Scrollbar(self.dialog, orient='vertical', command=self._canvas.yview)
        self._content = tk.Frame(self._canvas, bg="white")
        self._content.bind("<Configure>", lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        # Canvas内のフレームをキャンバス幅に追従させる（見切れ防止）
        self._content_win = self._canvas.create_window((0,0), window=self._content, anchor='nw')
        self._canvas.bind('<Configure>', lambda e: self._canvas.itemconfigure(self._content_win, width=e.width))
        self._canvas.configure(yscrollcommand=self._vscroll.set)
        self._canvas.pack(side='left', fill='both', expand=True)
        self._vscroll.pack(side='right', fill='y')

        # スクロール用マウスホイール
        def _on_mousewheel(event):
            try:
                delta = -1*(event.delta//120)
            except Exception:
                delta = 1 if event.num == 5 else -1
            self._canvas.yview_scroll(delta, "units")
        self._canvas.bind_all('<MouseWheel>', _on_mousewheel)
        self._canvas.bind_all('<Button-4>', _on_mousewheel)
        self._canvas.bind_all('<Button-5>', _on_mousewheel)

        # タイトル
        title_label = tk.Label(
            self._content,
            text="フォルダ設定（シンプル）",
            font=("Arial", 16, "bold"),
            bg="white"
        )
        title_label.pack(pady=(16, 6))

        # サブ説明
        tk.Label(
            self._content,
            text="このフォルダでの自動命名ルールをシンプルに設定します。",
            bg="white", fg="#666"
        ).pack(padx=20, anchor="w")

        # フォルダパス表示
        path_frame = tk.Frame(self._content, bg="white")
        path_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(path_frame, text="フォルダ:", bg="white", font=("Arial", 10, "bold")).pack(anchor="w")
        
        path_label = tk.Label(
            path_frame,
            text=self.folder_info.get('path', ''),
            bg="white",
            relief="sunken",
            anchor="w",
            font=("Arial", 9),
            wraplength=450
        )
        path_label.pack(fill="x", pady=5)
        
        # 設定オプション
        settings_frame = tk.LabelFrame(
            self._content,
            text="基本オプション",
            bg="white",
            font=("Arial", 12, "bold"),
            padx=15,
            pady=15
        )
        settings_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # 監視有効/無効
        self.enabled_var = tk.BooleanVar(value=self.folder_info.get('enabled', True))
        enabled_check = tk.Checkbutton(
            settings_frame,
            text="🔍 このフォルダの監視を有効にする",
            variable=self.enabled_var,
            font=("Arial", 11),
            bg="white"
        )
        enabled_check.pack(anchor="w", pady=5)
        
        # 日付/名前の付加（基本設定で選択）
        self.include_date_var = tk.BooleanVar(value=self.folder_info.get('include_date', False))
        self.include_names_var = tk.BooleanVar(value=self.folder_info.get('include_names', False))

        date_check = tk.Checkbutton(
            settings_frame,
            text="📅 ファイル名に日付を付ける",
            variable=self.include_date_var,
            font=("Arial", 11),
            bg="white"
        )
        date_check.pack(anchor="w", pady=4)

        names_check = tk.Checkbutton(
            settings_frame,
            text="👤 ファイル名に名前を付ける（会社名/氏名）",
            variable=self.include_names_var,
            font=("Arial", 11),
            bg="white"
        )
        names_check.pack(anchor="w", pady=4)

        # かんたんAI設定（プリセット + キーワード）
        easy_frame = tk.LabelFrame(
            self._content,
            text="AI命名のかんたん設定",
            bg="white",
            font=("Arial", 12, "bold"),
            padx=15,
            pady=12
        )
        easy_frame.pack(pady=(0, 20), padx=20, fill="both", expand=True)

        # 自然言語でAIに指示 + 有効/無効トグル
        top_row = tk.Frame(easy_frame, bg='white')
        top_row.pack(fill='x')
        tk.Label(top_row, text="AI命名の指示（自然文でOK）", bg="white", font=("Arial", 11, "bold")).pack(side='left')
        self.use_custom_instruction = tk.BooleanVar(value=bool(self.folder_info.get('use_custom_instruction', True)))
        ttk.Checkbutton(top_row, text='この指示を使う', variable=self.use_custom_instruction, command=lambda: self._toggle_instruction()).pack(side='right')

        self.instruction_text = tk.Text(easy_frame, height=5, bg='#F8FAFC')
        self.instruction_text.pack(fill='x')
        # 既存のカスタム指示を復元、なければ例文を提示
        existing = (self.folder_info.get('custom_classify_prompt') or '').strip()
        if existing:
            self.instruction_text.insert('1.0', existing)
        else:
            example = (
                "例:\n"
                "1) 1ページ目の見出しをそのまま短く使う。種類名を優先（見積書/契約書など）。\n"
                "2) 余計な説明や会社名は含めない。名詞句のみで8〜16文字程度。"
            )
            self.instruction_text.insert('1.0', example)
        tk.Label(easy_frame, text="ヒント: 雰囲気ではなく“こう名付けたい”を短く書くと安定します。", bg="white", fg="#666").pack(anchor='w', pady=(6, 0))
        # 互換: 旧プリセット/キーワード参照が残っている箇所へのダミー変数
        try:
            self.prompt_preset_var = tk.StringVar(value=self.folder_info.get('prompt_preset', 'auto'))
            self.prompt_keywords_var = tk.StringVar(value='')
        except Exception:
            pass
        # 初期トグル反映
        self._toggle_instruction()

        
        
        # ボタンフレーム
        button_frame = tk.Frame(self.dialog, bg="white")
        button_frame.pack(side='bottom', fill='x', pady=10)

        # OK・キャンセルボタン

        ok_btn = tk.Button(
            button_frame,
            text="OK",
            command=self.on_ok,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11),
            padx=20
        )
        ok_btn.pack(side="left", padx=(0, 10))
        
        cancel_btn = tk.Button(
            button_frame,
            text="キャンセル",
            command=self.on_cancel,
            bg="#f44336",
            fg="white",
            font=("Arial", 11),
            padx=20
        )
        cancel_btn.pack(side="left")
    
    # 旧・詳細設定/出力設定のトグル機能は廃止

    def toggle_fullscreen(self):
        """設定ダイアログの全画面/元サイズを切り替え"""
        try:
            if not self._is_fullscreen:
                # 保存
                self._prev_geometry = self.dialog.geometry()
                try:
                    # Windowsで有効
                    self.dialog.state('zoomed')
                except Exception:
                    # 他環境向けフォールバック
                    sw = self.parent.winfo_screenwidth()
                    sh = self.parent.winfo_screenheight()
                    self.dialog.geometry(f"{sw}x{sh}+0+0")
                self._is_fullscreen = True
            else:
                try:
                    self.dialog.state('normal')
                except Exception:
                    pass
                if self._prev_geometry:
                    self.dialog.geometry(self._prev_geometry)
                self._is_fullscreen = False
        except Exception:
            pass
    
    # 旧・命名ルールCRUDは廃止
    
    def _build_simple_prompt(self, preset_key: str, keywords: str) -> str:
        preset = (preset_key or 'auto').lower()
        # 強調キーワードは廃止
        header = (
            "日本語のPDF書類について、1ページ目を主たるページとして優先し、"
            "文書の種類名を短く1つだけ返してください。説明・語尾は不要です。"
        )
        if preset == 'legal':
            body = "候補例: 契約書/覚書/規程/議事録/定款/委任状/就任承諾書/印鑑証明/登記事項証明書"
        elif preset == 'business':
            body = "候補例: 見積書/請求書/領収書/注文書/納品書/仕様書/送付状/請求明細"
        elif preset == 'realestate':
            body = "候補例: 登記事項証明書/重要事項説明/売買契約書/賃貸借契約書/不動産契約書"
        else:
            body = "一般的なビジネス文書を前提として、分かりやすい種別名だけを返してください。"
        return "\n".join([header, body]).strip()

    def _update_prompt_preview(self):
        try:
            txt = '🧠 ' + self._build_simple_prompt(self.prompt_preset_var.get(), '')
            self._set_preview_text(txt)
        except Exception:
            pass

    # ---------- かんたんUIユーティリティ ----------
    def _init_kw_placeholder(self):
        try:
            pass
        except Exception:
            pass

    def _render_kw_tags(self):
        try:
            pass
        except Exception:
            pass

    def _toggle_instruction(self):
        try:
            state = 'normal' if self.use_custom_instruction.get() else 'disabled'
            self.instruction_text.configure(state=state)
        except Exception:
            pass

    def _set_preview_text(self, text: str):
        try:
            self.prompt_preview.configure(state='normal')
            self.prompt_preview.delete('1.0', 'end')
            self.prompt_preview.insert('1.0', text)
            self.prompt_preview.configure(state='disabled')
        except Exception:
            pass
    
    def on_ok(self):
        """OK ボタンクリック時の処理（簡素版）"""
        try:
            preset_key = self.prompt_preset_var.get()
            keywords = self.prompt_keywords_var.get()
            # ユーザーの自然文指示をそのまま使う（無効なら空）
            custom_prompt_val = (self.instruction_text.get('1.0', 'end') or '').strip() if self.use_custom_instruction.get() else ''

            self.result = {
                'path': self.folder_info['path'],
                'enabled': self.enabled_var.get(),
                'include_date': self.include_date_var.get(),
                'include_names': self.include_names_var.get(),
                'prompt_preset': preset_key,
                'use_custom_instruction': self.use_custom_instruction.get(),
                'custom_classify_prompt': custom_prompt_val if custom_prompt_val else None,
            }
            print(f"設定結果: {self.result}")
            self.dialog.destroy()
        except Exception as e:
            print(f"OK処理エラー: {e}")
            self.dialog.destroy()
    
    def on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        self.result = None
        self.dialog.destroy()

    # フォルダ設定ダイアログのメソッド末尾（ここにはウェルカムは置かない）

class AutoPDFWatcherAdvanced:
    def __init__(self):
        # GUI初期化（シングルトン確保後に実行されることが前提）
        self.window = tk.Tk()
        self.window.title("紙の名は。")
        self.window.geometry("1200x900")
        try:
            self.window.minsize(1000, 800)
        except Exception:
            pass
        self.window.configure(bg="white")
        
        # 包括的アイコンデバッグ実行（一時的に無効化）
        # self.debug_icon_comprehensive()
        
        # ウィンドウアイコン（Windows / PyInstaller対応）
        self.setup_window_icon_robust()
        
        # 設定ファイル（AppDataに保存）
        try:
            self.config_file = os.path.join(self._get_appdata_dir(), 'config.json')
        except Exception:
            self.config_file = "auto_watcher_advanced_config.json"  # フォールバック
        # 旧位置からの移行
        try:
            self._migrate_legacy_config()
        except Exception:
            pass
        self.config = self.load_config()
        
        # 単一インスタンス制御は起動前に実施（main側）
        
        # Claude API
        self.claude_client = None
        self.init_claude_api()
        # 使用モデル名（ユーザー設定可能）
        # 既定は Claude 4 Sonnet（API ID: claude-sonnet-4-20250514）
        self.model_name = self.config.get('model', 'claude-sonnet-4-20250514')
        
        # 監視関連（フォルダ別設定対応）
        self.observers = []
        self.is_watching = False
        self.watch_folders = self.config.get('watch_folders', [])
        
        # グローバル設定オプション
        self.auto_startup = tk.BooleanVar(value=self.config.get('auto_startup', False))
        # 最小化時はトレイに格納（常駐）
        self.minimize_to_tray = tk.BooleanVar(value=True)
        
        # 除外ワード設定（廃止 - 差出人/宛名の文脈判断に移行）
        
        # システムトレイ
        self.tray_icon = None
        self.is_minimized_to_tray = False
        self._tray_detached = False
        self._tray_thread_started = False
        
        # ログ保持（一定時間で自動クリア）※GUI構築より前に変数を用意
        self.log_history = []  # (timestamp, text)
        self.log_retention_minutes = tk.IntVar(value=int(self.config.get('log_retention_minutes', 60)))
        # ファイル名の最大長（ユーザー変更可）
        self.max_filename_length = tk.IntVar(value=int(self.config.get('max_filename_length', 32)))

        # GUI構築
        self.setup_gui()
        
        # ウィンドウイベント
        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)
        self.window.bind('<Unmap>', self.on_minimize)
        
        # 起動時はトレイを起動しない（最小化時に起動）
        
        # 自動監視開始
        if self.config.get('auto_start_monitoring', False) and self.watch_folders and self.claude_client:
            self.window.after(1000, self.start_watching)

        # ログ自動剪定のスケジュール
        self.window.after(60_000, self._prune_log_periodic)

        # UIフォント/テーマの軽いモダン化
        try:
            default_font = tkfont.nametofont("TkDefaultFont")
            fam = None
            if platform.system() == 'Windows':
                # 日本語に適した見やすいUIフォント
                for candidate in ("Yu Gothic UI", "Meiryo UI", "Segoe UI"):
                    try:
                        tkfont.Font(family=candidate)
                        fam = candidate
                        break
                    except Exception:
                        continue
            else:
                fam = "Segoe UI"
            if fam:
                default_font.configure(family=fam, size=11)
            # UIフォント名を保存（ログや他でも使用）
            try:
                self.ui_font_family = fam or default_font.actual('family')
            except Exception:
                self.ui_font_family = 'Segoe UI'
            self.window.option_add("*Font", default_font)
            style = ttk.Style()
            if 'clam' in style.theme_names():
                style.theme_use('clam')
            # スタイルセット
            self.setup_styles()
        except Exception:
            pass

        # 初回オンボーディングは表示しない（ユーザー主導で設定）

    def show_welcome_dialog(self):
        try:
            dlg = tk.Toplevel(self.window)
            dlg.title("ようこそ")
            dlg.configure(bg='white')
            dlg.geometry("520x360")
            tk.Label(dlg, text='PDF自動整理アシスタントへようこそ', bg='white', fg='#111', font=("Arial", 14, 'bold')).pack(pady=(16,8))
            msg = (
                "このアプリは、フォルダを監視してPDFの名前を自動で分かりやすく整えます。\n\n"
                "簡単3ステップ:\n"
                "1) 監視フォルダを追加\n2) 用途プリセットとキーワードを設定\n3) サンプルPDFで動作を確認"
            )
            tk.Label(dlg, text=msg, bg='white', fg='#374151', justify='left').pack(padx=18, anchor='w')
            btns = tk.Frame(dlg, bg='white')
            btns.pack(side='bottom', pady=12)
            def start():
                try:
                    self.config['welcome_shown'] = True
                    self.save_config()
                except Exception:
                    pass
                dlg.destroy()
                if not self.claude_client:
                    self.show_api_setup_dialog()
            def skip():
                try:
                    self.config['welcome_shown'] = True
                    self.save_config()
                except Exception:
                    pass
                dlg.destroy()
            tk.Button(btns, text='🚀 はじめる', command=start, bg='#3B82F6', fg='white', padx=18, pady=6).pack(side='left', padx=6)
            tk.Button(btns, text='⏭ スキップ', command=skip, bg='#E5E7EB', padx=16, pady=6).pack(side='left')
            dlg.transient(self.window)
            dlg.grab_set()
        except Exception:
            pass

    def setup_styles(self):
        try:
            style = ttk.Style()
            # ボタン
            style.configure('Primary.TButton', padding=(12, 8))
            style.map('Primary.TButton',
                      background=[('!disabled', '#3B82F6'), ('active', '#2563EB')],
                      foreground=[('!disabled', 'white')])
            style.configure('Secondary.TButton', padding=(12, 8))
            style.map('Secondary.TButton',
                      background=[('!disabled', '#E5E7EB'), ('active', '#D1D5DB')],
                      foreground=[('!disabled', '#1F2937')])
            # 危険色（停止ボタン用）
            style.configure('Danger.TButton', padding=(12, 8))
            style.map('Danger.TButton',
                      background=[('!disabled', '#EF4444'), ('active', '#DC2626')],
                      foreground=[('!disabled', 'white')])
            # フレーム
            style.configure('Top.TFrame', background='white')
            style.configure('Card.TFrame', background='#F8F9FA')
            # Treeview
            style.configure('Modern.Treeview', rowheight=28)
            # タグ風ラベル
            try:
                style.configure('Tag.TLabel', background='#E5E7EB', foreground='#374151', padding=(6, 2))
            except Exception:
                pass
        except Exception:
            pass

    
    
    def load_config(self):
        """設定ファイルを読み込み"""
        try:
            cfg_dir = os.path.dirname(self.config_file)
            if cfg_dir and not os.path.exists(cfg_dir):
                os.makedirs(cfg_dir, exist_ok=True)
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"設定読み込みエラー: {e}")
            return {}
    
    def save_config(self):
        """設定ファイルを保存"""
        try:
            cfg_dir = os.path.dirname(self.config_file)
            if cfg_dir and not os.path.exists(cfg_dir):
                os.makedirs(cfg_dir, exist_ok=True)
            self.config['watch_folders'] = self.watch_folders
            self.config['auto_startup'] = self.auto_startup.get()
            self.config['minimize_to_tray'] = self.minimize_to_tray.get()
            # 初回案内の表示状態
            self.config['welcome_shown'] = bool(self.config.get('welcome_shown', False))
            # 最大ファイル名長
            try:
                self.config['max_filename_length'] = int(self.max_filename_length.get())
            except Exception:
                pass
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
                
            self.update_startup_setting()
            
        except Exception as e:
            print(f"設定保存エラー: {e}")

    def _migrate_legacy_config(self):
        """旧カレントディレクトリの設定をAppDataへ移行"""
        try:
            legacy = 'auto_watcher_advanced_config.json'
            if os.path.isabs(legacy):
                return
            if os.path.exists(legacy) and not os.path.exists(self.config_file):
                # 読み取り→AppDataへコピー
                with open(legacy, 'r', encoding='utf-8') as f:
                    data = f.read()
                cfg_dir = os.path.dirname(self.config_file)
                os.makedirs(cfg_dir, exist_ok=True)
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    f.write(data)
                # 旧ファイルは残してもよいが、気になる場合はコメントアウト解除で削除
                # os.remove(legacy)
        except Exception:
            pass

    def export_config(self):
        try:
            path = filedialog.asksaveasfilename(title="設定のエクスポート", defaultextension=".json", filetypes=[["JSON","*.json"]])
            if not path:
                return
            self.save_config()
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("完了", "設定をエクスポートしました")
        except Exception as e:
            messagebox.showerror("エラー", f"エクスポートに失敗しました: {e}")

    def import_config(self):
        try:
            path = filedialog.askopenfilename(title="設定のインポート", filetypes=[["JSON","*.json"]])
            if not path:
                return
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # 安全なキーのみ適用
            self.watch_folders = data.get('watch_folders', self.watch_folders)
            self.auto_startup.set(bool(data.get('auto_startup', self.auto_startup.get())))
            self.minimize_to_tray.set(bool(data.get('minimize_to_tray', self.minimize_to_tray.get())))
            # 反映
            self.update_folder_tree()
            self.save_config()
            messagebox.showinfo("完了", "設定をインポートしました")
        except Exception as e:
            messagebox.showerror("エラー", f"インポートに失敗しました: {e}")
    
    def init_claude_api(self):
        """Claude API初期化"""
        try:
            api_key = os.environ.get('ANTHROPIC_API_KEY') or self._read_api_key_from_appdata()
            if api_key:
                self.claude_client = anthropic.Anthropic(api_key=api_key)
        except Exception as e:
            print(f"Claude API初期化エラー: {e}")

    def _get_appdata_dir(self) -> str:
        base = os.environ.get('APPDATA') or os.environ.get('LOCALAPPDATA') or os.path.expanduser('~/.config')
        return os.path.join(base, 'AutoPDFWatcherAdvanced')

    def _api_secret_path(self) -> str:
        return os.path.join(self._get_appdata_dir(), 'credentials.json')

    def _read_api_key_from_appdata(self) -> str | None:
        try:
            path = self._api_secret_path()
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    key = (data.get('anthropic_api_key') or '').strip()
                    if key:
                        return key
            return None
        except Exception:
            return None

    def _write_api_key_to_appdata(self, key: str):
        try:
            d = self._get_appdata_dir()
            os.makedirs(d, exist_ok=True)
            path = self._api_secret_path()
            data = {'anthropic_api_key': key}
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            # 権限を制限（POSIX）
            try:
                os.chmod(path, 0o600)
            except Exception:
                pass
        except Exception as e:
            raise e

    def _resolve_ico_path(self) -> str | None:
        r"""ico の探索優先度（WSLパス依存を除去）:
        1) PyInstaller バンドル (_MEIPASS) - 最優先
        2) 環境変数 KAMINONAHA_ICON
        3) EXE/スクリプトと同じフォルダ
        4) AppData\\Kaminonaha\\kaminonaha_latest.ico
        5) 開発用固定パス
        """
        try:
            cand = []
            
            # 1) PyInstaller bundled - 最優先（配布時）
            base_mei = getattr(sys, '_MEIPASS', None)
            if base_mei:
                bundled_path = os.path.join(base_mei, 'kaminonaha_latest.ico')
                cand.append(bundled_path)
                print(f"PyInstallerバンドルパス候補: {bundled_path}")
            
            # 2) env override
            env = os.environ.get('KAMINONAHA_ICON')
            if env:
                cand.append(env)
                print(f"環境変数パス候補: {env}")
            
            # 3) alongside - WSL環境を考慮
            if getattr(sys, 'frozen', False):
                # EXE実行時：実行ファイルと同じフォルダ
                exe_dir = os.path.dirname(sys.executable)
                alongside_path = os.path.join(exe_dir, 'kaminonaha_latest.ico')
                cand.append(alongside_path)
                print(f"EXE同フォルダパス候補: {alongside_path}")
            else:
                # スクリプト実行時：WSL環境を検出
                script_dir = os.path.dirname(__file__)
                if '\\wsl.localhost\\' in script_dir:
                    print(f"WSL環境を検出: {script_dir}")
                    # WSLパスの場合、開発用パスを優先
                    dev_paths = [
                        r'C:\dev\kaminonaha\kaminonaha_latest.ico',  # 開発環境
                        os.path.join(script_dir, 'kaminonaha_latest.ico')  # WSLパス（フォールバック）
                    ]
                    cand.extend(dev_paths)
                    print(f"WSL用候補パス: {dev_paths}")
                else:
                    # 通常のWindows環境
                    alongside_path = os.path.join(script_dir, 'kaminonaha_latest.ico')
                    cand.append(alongside_path)
                    print(f"スクリプト同フォルダパス候補: {alongside_path}")
            
            # 4) AppData
            appdata = os.environ.get('APPDATA') or os.path.expanduser('~/.config')
            appdata_path = os.path.join(appdata, 'Kaminonaha', 'kaminonaha_latest.ico')
            cand.append(appdata_path)
            print(f"AppDataパス候補: {appdata_path}")
            
            # パス探索実行
            for p in cand:
                if p and os.path.exists(p):
                    print(f"✅ 発見されたICOパス: {p}")
                    return p
                else:
                    print(f"❌ 存在しないパス: {p}")
            
            print("❌ ICOファイルが見つかりませんでした")
            return None
        except Exception as e:
            print(f"ICOパス解決エラー: {e}")
            return None

    def debug_icon_comprehensive(self):
        """アイコン問題の包括的デバッグ"""
        print("=" * 60)
        print("🔍 包括的アイコンデバッグ開始")
        print("=" * 60)
        
        # 1. 実行環境の詳細確認
        print("📋 実行環境:")
        print(f"  OS: {platform.system()} {platform.release()}")
        print(f"  Python: {sys.version.split()[0]}")
        print(f"  実行モード: {'PyInstaller' if getattr(sys, 'frozen', False) else 'スクリプト'}")
        print(f"  カレントディレクトリ: {os.getcwd()}")
        print(f"  __file__: {__file__ if '__file__' in globals() else 'N/A'}")
        print(f"  sys.executable: {sys.executable}")
        if hasattr(sys, '_MEIPASS'):
            print(f"  PyInstaller _MEIPASS: {sys._MEIPASS}")
        print()
        
        # 2. 現在のウィンドウアイコン確認
        print("🖼️ 現在のウィンドウアイコン:")
        try:
            current_icon = self.window.tk.call('wm', 'iconbitmap', self.window._w)
            print(f"  現在設定されているアイコン: {current_icon}")
            if current_icon and os.path.exists(current_icon):
                size = os.path.getsize(current_icon)
                print(f"  ファイルサイズ: {size} bytes")
            elif current_icon:
                print(f"  ⚠️ 設定されているが、ファイルが存在しない")
            else:
                print(f"  ❌ アイコンが設定されていない")
        except Exception as e:
            print(f"  ❌ アイコン取得エラー: {e}")
        print()
        
        # 3. _resolve_ico_path の詳細実行
        print("🔍 アイコンパス解決プロセス:")
        ico_path = self._resolve_ico_path_debug()
        print()
        
        # 4. ファイルシステム上のアイコンファイル検索
        print("📁 ファイルシステム上のアイコンファイル検索:")
        self._search_icon_files_systemwide()
        print()
        
        # 5. Windowsレジストリの確認（PyInstaller関連）
        if platform.system() == 'Windows':
            print("📝 Windowsアイコンキャッシュ情報:")
            self._check_windows_icon_cache()
            print()
        
        # 6. 既知の問題パターンチェック
        print("⚠️ 既知の問題パターンチェック:")
        self._check_known_icon_issues()
        print()
        
        print("=" * 60)
        print("🔍 デバッグ完了")
        print("=" * 60)

    def _resolve_ico_path_debug(self) -> str | None:
        """デバッグ版: アイコンパス解決の詳細出力"""
        try:
            candidates = []
            
            # 1) 環境変数
            env = os.environ.get('KAMINONAHA_ICON')
            print(f"  1. 環境変数 KAMINONAHA_ICON: {env if env else 'なし'}")
            if env:
                candidates.append(('環境変数', env))
            
            # 2) PyInstaller バンドル
            if hasattr(sys, '_MEIPASS'):
                bundled = os.path.join(sys._MEIPASS, 'kaminonaha_latest.ico')
                print(f"  2. PyInstaller バンドル: {bundled}")
                candidates.append(('バンドル', bundled))
            else:
                print(f"  2. PyInstaller バンドル: 該当なし（スクリプト実行）")
            
            # 3) 実行ファイル隣接
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
                print(f"  3. 実行ファイルディレクトリ: {exe_dir}")
            else:
                exe_dir = os.path.dirname(os.path.abspath(__file__))
                print(f"  3. スクリプトディレクトリ: {exe_dir}")
            
            adjacent_files = ['kaminonaha_latest.ico', 'icon.ico', 'app.ico']
            for filename in adjacent_files:
                path = os.path.join(exe_dir, filename)
                candidates.append(('実行ファイル隣接', path))
            
            # 4) AppData
            try:
                appdata = os.environ.get('APPDATA') or os.environ.get('LOCALAPPDATA')
                if appdata:
                    appdata_path = os.path.join(appdata, 'Kaminonaha', 'kaminonaha_latest.ico')
                    print(f"  4. AppData: {appdata_path}")
                    candidates.append(('AppData', appdata_path))
                else:
                    print(f"  4. AppData: 環境変数が見つからない")
            except Exception as e:
                print(f"  4. AppData: エラー - {e}")
            
            # 5) 固定パス
            fixed_paths = [
                r'C:\dev\kaminonaha\kaminonaha_latest.ico',
                r'C:\Program Files\Kaminonaha\kaminonaha_latest.ico'
            ]
            print(f"  5. 固定パス候補:")
            for path in fixed_paths:
                print(f"     {path}")
                candidates.append(('固定パス', path))
            
            # 候補の詳細チェック
            print(f"  📋 候補ファイルの詳細チェック:")
            for i, (source, path) in enumerate(candidates, 1):
                if os.path.exists(path):
                    try:
                        size = os.path.getsize(path)
                        # ファイルの先頭数バイトを確認（ICO形式チェック）
                        with open(path, 'rb') as f:
                            header = f.read(4)
                            is_ico = header[:2] == b'\x00\x00'
                        
                        status = "✅ 有効なICO" if is_ico and size > 100 else "⚠️ 無効"
                        print(f"    {i}. {status} [{source}] {path} ({size} bytes)")
                        
                        if is_ico and size > 100:
                            print(f"       👆 このファイルを使用します")
                            return path
                            
                    except Exception as e:
                        print(f"    {i}. ❌ 読み込みエラー [{source}] {path} - {e}")
                else:
                    print(f"    {i}. ❌ 存在しない [{source}] {path}")
            
            print(f"  ⚠️ 有効なアイコンファイルが見つかりませんでした")
            return None
            
        except Exception as e:
            print(f"  ❌ アイコンパス解決エラー: {e}")
            return None

    def _search_icon_files_systemwide(self):
        """システム全体でアイコンファイルを検索"""
        try:
            import glob
            
            # 検索対象ディレクトリ
            search_dirs = [
                os.getcwd(),  # カレントディレクトリ
                os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__),
                os.path.expanduser('~'),  # ホームディレクトリ
                os.environ.get('APPDATA', ''),
                os.environ.get('LOCALAPPDATA', ''),
            ]
            
            # 検索パターン
            patterns = ['*kaminonaha*.ico', '*icon*.ico', '*.ico']
            
            print(f"  検索対象ディレクトリ:")
            for directory in search_dirs:
                if directory and os.path.exists(directory):
                    print(f"    📁 {directory}")
                    
                    for pattern in patterns:
                        try:
                            search_path = os.path.join(directory, '**', pattern)
                            files = glob.glob(search_path, recursive=True)[:10]  # 最大10個
                            
                            for file in files:
                                try:
                                    size = os.path.getsize(file)
                                    if size > 100:  # 意味のあるサイズのファイルのみ
                                        print(f"      🔍 発見: {file} ({size} bytes)")
                                except Exception:
                                    pass
                                    
                        except Exception as e:
                            if "Access is denied" not in str(e):
                                print(f"      ⚠️ 検索エラー: {e}")
                else:
                    print(f"    ❌ {directory} (存在しない)")
                    
        except Exception as e:
            print(f"  ❌ システム検索エラー: {e}")

    def _check_windows_icon_cache(self):
        """Windowsアイコンキャッシュの確認"""
        try:
            if platform.system() != 'Windows':
                return
            
            # アイコンキャッシュファイルの場所
            cache_locations = [
                os.path.join(os.environ.get('LOCALAPPDATA', ''), 'IconCache.db'),
                os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Explorer'),
            ]
            
            print(f"  Windowsアイコンキャッシュ:")
            for location in cache_locations:
                if os.path.exists(location):
                    if os.path.isfile(location):
                        size = os.path.getsize(location)
                        mtime = os.path.getmtime(location)
                        from datetime import datetime
                        mtime_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                        print(f"    📄 {location} ({size} bytes, 更新: {mtime_str})")
                    else:
                        # ディレクトリの場合、中身を確認
                        try:
                            files = os.listdir(location)
                            icon_files = [f for f in files if 'icon' in f.lower()][:5]
                            print(f"    📁 {location} (アイコン関連: {len(icon_files)}個)")
                            for file in icon_files:
                                print(f"      - {file}")
                        except Exception:
                            print(f"    📁 {location} (アクセス不可)")
                else:
                    print(f"    ❌ {location} (存在しない)")
                    
        except Exception as e:
            print(f"  ❌ キャッシュ確認エラー: {e}")

    def _check_known_icon_issues(self):
        """既知のアイコン問題をチェック"""
        issues = []
        
        # WSLパス問題
        if '\\wsl.localhost\\' in os.getcwd() or '\\wsl.localhost\\' in __file__:
            issues.append("🔶 WSL環境での実行が検出されました。他のPCでは異なるアイコンが表示される可能性があります。")
        
        # PyInstaller問題
        if getattr(sys, 'frozen', False) and not hasattr(sys, '_MEIPASS'):
            issues.append("🔶 PyInstaller実行ですが_MEIPASSが見つかりません。--add-dataオプションでアイコンがバンドルされていない可能性があります。")
            issues.append("🔶 EXEファイルのタスクバーアイコンには --icon オプションでアイコンを埋め込む必要があります。")
        
        # 権限問題
        try:
            test_path = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), 'test_write.tmp')
            with open(test_path, 'w') as f:
                f.write('test')
            os.remove(test_path)
        except Exception:
            issues.append("🔶 実行ディレクトリに書き込み権限がありません。アイコンファイルの配置ができない可能性があります。")
        
        # アイコンファイル形式問題
        ico_path = self._resolve_ico_path()
        if ico_path and os.path.exists(ico_path):
            try:
                with open(ico_path, 'rb') as f:
                    header = f.read(4)
                    if header[:2] != b'\x00\x00':
                        issues.append(f"🔶 アイコンファイル形式が正しくありません: {ico_path}")
            except Exception:
                issues.append(f"🔶 アイコンファイルの読み込みに失敗しました: {ico_path}")
        
        if issues:
            for issue in issues:
                print(f"  {issue}")
        else:
            print(f"  ✅ 既知の問題は検出されませんでした")

    def setup_window_icon_robust(self):
        """確実なウィンドウアイコン設定"""
        if platform.system() != 'Windows':
            return
        
        try:
            print("[DEBUG] ウィンドウアイコン設定開始...")
            
            ico_path = self._resolve_ico_path()
            print(f"[DEBUG] 解決されたアイコンパス: {ico_path}")
            
            if ico_path and os.path.exists(ico_path):
                # 設定前の状態確認
                try:
                    current = self.window.tk.call('wm', 'iconbitmap', self.window._w)
                    print(f"[DEBUG] 設定前のアイコン: {current}")
                except:
                    print(f"[DEBUG] 設定前のアイコン: なし")
                
                # 方法1: 標準のiconbitmap設定
                try:
                    print(f"[DEBUG] アイコン設定試行: {ico_path}")
                    self.window.iconbitmap(ico_path)
                    print(f"[DEBUG] iconbitmap呼び出し成功")
                    
                    # 設定後の確認
                    try:
                        after = self.window.tk.call('wm', 'iconbitmap', self.window._w)
                        print(f"[DEBUG] 設定後のアイコン: {after}")
                        if after:
                            print(f"[DEBUG] ✅ ウィンドウアイコン設定成功")
                        else:
                            print(f"[DEBUG] ❌ ウィンドウアイコン設定失敗（設定値が空）")
                    except Exception as e:
                        print(f"[DEBUG] 設定確認エラー: {e}")
                        
                except Exception as e1:
                    print(f"[DEBUG] ❌ iconbitmap設定失敗: {e1}")
                    import traceback
                    traceback.print_exc()
                
                # 方法2: iconphoto設定（フォールバック）
                try:
                    print(f"[DEBUG] iconphoto設定試行...")
                    from PIL import Image, ImageTk
                    
                    with Image.open(ico_path) as img:
                        # 複数サイズのPhotoImageを作成
                        sizes_to_try = [(16, 16), (32, 32), (48, 48)]
                        icon_images = []
                        
                        n_frames = getattr(img, "n_frames", 1)
                        print(f"[DEBUG] ICOフレーム数: {n_frames}")
                        
                        for i in range(n_frames):
                            img.seek(i)
                            w, h = img.size
                            print(f"[DEBUG]   フレーム {i}: {w}x{h}px")
                            
                            # 複数サイズのPhotoImageを作成
                            for size in sizes_to_try:
                                try:
                                    if i == 0:  # 最初のフレームのみ使用
                                        resized = img.resize(size, Image.Resampling.LANCZOS)
                                        photo = ImageTk.PhotoImage(resized)
                                        icon_images.append(photo)
                                        print(f"[DEBUG]     {size[0]}x{size[1]}px PhotoImage作成成功")
                                except Exception as e2:
                                    print(f"[DEBUG]     {size[0]}x{size[1]}px PhotoImage作成失敗: {e2}")
                        
                        # iconphotoで設定
                        if icon_images:
                            self.window.iconphoto(True, *icon_images)
                            print(f"[DEBUG] ✅ iconphoto設定成功 ({len(icon_images)}個のサイズ)")
                        else:
                            print(f"[DEBUG] ❌ iconphoto用画像作成失敗")
                            
                except Exception as e2:
                    print(f"[DEBUG] ❌ iconphoto設定失敗: {e2}")
                    import traceback
                    traceback.print_exc()
                
                # Windows 11対応: 強化版Windows API でタスクバーアイコンを直接設定
                self.force_set_taskbar_icon_debug(ico_path)
                
                print(f"[DEBUG] ウィンドウアイコン設定処理完了")
                    
            else:
                print(f"[DEBUG] ❌ 有効なアイコンファイルが見つからない")
                
        except Exception as e:
            print(f"[DEBUG] ❌ ウィンドウアイコン設定エラー: {e}")
            import traceback
            traceback.print_exc()

    def force_set_taskbar_icon_debug(self, ico_path):
        try:
            import ctypes
            from ctypes import wintypes
            
            hwnd = self.window.winfo_id()
            print(f"[DEBUG] ウィンドウハンドル: {hwnd}")
            
            # 複数サイズでアイコン設定
            for size in [16, 32]:
                print(f"[DEBUG] {size}x{size}アイコン読み込み試行: {ico_path}")
                
                hicon = ctypes.windll.user32.LoadImageW(
                    None, ico_path, 1, size, size, 0x00000010
                )
                
                if hicon:
                    print(f"[DEBUG] {size}x{size}アイコンハンドル取得成功: {hicon}")
                    # WM_SETICON
                    icon_type = 0 if size == 16 else 1
                    result = ctypes.windll.user32.SendMessageW(hwnd, 0x0080, icon_type, hicon)
                    print(f"[DEBUG] SendMessage結果 ({size}px): {result}")
                else:
                    print(f"[DEBUG] {size}x{size}アイコンハンドル取得失敗")
                    
            # タスクバー更新を強制
            print(f"[DEBUG] タスクバー更新を強制中...")
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
            
            # ウィンドウ更新
            print(f"[DEBUG] ウィンドウ更新中...")
            self.window.update()
            
            print(f"[DEBUG] 強化版Windows API設定完了")
            
        except Exception as e:
            print(f"[DEBUG] Windows API設定エラー: {e}")
            import traceback
            traceback.print_exc()

    def _ensure_appdata_icon(self):
        """AppDataにアイコンを配置（存在しない/古い場合のみ）。
        - onefile同梱やEXE隣のicoを見つけて、%APPDATA%\Kaminonaha\kaminonaha_latest.ico へコピー
        - 失敗は無視（フォールバック探索に任せる）
        """
        try:
            appdata = os.environ.get('APPDATA') or os.path.expanduser('~/.config')
            target_dir = os.path.join(appdata, 'Kaminonaha')
            target = os.path.join(target_dir, 'kaminonaha_latest.ico')
            # 既に適切なファイルがあるならスキップ
            src = None
            # 環境変数最優先
            env = os.environ.get('KAMINONAHA_ICON')
            if env and os.path.exists(env):
                src = env
            if not src:
                base_mei = getattr(sys, '_MEIPASS', None)
                if base_mei:
                    p = os.path.join(base_mei, 'kaminonaha_latest.ico')
                    if os.path.exists(p):
                        src = p
            if not src:
                exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
                p = os.path.join(exe_dir, 'kaminonaha_latest.ico')
                if os.path.exists(p):
                    src = p
            if not src:
                p = r'C:\\dev\\kaminonaha\\kaminonaha_latest.ico'
                if os.path.exists(p):
                    src = p
            if not src:
                return
            os.makedirs(target_dir, exist_ok=True)
            # コピー（内容差分で更新）
            try:
                if not os.path.exists(target) or os.path.getsize(target) != os.path.getsize(src):
                    import shutil
                    shutil.copyfile(src, target)
            except Exception:
                pass
        except Exception:
            pass

    def tray_target_px_for_dpi(self, dpi: int) -> int:
        """DPIに基づいてトレイアイコンの最適サイズを返す（より高解像度対応）"""
        if dpi < 120:       # ～100%
            return 32  # 従来24を32に変更
        elif dpi < 144:     # ～125%
            return 40
        elif dpi < 168:     # ～150%
            return 48
        elif dpi < 216:     # ～200%弱
            return 56
        else:               # 200%以上
            return 64

    def _get_system_dpi(self) -> int:
        """システムのDPIを取得する"""
        try:
            if platform.system() != 'Windows':
                return 96
            # Windows 10+ DPI取得
            try:
                # 可能なら高精度の DPI 認識へ
                try:
                    shcore = ctypes.windll.shcore
                    # PROCESS_PER_MONITOR_DPI_AWARE = 2
                    shcore.SetProcessDpiAwareness(2)
                except Exception:
                    ctypes.windll.user32.SetProcessDPIAware()
                user32 = ctypes.windll.user32
                dpi = user32.GetDpiForWindow(self.window.winfo_id())  # 96, 120, 144, 192...
                return dpi
            except Exception:
                # フォールバック: スクリーンDPI推定
                return 96
        except Exception:
            return 96

    def load_icon_from_ico(self, ico_path: str, target_px: int) -> Image.Image:
        """ICOファイルから target_px に最も近いフレームを選んで返す"""
        try:
            im = Image.open(ico_path)
            best_index = 0
            best_size = (0, 0)
            n = getattr(im, "n_frames", 1)
            
            # まず利用可能なサイズを全て調べる
            available_sizes = []
            for i in range(n):
                im.seek(i)
                w, h = im.size
                available_sizes.append((i, w, h))
            
            # 最適なフレームを選択する戦略を改善
            # 1. 完全一致を優先
            # 2. target_px以上の最小サイズを優先（ダウンスケール）
            # 3. 最大サイズを選択（アップスケール）
            perfect_match = None
            larger_sizes = []
            
            for i, w, h in available_sizes:
                if w == target_px and h == target_px:
                    perfect_match = (i, w, h)
                    break
                elif w >= target_px and h >= target_px:
                    larger_sizes.append((i, w, h))
            
            if perfect_match:
                best_index, best_w, best_h = perfect_match
            elif larger_sizes:
                # target_px以上で最小のものを選択（ダウンスケールの方が高品質）
                best_index, best_w, best_h = min(larger_sizes, key=lambda x: x[1] * x[2])
            else:
                # target_px未満の場合は最大のものを選択
                best_index, best_w, best_h = max(available_sizes, key=lambda x: x[1] * x[2])
            
            im.seek(best_index)
            icon = im.convert("RGBA")
            
            # リサイズが必要な場合は高品質リサンプリングを使用
            if icon.size != (target_px, target_px):
                # 透明度を保持しながらリサイズ
                icon = icon.resize((target_px, target_px), Image.Resampling.LANCZOS)
            
            return icon
            
        except Exception as e:
            print(f"ICOファイル読み込みエラー: {e}")
            # フォールバック: 白い背景に「紙」の文字
            image = Image.new('RGBA', (target_px, target_px), (255, 255, 255, 255))
            draw = ImageDraw.Draw(image)
            # 境界線
            draw.rectangle([1, 1, target_px-2, target_px-2], outline=(0, 0, 0, 255), width=1)
            # 文字サイズを調整
            font_size = max(8, target_px // 3)
            try:
                # システムフォントを試す
                from PIL import ImageFont
                font = ImageFont.truetype("msgothic.ttc", font_size)
            except:
                font = None
            
            text_pos = (target_px // 4, target_px // 4)
            draw.text(text_pos, '紙', fill=(0, 0, 0, 255), font=font)
            return image
    
    def create_tray_icon(self):
        """システムトレイアイコンを作成（シンプル版）"""
        return self.create_tray_icon_simple()
    
    def create_tray_icon_simple(self):
        """シンプルで確実なICO読み込み"""
        print("=== システムトレイアイコン作成開始 ===")
        
        ico_path = self._resolve_ico_path()
        print(f"トレイアイコン用ICOパス: {ico_path}")
        
        if not ico_path or not os.path.exists(ico_path):
            print(f"❌ ICOファイルが見つかりません: {ico_path}")
            raise FileNotFoundError(f"ICOファイルが見つかりません: {ico_path}")
        
        print(f"ICOファイルサイズ: {os.path.getsize(ico_path)} bytes")
        
        # 直接ICOファイルを読み込み
        with Image.open(ico_path) as img:
            print(f"ICOフォーマット: {img.format}")
            
            # システムトレイに適したサイズのフレームを取得
            n_frames = getattr(img, "n_frames", 1)
            print(f"利用可能フレーム数: {n_frames}")
            
            # 32px付近のフレームを優先して選択
            target_sizes = [32, 48, 24, 64, 16]  # 優先順位
            best_frame = 0
            best_diff = float('inf')
            
            all_frames = []
            for i in range(n_frames):
                img.seek(i)
                w, h = img.size
                all_frames.append((i, w, h))
                print(f"  フレーム {i}: {w}x{h}px")
            
            # 32px付近を優先選択
            for target in target_sizes:
                for i, w, h in all_frames:
                    if w == target and h == target:
                        best_frame = i
                        print(f"完全一致フレーム発見: {target}x{target}px")
                        break
                else:
                    continue
                break
            else:
                # 完全一致がない場合、32pxに最も近いサイズを選択
                for i, w, h in all_frames:
                    diff = abs(w - 32)
                    if diff < best_diff:
                        best_diff = diff
                        best_frame = i
            
            print(f"選択フレーム: {best_frame} ({all_frames[best_frame][1]}x{all_frames[best_frame][2]}px)")
            
            # 最適フレームを読み込み
            img.seek(best_frame)
            icon = img.convert("RGBA")
            original_size = icon.size
            print(f"変換後サイズ: {original_size}")
            
            # 必要に応じて32pxにリサイズ
            if original_size != (32, 32):
                if original_size[0] > 64:
                    # 64pxより大きい場合は品質重視でリサイズ
                    print(f"大きなフレーム({original_size})を32pxにリサイズ")
                    icon = icon.resize((32, 32), Image.Resampling.LANCZOS)
                elif original_size[0] < 24:
                    # 24px未満は拡大
                    print(f"小さなフレーム({original_size})を32pxに拡大")
                    icon = icon.resize((32, 32), Image.Resampling.LANCZOS)
                else:
                    # 24-64pxの範囲はそのまま使用
                    print(f"適切なサイズ({original_size})のためリサイズしない")
                
                print(f"最終サイズ: {icon.size}")
            else:
                print("32pxフレーム - リサイズ不要")
            
            print("✅ トレイアイコン作成成功")
            return icon
    
    def init_tray(self):
        """システムトレイを初期化（重複防止版）"""
        try:
            print("=== システムトレイ初期化開始 ===")
            
            # 既存のトレイアイコンをチェック（デタッチ済みを含む）
            if self.tray_icon:
                # デタッチ済みのアイコンがある場合はそのまま使用
                if self._tray_detached:
                    print("既存のデタッチ済みトレイアイコンを使用 - 初期化をスキップ")
                    return self.tray_icon
                    
                # 表示中のアイコンがある場合もそのまま使用
                if hasattr(self.tray_icon, 'visible') and self.tray_icon.visible:
                    print("既存の表示中トレイアイコンが存在 - 初期化をスキップ")
                    return self.tray_icon
                
                # 古いアイコンをクリーンアップ
                print("古い非表示トレイアイコンをクリーンアップ中...")
                try:
                    self.tray_icon.stop()
                    import time
                    time.sleep(0.1)  # クリーンアップ待機
                except:
                    pass
                self.tray_icon = None
                self._tray_detached = False
            
            # ICOを直接読み込み
            print("トレイアイコン画像を作成中...")
            icon_image = self.create_tray_icon_simple()
            print(f"作成されたアイコンサイズ: {icon_image.size if icon_image else 'None'}")
            
            print("トレイメニューを作成中...")
            menu = pystray.Menu(
                pystray.MenuItem("表示", self.show_window, default=True),
                pystray.MenuItem("監視開始", self.start_watching_from_tray, enabled=lambda item: not self.is_watching),
                pystray.MenuItem("監視停止", self.stop_watching_from_tray, enabled=lambda item: self.is_watching),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("設定", self.show_window),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("終了", self.quit_application)
            )
            
            print("pystray.Iconオブジェクトを作成中...")
            self.tray_icon = pystray.Icon(
                "Kaminonaha",
                icon_image,
                "紙の名は。",
                menu
            )
            
            print("✅ システムトレイ初期化成功")
            return self.tray_icon
        except Exception as e:
            print(f"❌ システムトレイ初期化エラー: {e}")
            raise

    # システムトレイ切替機能はなし（最小化時のみ起動）

    def toggle_log_section(self):
        try:
            if getattr(self, '_log_collapsed', False):
                self.log_frame.pack(pady=20, padx=20, fill="both", expand=True)
                self.log_toggle_btn.config(text='詳細ログを隠す')
                self._log_collapsed = False
            else:
                self.log_frame.pack_forget()
                self.log_toggle_btn.config(text='詳細ログを表示')
                self._log_collapsed = True
        except Exception:
            pass

    def show_toast(self, title: str, message: str, duration_ms: int = 3000):
        try:
            tw = tk.Toplevel(self.window)
            tw.overrideredirect(True)
            tw.attributes('-topmost', True)
            bg = '#111827'
            fg = 'white'
            frame = tk.Frame(tw, bg=bg, padx=14, pady=10)
            frame.pack()
            tk.Label(frame, text=title, bg=bg, fg=fg, font=("Arial", 10, 'bold')).pack(anchor='w')
            tk.Label(frame, text=message, bg=bg, fg=fg, font=("Arial", 10)).pack(anchor='w')
            # 位置: 右下
            self.window.update_idletasks()
            x = self.window.winfo_rootx() + self.window.winfo_width() - 280
            y = self.window.winfo_rooty() + self.window.winfo_height() - 110
            tw.geometry(f"260x80+{x}+{y}")
            def _close():
                try:
                    tw.destroy()
                except Exception:
                    pass
            tw.after(duration_ms, _close)
        except Exception:
            pass
    
    def setup_gui(self):
        """GUI構築"""
        # トップバー（右側にAPI設定のみ）- 余白を最小化
        topbar = ttk.Frame(self.window, style='Top.TFrame')
        topbar.pack(fill='x', pady=(2, 0), padx=12)
        # 右上は何も置かない（API設定はシステム設定へ移動）
        right_tb = ttk.Frame(topbar, style='Top.TFrame')
        right_tb.pack(side='right')
        
        # ヘッダー（タイトル+サブ）
        header = tk.Frame(self.window, bg='white')
        header.pack(fill='x', padx=16, pady=(2,0))
        tk.Label(header, text='紙の名は。', bg='white', fg='#111', font=("Arial", 20, 'bold')).pack(anchor='w')
        tk.Label(header, text='フォルダを監視してAIがわかりやすい名前に', bg='white', fg='#6B7280', font=("Arial", 11)).pack(anchor='w', pady=(0,6))

        # タイトル下に開始/停止ボタン配置
        header_btns = ttk.Frame(header, style='Top.TFrame')
        header_btns.pack(fill='x', pady=(0,6))
        self.btn_start = ttk.Button(header_btns, text='▶ 開始', command=self.start_watching, style='Primary.TButton')
        self.btn_start.pack(side='left')
        self.btn_stop = ttk.Button(header_btns, text='⏹ 停止', command=self.stop_watching, style='Secondary.TButton')
        self.btn_stop.pack(side='left', padx=(8,0))
        try:
            self.btn_stop.state(['disabled'])
        except Exception:
            pass
        
        # ステータス（トップバー下に配置）
        self.status_label = tk.Label(
            self.window,
            text="監視停止中",
            font=("Arial", 12, "bold"),
            bg="white",
            fg="#666666"
        )
        self.status_label.pack(pady=(0, 8))
        
        # 監視フォルダ設定
        folder_frame = tk.LabelFrame(
            self.window,
            text="監視フォルダ管理",
            bg="white",
            font=("Arial", 12, "bold"),
            padx=15,
            pady=15
        )
        folder_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # フォルダリストとボタンのフレーム
        list_button_frame = tk.Frame(folder_frame, bg="white")
        list_button_frame.pack(fill="both", expand=True)
        
        # フォルダリスト（Treeview使用）
        list_frame = tk.Frame(list_button_frame, bg="white")
        list_frame.pack(side="left", fill="both", expand=True)
        
        # Treeviewでフォルダ一覧表示
        columns = ("path", "enabled", "settings")
        self.folder_tree = ttk.Treeview(list_frame, columns=columns, show="tree headings", height=10, style='Modern.Treeview')
        
        # カラム設定
        self.folder_tree.heading("#0", text="フォルダ")
        self.folder_tree.heading("enabled", text="監視")
        self.folder_tree.heading("settings", text="設定")
        
        self.folder_tree.column("#0", width=400)
        self.folder_tree.column("enabled", width=60)
        self.folder_tree.column("settings", width=120)
        
        # スクロールバー
        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.folder_tree.yview)
        self.folder_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.folder_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        
        # ボタンフレーム（縦配置で幅を広く）
        button_frame = tk.Frame(list_button_frame, bg="white")
        button_frame.pack(side="right", padx=(15, 0), fill="y")
        
        # フォルダ追加ボタン
        add_folder_btn = tk.Button(
            button_frame,
            text="フォルダ追加",
            command=self.add_watch_folder,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10),
            width=15
        )
        add_folder_btn.pack(pady=(0, 8))
        
        # フォルダ設定ボタン
        config_folder_btn = tk.Button(
            button_frame,
            text="フォルダ設定",
            command=self.configure_selected_folder,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10),
            width=15
        )
        config_folder_btn.pack(pady=(0, 8))
        
        # フォルダ削除ボタン
        remove_folder_btn = tk.Button(
            button_frame,
            text="フォルダ削除",
            command=self.remove_watch_folder,
            bg="#f44336",
            fg="white",
            font=("Arial", 10),
            width=15
        )
        remove_folder_btn.pack(pady=(0, 8))
        
        # 全削除ボタン
        clear_folders_btn = tk.Button(
            button_frame,
            text="全削除",
            command=self.clear_watch_folders,
            bg="#FF9800",
            fg="white",
            font=("Arial", 10),
            width=15
        )
        clear_folders_btn.pack()
        
        # 監視フォルダ一覧を更新
        self.update_folder_tree()
        
        # 旧制御ボタンはトップバーへ集約
        
        # システム設定
        system_frame = tk.LabelFrame(
            self.window,
            text="システム設定",
            bg="white",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=10
        )
        system_frame.pack(pady=15, padx=30, fill="x")
        
        # 設定項目を横に並べる
        system_options_frame = tk.Frame(system_frame, bg="white")
        system_options_frame.pack(fill="x")
        
        # スタートアップオプション
        startup_check = tk.Checkbutton(
            system_options_frame,
            text="🚀 Windows起動時に自動で開始",
            variable=self.auto_startup,
            font=("Arial", 10),
            bg="white",
            command=self.save_config
        )
        startup_check.pack(side="left", padx=(0, 20))
        
        # システムトレイ常駐は最小化時のみ（UI項目なし）
        
        # API設定ボタン（ここに集約）
        api_btn = tk.Button(
            system_options_frame,
            text="⚙ API設定",
            command=self.setup_api,
            font=("Arial", 10)
        )
        api_btn.pack(side='right')

        # インポート/エクスポートはUIから非表示（要望により整理）
        
        # ログ保持期間設定
        retention_frame = tk.Frame(system_frame, bg="white")
        retention_frame.pack(fill="x", pady=(10, 0))
        tk.Label(retention_frame, text="ログ保持（分）:", bg="white").pack(side="left")
        retention_spin = tk.Spinbox(retention_frame, from_=1, to=1440, width=6, textvariable=self.log_retention_minutes, command=self.save_config)
        retention_spin.pack(side="left", padx=(5, 20))
        # 最大ファイル名長
        tk.Label(retention_frame, text="最大ファイル名（文字）:", bg="white").pack(side="left")
        length_spin = tk.Spinbox(retention_frame, from_=20, to=80, width=6, textvariable=self.max_filename_length, command=self.save_config)
        length_spin.pack(side="left", padx=(5, 0))
        
        # ログ表示（常時表示）
        log_frame = tk.LabelFrame(
            self.window,
            text="処理ログ",
            bg="white",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=10
        )
        log_frame.pack(pady=12, padx=16, fill="both", expand=True)
        self.log_frame = log_frame
        
        # ログフォントは日本語向けのUIフォントを優先
        lfam = getattr(self, 'ui_font_family', 'Yu Gothic UI')
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=10,
            font=(lfam, 10),
            bg="white"
        )
        self.log_text.pack(fill="both", expand=True)
        
        # 初期ログメッセージ
        self.log_message("紙の名は。を起動しました")
        if not self.claude_client:
            self.log_message("⚠️ Claude APIキーが未設定です（システム設定→API設定から登録できます）")
            # 起動時にAPI設定ダイアログを案内
            try:
                self.window.after(600, self.show_api_setup_dialog)
            except Exception:
                pass
        
        folder_count = len(self.watch_folders)
        if folder_count > 0:
            self.log_message(f"📁 {folder_count}個のフォルダが登録されています")
    
    def update_folder_tree(self):
        """フォルダツリーを更新"""
        # 既存のアイテムを削除
        for item in self.folder_tree.get_children():
            self.folder_tree.delete(item)
        
        # フォルダ情報を追加
        for i, folder_info in enumerate(self.watch_folders):
            path = folder_info.get('path', '')
            enabled = "✓" if folder_info.get('enabled', True) else "✗"
            
            settings_info = []
            if folder_info.get('include_date', True):
                settings_info.append("日付")
            if folder_info.get('include_names', True):
                settings_info.append("名前")
            # 指示の有無を表示
            use_instr = bool(folder_info.get('use_custom_instruction', True))
            instr = (folder_info.get('custom_classify_prompt') or '').strip()
            has_instr = use_instr and bool(instr)
            settings_info.append("指示:あり" if has_instr else "指示:なし")
            
            settings_text = ", ".join(settings_info) if settings_info else "基本"
            
            self.folder_tree.insert(
                "",
                "end",
                iid=str(i),
                text=os.path.basename(path) or path,
                values=(path, enabled, settings_text)
            )

    def run_sample_test(self):
        """サンプルPDFを生成して最初の監視フォルダに投入"""
        try:
            if not self.watch_folders:
                messagebox.showwarning("警告", "先に監視フォルダを追加してください")
                return
            target = None
            for f in self.watch_folders:
                if f.get('enabled', True) and os.path.isdir(f.get('path','')):
                    target = f.get('path')
                    break
            if not target:
                messagebox.showwarning("警告", "有効な監視フォルダが見つかりません")
                return
            # PyMuPDFで簡易PDF生成
            doc = fitz.open()
            page = doc.new_page()
            text = "請求書\n株式会社テスト\n合計金額 12,300円\n2024/03/31"
            rect = fitz.Rect(50, 60, 550, 800)
            page.insert_textbox(rect, text, fontsize=14)
            os.makedirs(target, exist_ok=True)
            name = f"sample_{int(time.time())}.pdf"
            out = os.path.join(target, name)
            doc.save(out)
            doc.close()
            self.log_message(f"🧪 サンプルPDFを投入: {out}")
        except Exception as e:
            messagebox.showerror("エラー", f"サンプル作成に失敗: {e}")
    
    def add_watch_folder(self):
        """監視フォルダを追加"""
        folder = filedialog.askdirectory(title="監視するフォルダを選択")
        if folder:
            # 既存チェック
            for folder_info in self.watch_folders:
                if folder_info.get('path') == folder:
                    messagebox.showwarning("警告", "このフォルダは既に登録されています")
                    return
            
            # 新しいフォルダ情報を追加
            new_folder_info = {
                'path': folder,
                'enabled': True,
                'include_date': False,
                'include_names': False,
                'prompt_preset': 'auto',
                'prompt_keywords': ''
            }
            
            self.watch_folders.append(new_folder_info)
            self.update_folder_tree()
            self.save_config()
            self.log_message(f"📁 監視フォルダを追加: {folder}")
            
            # 監視中の場合は動的に監視対象に追加
            if self.is_watching and hasattr(self, 'observers'):
                self.add_folder_to_active_monitoring(new_folder_info)
    
    def add_folder_to_active_monitoring(self, folder_info):
        """アクティブな監視セッションに新しいフォルダを追加"""
        try:
            if not os.path.exists(folder_info['path']):
                self.log_message(f"⚠️ フォルダが見つかりません: {folder_info['path']}")
                return
            
            # 新しいObserverを作成して追加
            event_handler = PDFWatcherHandler(self)
            event_handler.folder_settings = folder_info
            
            observer = Observer()
            observer.schedule(event_handler, folder_info['path'], recursive=True)
            observer.start()
            
            # observers リストに追加
            if not hasattr(self, 'observers'):
                self.observers = []
            self.observers.append(observer)
            
            self.log_message(f"📁 監視に追加: {folder_info['path']}")
            
            # ステータス更新
            if hasattr(self, 'status_label'):
                folder_count = len([f for f in self.watch_folders if f.get('enabled', True) and os.path.exists(f.get('path', ''))])
                self.status_label.config(
                    text=f"監視中: {folder_count}個のフォルダ", 
                    fg="#4CAF50"
                )
                
        except Exception as e:
            self.log_message(f"❌ フォルダ監視追加エラー: {e}")
    
    def remove_folder_from_active_monitoring(self, folder_path):
        """アクティブな監視セッションから特定のフォルダを削除"""
        try:
            if not hasattr(self, 'observers'):
                return
            
            # 該当するObserverを見つけて停止
            observers_to_remove = []
            for i, observer in enumerate(self.observers):
                # Observerが監視している各patへアクセスしてチェック
                try:
                    # watchdog のObserver内部の監視パスをチェック
                    # _watchesからパスを取得
                    if hasattr(observer, '_watches'):
                        for watch in observer._watches.values():
                            if hasattr(watch, 'path') and watch.path == folder_path:
                                observer.stop()
                                observer.join(timeout=2)  # 最大2秒待機
                                observers_to_remove.append(i)
                                self.log_message(f"📁 監視停止: {folder_path}")
                                break
                except Exception:
                    # フォールバック: パス比較ができない場合は継続
                    pass
            
            # リストから削除（逆順で削除してインデックスの問題を回避）
            for i in sorted(observers_to_remove, reverse=True):
                if i < len(self.observers):
                    self.observers.pop(i)
            
            # ステータス更新
            if hasattr(self, 'status_label'):
                folder_count = len([f for f in self.watch_folders if f.get('enabled', True) and os.path.exists(f.get('path', ''))])
                self.status_label.config(
                    text=f"監視中: {folder_count}個のフォルダ", 
                    fg="#4CAF50"
                )
                
        except Exception as e:
            self.log_message(f"❌ フォルダ監視停止エラー: {e}")
    
    def update_folder_monitoring(self, old_folder_info, new_folder_info):
        """フォルダ設定変更時の監視状態を更新"""
        try:
            folder_path = old_folder_info['path']
            old_enabled = old_folder_info.get('enabled', True)
            new_enabled = new_folder_info.get('enabled', True)
            
            # 有効 → 無効の場合：監視を停止
            if old_enabled and not new_enabled:
                self.remove_folder_from_active_monitoring(folder_path)
                self.log_message(f"📁 監視無効化: {folder_path}")
            
            # 無効 → 有効の場合：監視を開始
            elif not old_enabled and new_enabled:
                self.add_folder_to_active_monitoring(new_folder_info)
                self.log_message(f"📁 監視有効化: {folder_path}")
            
            # 有効のまま設定変更の場合：一旦停止して再開始で設定更新
            elif old_enabled and new_enabled:
                self.remove_folder_from_active_monitoring(folder_path)
                # 少し待ってから再追加
                self.window.after(100, lambda: self.add_folder_to_active_monitoring(new_folder_info))
                self.log_message(f"📁 監視設定更新: {folder_path}")
                
        except Exception as e:
            self.log_message(f"❌ 監視状態更新エラー: {e}")
    
    def restart_after_settings_change(self):
        """設定変更後の監視再開"""
        try:
            self.log_message("▶️ 設定変更完了、監視を再開...")
            self.start_monitoring()
        except Exception as e:
            self.log_message(f"❌ 監視再開エラー: {e}")
    
    def configure_selected_folder(self):
        """選択されたフォルダの設定"""
        selection = self.folder_tree.selection()
        if selection:
            index = int(selection[0])
            folder_info = self.watch_folders[index]
            
            # 設定ダイアログを表示
            dialog = FolderSettingsDialog(self.window, folder_info)
            self.window.wait_window(dialog.dialog)
            
            if dialog.result:
                old_folder_info = folder_info.copy()
                self.watch_folders[index] = dialog.result
                self.update_folder_tree()
                self.save_config()
                self.log_message(f"⚙️ フォルダ設定を更新: {dialog.result['path']}")
                
                # 監視中の場合は一時停止して再開（二重監視を防止）
                if self.is_watching and hasattr(self, 'observers'):
                    self.log_message("⏸️ 設定変更のため監視を一時停止...")
                    was_watching = True
                    self.stop_monitoring()
                    # 短い遅延の後に再開
                    threading.Timer(1.0, lambda: self.restart_after_settings_change()).start()
                else:
                    was_watching = False
        else:
            messagebox.showwarning("警告", "設定するフォルダを選択してください")
    
    def remove_watch_folder(self):
        """選択されたフォルダを削除"""
        selection = self.folder_tree.selection()
        if selection:
            index = int(selection[0])
            folder_info = self.watch_folders[index]
            
            result = messagebox.askyesno("確認", f"以下のフォルダを削除しますか？\n\n{folder_info['path']}")
            if result:
                # 監視中の場合は該当するObserverを停止
                if self.is_watching and hasattr(self, 'observers'):
                    self.remove_folder_from_active_monitoring(folder_info['path'])
                
                self.watch_folders.pop(index)
                self.update_folder_tree()
                self.save_config()
                self.log_message(f"🗑️ 監視フォルダを削除: {folder_info['path']}")
        else:
            messagebox.showwarning("警告", "削除するフォルダを選択してください")
    
    def clear_watch_folders(self):
        """全てのフォルダを削除"""
        if self.watch_folders:
            result = messagebox.askyesno("確認", "全ての監視フォルダを削除しますか？")
            if result:
                folder_count = len(self.watch_folders)
                self.watch_folders.clear()
                self.update_folder_tree()
                self.save_config()
                self.log_message(f"🧹 全ての監視フォルダを削除（{folder_count}個）")
        else:
            messagebox.showinfo("情報", "削除するフォルダがありません")
    
    def start_watching(self):
        """監視開始（フォルダ別設定対応）"""
        if not self.watch_folders:
            messagebox.showerror("エラー", "監視フォルダが設定されていません")
            return
        
        if not self.claude_client:
            messagebox.showerror("エラー", "Claude APIキーを設定してください")
            return
        
        # 有効なフォルダをチェック
        valid_folders = []
        for folder_info in self.watch_folders:
            if folder_info.get('enabled', True) and os.path.exists(folder_info['path']):
                valid_folders.append(folder_info)
            elif folder_info.get('enabled', True):
                self.log_message(f"⚠️ フォルダが見つかりません: {folder_info['path']}")
        
        if not valid_folders:
            messagebox.showerror("エラー", "有効な監視フォルダがありません")
            return
        
        try:
            # フォルダごとにObserverを作成
            self.observers = []
            
            for folder_info in valid_folders:
                event_handler = PDFWatcherHandler(self)
                # フォルダ設定をハンドラーに渡す
                event_handler.folder_settings = folder_info
                
                observer = Observer()
                observer.schedule(event_handler, folder_info['path'], recursive=True)
                observer.start()
                self.observers.append(observer)
                self.log_message(f"📁 監視開始: {folder_info['path']}")
            
            self.is_watching = True
            
            # GUI更新
            try:
                self.btn_start.state(['disabled'])
                self.btn_stop.state(['!disabled'])
                # 停止ボタンを赤（Danger）に、開始はニュートラルのまま
                self.btn_stop.configure(style='Danger.TButton')
            except Exception:
                pass
            self.status_label.config(
                text=f"監視中: {len(valid_folders)}個のフォルダ", 
                fg="#4CAF50"
            )
            
            self.log_message(f"🔄 {len(valid_folders)}個のフォルダの監視を開始しました")
            
            self.config['auto_start_monitoring'] = True
            self.save_config()
            
        except Exception as e:
            self.log_message(f"❌ 監視開始エラー: {e}")
            messagebox.showerror("エラー", f"監視開始に失敗しました: {e}")
    
    def stop_watching(self):
        """監視停止"""
        if self.observers and self.is_watching:
            for observer in self.observers:
                observer.stop()
                observer.join()
            
            self.observers.clear()
            self.is_watching = False
            
            try:
                self.btn_start.state(['!disabled'])
                self.btn_stop.state(['disabled'])
                # ニュートラルへ戻す
                self.btn_stop.configure(style='Secondary.TButton')
            except Exception:
                pass
            self.status_label.config(text="監視停止中", fg="#666666")
            
            self.log_message("⏹️ 全フォルダの監視を停止しました")
            
            self.config['auto_start_monitoring'] = False
            self.save_config()
    
    def start_watching_from_tray(self):
        """トレイメニューから監視開始"""
        self.start_watching()
    
    def stop_watching_from_tray(self):
        """トレイメニューから監視停止"""
        self.stop_watching()
    
    def process_new_file(self, file_path):
        """新しいPDFファイルを処理（フォルダ別設定対応）"""
        try:
            filename = os.path.basename(file_path)
            folder_path = os.path.dirname(file_path)
            folder_name = os.path.basename(folder_path)
            
            # ファイルが属するフォルダの設定を取得
            folder_settings = None
            for folder_info in self.watch_folders:
                if folder_path.startswith(folder_info['path']):
                    folder_settings = folder_info
                    break
            
            if not folder_settings:
                folder_settings = {
                    'include_date': True,
                    'include_names': True,
                    'use_custom_output': False,
                    'output_folder': ''
                }
            
            self.log_message(f"🔄 処理開始: {filename} ({folder_name})")
            
            if not os.path.exists(file_path):
                self.log_message(f"❌ ファイルが見つかりません: {filename}")
                return
            
            # PDFを画像に変換（先頭2ページまで）
            images = self.pdf_to_images(file_path, max_pages=2)
            if not images:
                self.log_message(f"❌ PDF変換失敗: {filename}")
                return
            
            prompt_override = None
            try:
                if folder_settings.get('use_custom_instruction', True):
                    c = (folder_settings.get('custom_classify_prompt') or '').strip()
                    prompt_override = c if c else None
            except Exception:
                prompt_override = (folder_settings.get('custom_classify_prompt') or None)
            extracted_text = self.extract_text_from_pdf(file_path, max_pages=2, max_chars=4000)

            # まず文書種別を軽く判定（登記事項系の特別処理用）
            self.log_message(f"🔎 種別判定: {filename}")
            preset_key_for_labels = folder_settings.get('prompt_preset', 'auto')
            if extracted_text and len(extracted_text) >= 200:
                doc_type = self.classify_with_text(extracted_text, prompt_override=prompt_override, preset_key=preset_key_for_labels)
            else:
                doc_type = self.classify_with_vision(images, prompt_override=prompt_override, preset_key=preset_key_for_labels)
            doc_type = (doc_type or '').strip()

            # 主たる/従たるの補正（計算書+資料などは主たる書類名に寄せる）
            try:
                doc_type = self.adjust_primary_document_type(extracted_text or '', doc_type)
            except Exception:
                pass

            # 宛名・日付はオプションで抽出
            names_info = {'surname': None, 'given_name': None, 'company_name': None}
            if folder_settings.get('include_names', False):
                names_info = self.extract_names_and_companies(images[0])
            document_date = None
            if folder_settings.get('include_date', False):
                document_date = datetime.now().strftime("%Y%m%d")

            # 登記事項証明系なら、不動産情報を抽出して専用命名
            registry_keywords = ['登記事項証明書', '登記情報', '登記簿', '全部事項証明書', '現在事項証明書', '建物事項証明書', '土地登記', '建物登記', '不動産登記']
            if any(k in doc_type for k in registry_keywords):
                self.log_message("🏷 登記系書類と判定 → 不動産情報を抽出")
                property_info = self.extract_property_info(images[0], doc_type)
                # ファイルリネーム（document_typeは固定で登記事項証明書を採用）
                new_path = self.rename_file(
                    file_path, '登記事項証明書', names_info, document_date, property_info, folder_settings
                )
            else:
                # 主: 1ページ目のタイトル重視 → 失敗時は全体から推定
                self.log_message(f"🧠 AI自由命名: {filename}")
                # レイアウト優先：上部の大きな文字を優先してタイトル候補に
                layout_title = self.extract_layout_title(file_path)
                base_name = layout_title
                if not base_name:
                    first_text = self.extract_text_from_pdf(file_path, max_pages=1, max_chars=1000)
                    if first_text and len(first_text) >= 40:
                        base_name = self.ai_name_from_text(first_text, prompt_override)
                if not base_name:
                    base_name = self.ai_name_from_vision([images[0]], prompt_override)
                if not base_name:
                    if extracted_text and len(extracted_text) >= 120:
                        base_name = self.ai_name_from_text(extracted_text, prompt_override)
                    else:
                        base_name = self.ai_name_from_vision(images, prompt_override)
                # AIが短く切った場合はレイアウトの候補で上書き（先頭一致）
                try:
                    if layout_title and base_name:
                        lt = re.sub(r"\s+", " ", layout_title).strip()
                        bn = base_name.strip()
                        if lt and bn and lt.upper().startswith(bn.upper()) and len(lt) <= 64:
                            base_name = lt
                except Exception:
                    pass
                if not base_name:
                    self.log_message(f"❌ 自由命名失敗: {filename}")
                    return
                # 連結（ベース名 + 任意追記）
                final_name = base_name
                if folder_settings.get('include_names', False):
                    if names_info.get('company_name'):
                        final_name += f"_{names_info['company_name']}"
                    elif names_info.get('surname') and names_info.get('given_name'):
                        final_name += f"_{names_info['surname']}{names_info['given_name']}"
                    elif names_info.get('surname'):
                        final_name += f"_{names_info['surname']}"
                if folder_settings.get('include_date', False):
                    if not re.search(r'(19|20)\d{6}', final_name):
                        final_name += f"_{document_date}"

                document_type = self.sanitize_filename(final_name)
                new_path = self.rename_file(
                    file_path, document_type, None, None, None, folder_settings
                )
            
            if new_path:
                new_filename = os.path.basename(new_path)
                self.log_message(f"✅ 成功: {filename} → {new_filename} ({folder_name})")
                
                # システムトレイ通知（安全版）
                self.safe_notify("PDF処理完了", f"→ {new_filename}")
                # トースト通知（控えめ）
                try:
                    self.show_toast('PDF処理完了', f"{new_filename}")
                except Exception:
                    pass
            else:
                self.log_message(f"❌ リネーム失敗: {filename}")
                
        except Exception as e:
            self.log_message(f"❌ 処理エラー: {filename} - {e}")

    def extract_layout_title(self, pdf_path: str) -> str | None:
        """1ページ目のレイアウトからタイトル候補を抽出。
        - スパンのフォントサイズを集計し、"大きめ"の文字群を抽出
        - 大きめの中では画面上部（yが小さい）を優先
        - ノイズ（あまりに短い/英数字のみ等）を除外
        """
        try:
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                return None
            page = doc[0]
            info = page.get_text('dict')
            lines_agg = []  # (y0, rep_size, full_line_text)
            for block in info.get('blocks', []):
                for line in block.get('lines', []):
                    spans = line.get('spans', []) or []
                    if not spans:
                        continue
                    # line bbox は spans/bbox から推測
                    y0s, sizes, texts = [], [], []
                    for span in spans:
                        text = (span.get('text') or '')
                        if not text.strip():
                            continue
                        texts.append(text)
                        sizes.append(float(span.get('size') or 0))
                        bbox = span.get('bbox') or [0,0,0,0]
                        y0s.append(bbox[1])
                    if not texts:
                        continue
                    full_text = ''.join(texts).strip()
                    # クリーンアップ
                    full_text = re.sub(r'[\s\u3000]+', ' ', full_text)
                    tclean = re.sub(r'\s+', '', full_text)
                    if len(tclean) < 2:
                        continue
                    if re.fullmatch(r'[\W_]+', tclean):
                        continue
                    if re.fullmatch(r'[0-9\-–—/\.]+', tclean):
                        continue
                    y0 = min(y0s) if y0s else 0.0
                    # 表示サイズの代表値は最大（見出しを優先）
                    rep_size = max(sizes) if sizes else 0.0
                    lines_agg.append((y0, rep_size, full_text))

            if not lines_agg:
                return None
            sizes = [s for _, s, _ in lines_agg]
            med = stats.median(sizes)
            thr = max(12.0, med * 1.25)
            large = [(y, sz, tx) for (y, sz, tx) in lines_agg if sz >= thr]
            if not large:
                # 代替: 上位サイズ10件から最上部を優先
                topN = sorted(lines_agg, key=lambda t: t[1], reverse=True)[:10]
                cand = sorted(topN, key=lambda t: (t[0], -t[1]))[0]
                title = cand[2]
                # 行継続（ハイフン区切りや同サイズで直下行が続く場合）
                title = self._maybe_join_next_line(title, cand, lines_agg)
            else:
                # 大きめ文字の中で最上部（yが小）を採用
                cand = sorted(large, key=lambda t: (t[0], -t[1]))[0]
                title = cand[2]
                title = self._maybe_join_next_line(title, cand, lines_agg)
            # 軽クリーニング
            title = title.strip().splitlines()[0]
            title = re.sub(r'[\s　]+', ' ', title)
            cleaned = self.clean_ai_filename_output(title)
            # 過度に短くなる場合は元の行テキストを優先
            title = cleaned if len(cleaned) >= 4 else title
            # 英単語途中の不自然な空白を除去（例: Organizat ion → Organization）
            title = re.sub(r'(?<=[a-z])\s+(?=[a-z])', '', title)
            title = self.sanitize_filename(title)
            doc.close()
            # 極端に短い場合は見なさない
            return title if len(title) >= 2 else None
        except Exception:
            try:
                doc.close()
            except Exception:
                pass
            return None

    def _maybe_join_next_line(self, title: str, cand: tuple, lines_agg: list[tuple]) -> str:
        """候補行の直下行が同等サイズで続きなら結合（ハイフン改行や単語分割対策）。"""
        try:
            y0, sz, text = cand
            # 直下行の候補を探索（y差が小さく、サイズ差が小さい）
            below = [t for t in lines_agg if (t[0] > y0 and abs(t[1]-sz) <= max(1.0, sz*0.1))]
            if not below:
                return title
            nxt = sorted(below, key=lambda t: t[0])[0]
            next_text = nxt[2].strip()
            # ハイフン改行や明らかな単語継続のとき結合
            if title.rstrip().endswith(('-', '‐', '‑', '–', '—')):
                return re.sub(r"[-‐‑–—]+\s*$", "", title) + next_text
            # 英字のみで、次行が英字で始まる場合も連結（スペースあり）
            if re.fullmatch(r"[A-Za-z0-9 \-]+", title) and re.match(r"^[A-Za-z]", next_text):
                # 末尾が小文字、先頭が小文字なら単語継続とみなしてスペースなし結合
                if re.search(r"[a-z]$", title) and re.match(r"^[a-z]", next_text):
                    return title + next_text
                # それ以外はスペースで結合
                return title + " " + next_text
            return title
        except Exception:
            return title

    def adjust_primary_document_type(self, text: str, ai_label: str) -> str:
        """全体テキストから主たる書類を推定し、AIのラベルを軽く補正する。
        例: 『計算書 + 資料』→『計算書』優先など。
        """
        try:
            label = (ai_label or '').strip()
            t = (text or '').lower()
            # 日本語キーワードの単純検出（大文字小文字鈍感にするためlower）
            # 主たる書類候補とキーワード
            primary = [
                ('計算書', ['計算書', '損益計算書', '貸借対照表', '決算書', '財務諸表']),
                ('契約書', ['契約書', '覚書', '合意書']),
                ('見積書', ['見積書']),
                ('請求書', ['請求書', '請求金額']),
                ('領収書', ['領収書', '受領']),
                ('納品書', ['納品書']),
            ]
            # 従たる/汎用キーワード
            secondary_words = ['資料', '添付資料', '参考資料', '別紙', '付録', '別添']

            def count_any(words):
                c = 0
                for w in words:
                    if w.lower() in t:
                        c += t.count(w.lower())
                return c

            # 二次的な語が多く、AIが『資料/その他』と判断した場合は主たる候補を再評価
            if any(sw in label for sw in ['資料', 'その他']) or label in ['', 'PDF文書']:
                best = (label, 0)
                for name, kws in primary:
                    c = count_any(kws)
                    if c > best[1]:
                        best = (name, c)
                if best[1] >= 1:
                    return best[1] and best[0] or label

            # AIが既に主たる名を返している場合でも、明らかな矛盾（資料多すぎ）はスキップ
            # もしくはAIが『受付のお知らせ』等の場合、財務系語が強ければ『計算書』へ
            if label in ['受付のお知らせ', '必要書類等一覧', 'PDF文書', 'その他書類']:
                calc_score = count_any(['計算書', '損益計算書', '貸借対照表', '決算書'])
                if calc_score >= 1:
                    return '計算書'
            return label
        except Exception:
            return ai_label

    # 旧・ルールベース名の適用は廃止

    def normalize_document_type(self, document_type: str) -> str:
        """AIの分類名を内蔵ルールで標準化（軽量辞書・部分一致/正規表現）"""
        if not document_type:
            return "PDF文書"
        dt = document_type.strip()
        try:
            mapping = [
                # 司法書士系の正規化
                (r"印鑑証明", "印鑑証明書"),
                (r"(登記事項証明|全部事項証明|現在事項証明)", "登記事項証明書"),
                # 一般書類
                (r"請求(書)?", "請求書"),
                (r"見積(書)?", "見積書"),
                (r"領収(書)?", "領収書"),
                (r"納品(書)?", "納品書"),
                (r"注文(書)?", "注文書"),
                (r"契約(書)?", "契約書"),
            ]
            import re as _re
            for pat, name in mapping:
                if _re.search(pat, dt):
                    return name
        except Exception:
            pass
        return dt

    def sanitize_filename(self, name: str, max_len: int | None = None) -> str:
        try:
            s = name.strip()
            orig = s
            # 禁止文字の除去
            s = re.sub(r'[<>:"/\\|?*]', '', s)
            # 制御文字の除去（U+0000〜U+001F）
            s = re.sub(r'[\x00-\x1F]', '', s)
            # 改行やタブをスペースに
            s = re.sub(r'[\r\n\t]+', ' ', s)
            # 連続スペースの縮約
            s = re.sub(r'\s{2,}', ' ', s)
            # 先頭末尾のピリオドやスペース除去
            s = s.strip(' .')
            # Windows予約語の避け（拡張子があっても不可なので根幹名を検査）
            try:
                root = s.split('.')[0].strip().upper()
            except Exception:
                root = s.strip().upper()
            reserved = {"CON","PRN","AUX","NUL","COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9","LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"}
            if root in reserved:
                s = f"_{s}" if s else "文書"
            # 長さ制限（設定優先）
            try:
                conf_limit = int(self.max_filename_length.get()) if hasattr(self, 'max_filename_length') else int(self.config.get('max_filename_length', 40))
            except Exception:
                conf_limit = int(self.config.get('max_filename_length', 40))
            limit = max_len if isinstance(max_len, int) and max_len > 0 else conf_limit
            # クランプ（20〜80の範囲で安全に）
            limit = min(80, max(20, limit))
            if len(s) > limit:
                # デリミタに合わせて気持ちよく切る（なければハードカット）
                cut = -1
                for d in ['　', '、', '（', '(', '・', ' ', '-', '—', '–', '_']:
                    p = s.rfind(d, 0, limit)
                    if p > cut:
                        cut = p
                if cut >= 10:  # 極端に短くならないように
                    s = s[:cut].rstrip(' 　、（(・-—–_')
                else:
                    s = s[:limit]
                # 可能なら省略記号を付与
                if len(orig) > len(s) and len(s) < limit:
                    s = (s[:max(0, limit-1)] + '…')[:limit]
            return s or '名称未設定'
        except Exception:
            return '名称未設定'

    def ai_name_from_text(self, text: str, prompt_override: str | None) -> str | None:
        try:
            base_prompt = (
                "この文書の主たるページ（1ページ目）に記載の『タイトルまたは種類名』を1つだけ返してください。\n"
                "条件: 句読点・説明なし、名詞句のみ。\n"
                "例: 見積書 / 契約書 / 登記事項証明書 / 受付のお知らせ\n"
                "ファイル名に不適切な記号 / \\ : * ? \" < > | は使わないこと。\n\n"
                f"【文書内容（抜粋）】\n{text[:1200]}\n"
            )
            prompt = (prompt_override + "\n\n" + base_prompt) if prompt_override else base_prompt
            message = self._anthropic_call_with_retry(
                [{"type": "text", "text": prompt}], max_tokens=64, temperature=0, timeout=30.0
            )
            resp = (message.content[0].text or '').strip()
            cleaned = self.clean_ai_filename_output(resp)
            return self.sanitize_filename(cleaned)
        except Exception as e:
            print(f"AI自由命名(テキスト) エラー: {e}")
            self.log_message(f"❌ 自由命名(テキスト) エラー: {str(e)}")
            return None

    def ai_name_from_vision(self, images, prompt_override: str | None) -> str | None:
        try:
            images = images if isinstance(images, (list, tuple)) else [images]
            img_blocks = []
            for img in images[:2]:
                buffer = io.BytesIO()
                img.save(buffer, format='PNG')
                image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
                img_blocks.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": image_data
                    }
                })
            base_prompt = (
                "この文書の主たるページ（1ページ目）に記載の『タイトルまたは種類名』を1つだけ返してください。\n"
                "条件: 句読点・説明なし、名詞句のみ。\n"
                "例: 見積書 / 契約書 / 登記事項証明書 / 受付のお知らせ\n"
                "ファイル名に不適切な記号 / \\ : * ? \" < > | は使わないこと。"
            )
            prompt = (prompt_override + "\n\n" + base_prompt) if prompt_override else base_prompt
            message = self._anthropic_call_with_retry(
                ([{"type": "text", "text": prompt}] + img_blocks), max_tokens=64, temperature=0, timeout=30.0
            )
            resp = (message.content[0].text or '').strip()
            cleaned = self.clean_ai_filename_output(resp)
            return self.sanitize_filename(cleaned)
        except Exception as e:
            print(f"AI自由命名(Vision) エラー: {e}")
            self.log_message(f"❌ 自由命名(Vision) エラー: {str(e)}")
            return None

    def clean_ai_filename_output(self, resp: str) -> str:
        try:
            s = (resp or '').strip()
            if not s:
                return ''
            # 1行目のみ
            s = s.splitlines()[0].strip()
            # 接頭辞の除去
            s = re.sub(r'^(ファイル名|タイトル|題名)[:：]\s*', '', s)
            # 「この文書は〜」以降を削除
            s = re.sub(r'この文書は.*$', '', s)
            # 最初の句点まで
            s = s.split('。')[0].strip()
            # コロンで説明が続く場合は左側を優先（ダッシュやハイフンはタイトルに含める）
            parts = re.split(r'\s*[:：]\s*', s)
            if parts:
                s = parts[0].strip()
            # 引用符の除去
            s = s.strip('「」"\'')
            # 冗長な語尾の簡易削除
            for tail in ['について', 'に関する', 'に係る', 'のご案内', 'の案内', 'の通知', 'のお願い']:
                if s.endswith(tail) and len(s) > 8:
                    s = s[: -len(tail)]
                    break
            return s
        except Exception:
            return resp or ''
    
    def rename_file(self, original_path, document_type, names_info=None, document_date=None, property_info=None, folder_settings=None):
        """ファイルをリネーム（フォルダ別設定対応）"""
        try:
            directory = os.path.dirname(original_path)
            name_parts = [document_type]
            
            # 登記関連書類の特殊処理：不動産地番等のみ
            if property_info and property_info.get('type'):
                # 不動産種別の略称を決定
                property_type_suffix = ""
                if property_info['type'] == '土地':
                    property_type_suffix = "（土地）"
                elif property_info['type'] == '建物':
                    property_type_suffix = "（建物）"
                elif property_info['type'] == '区分建物':
                    property_type_suffix = "（区分）"
                
                # 不動産情報から所在地と地番等を取得
                location = property_info.get('location', '')
                address_number = property_info.get('address_number', '')
                
                # 区分建物の場合、建物名称+部屋番号のみ表示
                if property_info['type'] == '区分建物':
                    if address_number:
                        # 建物名称+部屋番号のみ
                        property_part = f"{address_number}{property_type_suffix}"
                    else:
                        property_part = f"区分建物{property_type_suffix}"
                else:
                    # 土地・建物の場合
                    if location and address_number:
                        property_part = f"{location}{address_number}{property_type_suffix}"
                    elif location:
                        property_part = f"{location}{property_type_suffix}"
                    elif address_number:
                        property_part = f"{address_number}{property_type_suffix}"
                    else:
                        property_part = f"不動産情報{property_type_suffix}"
                
                name_parts.append(property_part)
            
            # 一般書類の名前・法人名を追加（登記関連書類でない場合のみ）
            elif folder_settings and folder_settings.get('include_names', True) and names_info:
                # 法人名を最優先
                if names_info.get('company_name'):
                    name_parts.append(names_info['company_name'])
                # 個人名を追加（姓・名の両方がある場合）
                elif names_info.get('surname') and names_info.get('given_name'):
                    full_name = f"{names_info['surname']}{names_info['given_name']}"
                    name_parts.append(full_name)
                # 姓のみの場合
                elif names_info.get('surname'):
                    name_parts.append(names_info['surname'])
            
            # 日付を追加（設定が有効で、登記関連書類でない場合）
            if folder_settings and folder_settings.get('include_date', False) and not (property_info and property_info.get('type')):
                # 既に8桁日付が含まれていれば重複を避ける
                pre = "_".join(name_parts)
                if not re.search(r'(19|20)\d{6}', pre):
                    date_to_use = document_date if document_date else datetime.now().strftime("%Y%m%d")
                    name_parts.append(date_to_use)
            
            base_filename = "_".join(name_parts)
            base_filename = self.sanitize_filename(base_filename)
            
            # 出力先ディレクトリを決定
            if folder_settings and folder_settings.get('use_custom_output', False):
                output_folder = folder_settings.get('output_folder', '')
                if output_folder and os.path.exists(output_folder):
                    directory = output_folder
            
            # 重複回避
            counter = 1
            while True:
                if counter == 1:
                    new_filename = f"{base_filename}.pdf"
                else:
                    new_filename = f"{base_filename}_{counter}.pdf"
                
                new_path = os.path.join(directory, new_filename)
                
                if not os.path.exists(new_path):
                    break
                counter += 1
            
            # ファイルを移動またはリネーム
            if directory != os.path.dirname(original_path):
                # 別フォルダに移動
                import shutil
                shutil.move(original_path, new_path)
            else:
                # 同じフォルダ内でリネーム
                os.rename(original_path, new_path)
            
            return new_path
            
        except Exception as e:
            print(f"リネームエラー: {e}")
            return None
    
    def log_message(self, message):
        """ログメッセージを表示（保持期間付き）"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_line = f"[{timestamp}] {message}\n"
        def _append():
            try:
                # 内部履歴
                self.log_history.append((time.time(), log_line))
                # 表示
                self.log_text.insert(tk.END, log_line)
                self.log_text.see(tk.END)
                # 軽い即時剪定
                self._prune_log(max_lines=2000)
            except Exception:
                pass
        self.window.after(0, _append)

    def _prune_log(self, max_lines=2000):
        try:
            now = time.time()
            keep_sec = max(60, int(self.log_retention_minutes.get()) * 60)
            # 時間でフィルタ
            self.log_history = [(t, l) for t, l in self.log_history if now - t <= keep_sec]
            # 行数で上限
            if len(self.log_history) > max_lines:
                self.log_history = self.log_history[-max_lines:]
            # テキストに反映
            self.log_text.delete('1.0', tk.END)
            self.log_text.insert(tk.END, "".join(l for _, l in self.log_history))
            self.log_text.see(tk.END)
            # 設定へ保持時間を反映
            self.config['log_retention_minutes'] = int(self.log_retention_minutes.get())
        except Exception:
            pass

    def _prune_log_periodic(self):
        try:
            self._prune_log()
        finally:
            try:
                self.window.after(60_000, self._prune_log_periodic)
            except Exception:
                pass
    
    # PDF処理関連メソッド（既存と同じ）
    def pdf_to_image(self, pdf_path):
        """PDFの1ページ目を画像に変換（互換API）"""
        imgs = self.pdf_to_images(pdf_path, max_pages=1)
        return imgs[0] if imgs else None

    def pdf_to_images(self, pdf_path, max_pages=2):
        """PDFの先頭max_pagesページを画像に変換（適応的解像度処理）"""
        try:
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                return []
            images = []
            page_count = min(len(doc), max_pages)
            for i in range(page_count):
                page = doc[i]
                # まず元サイズをチェック
                pix_check = page.get_pixmap()
                check_size = len(pix_check.tobytes("png"))
                # サイズに応じて適応的に解像度調整
                if check_size < 500000:  # 0.5MB未満 = 小さすぎる
                    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                elif check_size < 1500000:  # 1.5MB未満 = やや小さい
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                else:  # 1.5MB以上 = 十分大きい
                    pix = page.get_pixmap()

                img_data = pix.tobytes("png")
                final_size = len(img_data)
                if final_size <= 5000000:  # 5MB以下
                    images.append(Image.open(io.BytesIO(img_data)))
                else:
                    image = Image.open(io.BytesIO(img_data))
                    images.append(self.light_compress_for_api(image))
            doc.close()
            return images
        
        except Exception as e:
            print(f"PDF変換エラー: {e}")
            return []
    
    def light_compress_for_api(self, image):
        """5MB超過時の軽い圧縮（品質重視）"""
        try:
            max_size = 4900000  # 4.9MB（安全マージン）
            
            # RGBAの場合はRGBに変換
            if image.mode == 'RGBA':
                rgb_image = Image.new('RGB', image.size, (255, 255, 255))
                rgb_image.paste(image, mask=image.split()[-1])
                image = rgb_image
                print("RGBA→RGB変換完了")
            
            # 高品質のJPEG圧縮から開始
            quality = 95
            current_size = max_size + 1
            
            while current_size > max_size and quality >= 80:
                buffer = io.BytesIO()
                image.save(buffer, format='JPEG', quality=quality, optimize=True)
                current_size = len(buffer.getvalue())
                
                print(f"軽圧縮テスト: {current_size} bytes (品質: {quality})")
                
                if current_size <= max_size:
                    buffer.seek(0)
                    compressed_image = Image.open(buffer)
                    print(f"✅ 軽圧縮完了: {current_size} bytes (品質: {quality})")
                    return compressed_image
                
                quality -= 5  # 細かく調整
            
            # それでも大きい場合は従来の強圧縮
            print("強圧縮に切り替え")
            return self.compress_image_for_api(image)
            
        except Exception as e:
            print(f"軽圧縮エラー: {e}")
            return self.compress_image_for_api(image)
    
    def compress_image_for_api(self, image):
        """Claude API用に画像を圧縮（5MB制限対応・強制圧縮）"""
        try:
            # 5MB制限（安全のため4.5MBを上限とする）
            max_size = 4700000  # 4.7MB
            
            # 最初に画像サイズを大幅に縮小
            width, height = image.size
            print(f"元画像サイズ: {width}x{height}")
            
            # まず画像を800x800以下に縮小
            if width > 800 or height > 800:
                ratio = min(800/width, 800/height)
                new_width = int(width * ratio)
                new_height = int(height * ratio)
                image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                print(f"サイズ縮小: {new_width}x{new_height}")
            
            # RGBAの場合はRGBに変換
            if image.mode == 'RGBA':
                rgb_image = Image.new('RGB', image.size, (255, 255, 255))
                rgb_image.paste(image, mask=image.split()[-1])
                image = rgb_image
                print("RGBA→RGB変換完了")
            
            # JPEG形式で段階的に品質を下げて圧縮
            quality = 85
            current_size = max_size + 1
            
            while current_size > max_size and quality >= 30:
                buffer = io.BytesIO()
                image.save(buffer, format='JPEG', quality=quality, optimize=True)
                current_size = len(buffer.getvalue())
                
                print(f"圧縮テスト: {current_size} bytes (品質: {quality})")
                
                if current_size <= max_size:
                    # 圧縮成功
                    buffer.seek(0)
                    compressed_image = Image.open(buffer)
                    print(f"✅ 圧縮成功: {current_size} bytes (品質: {quality})")
                    return compressed_image
                
                # 品質を下げて再試行
                quality -= 10
            
            # それでもサイズが大きい場合は、さらに画像サイズを縮小
            if current_size > max_size:
                width, height = image.size
                while current_size > max_size and (width > 400 or height > 400):
                    ratio = 0.8  # 20%縮小
                    new_width = int(width * ratio)
                    new_height = int(height * ratio)
                    image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    width, height = new_width, new_height
                    
                    buffer = io.BytesIO()
                    image.save(buffer, format='JPEG', quality=50, optimize=True)
                    current_size = len(buffer.getvalue())
                    
                    print(f"さらに縮小: {current_size} bytes (サイズ: {width}x{height})")
                
                buffer.seek(0)
                final_image = Image.open(buffer)
                print(f"✅ 最終圧縮: {current_size} bytes")
                return final_image
            
            return image
            
        except Exception as e:
            print(f"画像圧縮エラー: {e}")
            # エラー時は元画像をJPEG形式で最低品質で保存
            try:
                buffer = io.BytesIO()
                if image.mode == 'RGBA':
                    rgb_image = Image.new('RGB', image.size, (255, 255, 255))
                    rgb_image.paste(image, mask=image.split()[-1])
                    image = rgb_image
                image.save(buffer, format='JPEG', quality=30, optimize=True)
                buffer.seek(0)
                return Image.open(buffer)
            except:
                return image
    
    def get_model(self):
        """UI上のモデル名/エイリアスを実モデルIDに解決"""
        configured = self.config.get('model', self.model_name or 'claude-sonnet-4-20250514')
        alias_map = {
            # Claude 4（UI表記）→ 正式API ID へ解決
            'Claude 4 Sonnet': 'claude-sonnet-4-20250514',
            'claude-4-sonnet': 'claude-sonnet-4-20250514',
        }
        return alias_map.get(configured, configured)

    def build_label_set(self, preset_key: str | None) -> list[str]:
        p = (preset_key or 'auto').lower()
        common = [
            "登記事項証明書", "印鑑証明書",
            "契約書", "覚書", "議事録", "定款", "委任状", "就任承諾書",
            "見積書", "請求書", "領収書", "注文書", "納品書", "仕様書", "送付状",
            "受付のお知らせ", "必要書類等一覧", "その他書類"
        ]
        if p == 'business':
            return ["見積書", "請求書", "領収書", "注文書", "納品書", "仕様書", "送付状", "請求明細", "その他書類"]
        if p == 'legal':
            return ["契約書", "覚書", "規程", "議事録", "定款", "委任状", "就任承諾書", "印鑑証明書", "その他書類"]
        if p == 'realestate':
            return ["登記事項証明書", "重要事項説明", "売買契約書", "賃貸借契約書", "不動産契約書", "その他書類"]
        return common

    def _anthropic_call_with_retry(self, content_blocks, *, max_tokens=100, temperature=0, timeout=30.0,
                                   models=None, retries=3):
        """Anthropic API呼び出し（モデルフォールバック + 529混雑時の指数バックオフ）"""
        if not self.claude_client:
            raise RuntimeError('Claude API未設定')
        primary = self.get_model()
        fb = 'claude-3-5-sonnet-20241022'
        try_models = models or ([primary] + ([fb] if fb != primary else []))
        last_err = None
        for m in try_models:
            delay = 1.5
            for attempt in range(retries):
                try:
                    return self.claude_client.messages.create(
                        model=m,
                        max_tokens=max_tokens,
                        temperature=temperature,
                        messages=[{"role": "user", "content": content_blocks}],
                        timeout=timeout
                    )
                except Exception as e:
                    last_err = e
                    es = str(e)
                    overloaded = ('529' in es) or ('overload' in es.lower())
                    if overloaded and attempt < retries - 1:
                        wait = delay * (1 + 0.25 * random.random())
                        self.log_message(f"⏳ モデル混雑のため再試行({attempt+1}/{retries-1}) {wait:.1f}s 待機: {m}")
                        time.sleep(wait)
                        delay *= 2
                        continue
                    # それ以外 or 最終試行は次のモデルへ
                    self.log_message(f"⚠️ モデル '{m}' 呼び出し失敗: {e}")
                    break
        # 全モデル失敗
        raise last_err if last_err else RuntimeError('Anthropic呼び出し失敗')

    def classify_with_text(self, text, prompt_override=None, preset_key=None):
        """テキストのみで文書分類（高精度・簡潔プロンプト）"""
        try:
            label_list = self.build_label_set(preset_key)
            labels = "、".join(label_list)
            base_prompt = (
                f"以下の文書テキストを読み、1ページ目を主とみなして、"
                f"次の中から最も適切な1つの文書種別のみを日本語で返してください（説明不要・厳密一致）。\n"
                f"候補: {labels}\n"
                f"どれにも当てはまらなければ『その他書類』と返してください。\n\n"
                f"【文書テキスト】\n{text[:4000]}"
            )
            prompt = (prompt_override + "\n\n" + base_prompt) if prompt_override else base_prompt

            message = self._anthropic_call_with_retry(
                [{"type": "text", "text": prompt}], max_tokens=50, temperature=0, timeout=30.0
            )

            response = message.content[0].text.strip()
            document_type = response.split('\n')[0].strip()
            document_type = re.sub(r'[<>:"/\\|?*]', '', document_type)
            return document_type or "その他書類"
        except Exception as e:
            print(f"分類API(テキスト) エラー: {e}")
            self.log_message(f"❌ 分類API(テキスト) エラー: {str(e)}")
            return "PDF文書"

    def extract_text_from_pdf(self, pdf_path, max_pages=2, max_chars=4000):
        try:
            doc = fitz.open(pdf_path)
            text_parts = []
            for i in range(min(len(doc), max_pages)):
                page = doc[i]
                text_parts.append(page.get_text("text"))
                if sum(len(t) for t in text_parts) >= max_chars:
                    break
            doc.close()
            text = "\n".join(text_parts).strip()
            return text[:max_chars]
        except Exception as e:
            print(f"PDFテキスト抽出エラー: {e}")
            return ""

    def classify_with_vision(self, image, prompt_override=None, preset_key=None):
        """Claude Vision APIで文書分類（先頭2ページ対応・簡潔プロンプト）"""
        try:
            # 画像引数をリストに正規化
            images = image if isinstance(image, (list, tuple)) else [image]

            img_blocks = []
            for img in images[:2]:
                buffer = io.BytesIO()
                img.save(buffer, format='PNG')
                image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
                img_blocks.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": image_data
                    }
                })

            label_list = self.build_label_set(preset_key)
            labels = "、".join(label_list)
            base_prompt = (
                f"1ページ目を主とみなして、次の中から最も適切な1つの文書種別のみを日本語で返してください（説明不要・厳密一致）。\n"
                f"候補: {labels}\n"
                f"どれにも当てはまらなければ『その他書類』と返してください。"
            )

            prompt = (prompt_override + "\n\n" + base_prompt) if prompt_override else base_prompt

            message = self._anthropic_call_with_retry(
                ([{"type": "text", "text": prompt}] + img_blocks), max_tokens=50, temperature=0, timeout=30.0
            )
            
            response = message.content[0].text.strip()
            print(f"分類API生レスポンス: {response}")
            
            # レスポンスをクリーンアップ
            document_type = response.split('\n')[0].strip()
            document_type = re.sub(r'^この文書は', '', document_type)
            document_type = re.sub(r'です$', '', document_type)
            document_type = re.sub(r'[<>:"/\\|?*]', '', document_type)
            
            print(f"分類結果: {document_type}")
            return document_type.strip()
            
        except Exception as e:
            print(f"分類API エラー: {e}")
            self.log_message(f"❌ 分類API エラー: {str(e)}")
            return "PDF文書"
    
    def extract_names_and_companies(self, image):
        """宛名（受取人）を抽出。説明を返された場合でも粘り強く再試行して3行形式を得る。"""
        try:
            buffer = io.BytesIO()
            image.save(buffer, format='PNG')
            image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')

            base_prompt = (
                "この日本の文書から『宛名（受取人）』のみを抽出してください。差出人（発行者）は除外してください。\n\n"
                "出力は以下の3行のみ。余計な説明・前置きは絶対に出力しない。\n"
                "法人名：[宛名の法人名または「なし」]\n"
                "姓：[宛名の姓または「なし」]\n"
                "名：[宛名の名または「なし」]\n\n"
                "複数の宛名がある場合は最初の1名（または法人）だけを対象にする。\n"
                "『様』『殿』の直前の名前を優先。肩書きや差出人の事務所名は出力しない。"
            )

            def call_blocks(prompt_text: str):
                return self._anthropic_call_with_retry(
                    [
                        {"type": "text", "text": prompt_text},
                        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": image_data}},
                    ],
                    max_tokens=120, temperature=0, timeout=30.0
                )

            message = call_blocks(base_prompt)
            response = (message.content[0].text or '').strip()
            print(f"名前抽出API生レスポンス: {response}")

            company_name, surname, given_name = self._parse_name_fields(response)

            if not (company_name or surname or given_name):
                strict_prompt = base_prompt + "\n\n絶対条件: 上記3行のみを出力。例や分析文・理由は出力しない。"
                message2 = call_blocks(strict_prompt)
                response2 = (message2.content[0].text or '').strip()
                print(f"名前抽出API再試行レスポンス: {response2}")
                company_name, surname, given_name = self._parse_name_fields(response2)

            return {'surname': surname, 'given_name': given_name, 'company_name': company_name}

        except Exception as e:
            print(f"名前抽出API エラー: {e}")
            self.log_message(f"❌ 名前抽出API エラー: {str(e)}")
            return {'surname': None, 'given_name': None, 'company_name': None}

    def _parse_name_fields(self, response_text: str):
        """『法人名/姓/名』の3行形式を緩くパース。全角・半角コロン対応、'なし'無視、姓名が同一行でも分割。"""
        company_name = None
        surname = None
        given_name = None
        try:
            for raw in (response_text or '').splitlines():
                line = raw.strip()
                if not line:
                    continue
                if line.startswith(('法人名：', '法人名:')):
                    val = re.split('[：:]', line, 1)[1].strip()
                    if val and val.lower() not in ('なし', 'none', 'null'):
                        company_name = val
                elif line.startswith(('姓：', '姓:')):
                    val = re.split('[：:]', line, 1)[1].strip()
                    if val and val.lower() not in ('なし', 'none', 'null'):
                        parts = re.split(r'\s+', val)
                        surname = parts[0] if parts else val
                        if len(parts) >= 2:
                            given_name = parts[1]
                elif line.startswith(('名：', '名:')):
                    val = re.split('[：:]', line, 1)[1].strip()
                    if val and val.lower() not in ('なし', 'none', 'null'):
                        given_name = val
            return company_name, surname, given_name
        except Exception:
            return None, None, None
    
    def extract_property_info(self, image, document_type):
        """登記関連書類から不動産情報を抽出（最適化版）"""
        # 登記関連書類でない場合はスキップ
        registry_keywords = [
            '登記事項証明書', '登記情報', '登記簿', 
            '全部事項証明書', '現在事項証明書', '建物事項証明書',
            '土地登記', '建物登記', '不動産登記'
        ]
        if not any(keyword in document_type for keyword in registry_keywords):
            return None
            
        try:
            buffer = io.BytesIO()
            image.save(buffer, format='PNG')
            image_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
            
            prompt = """この登記簿から不動産情報を抽出してください。

【出力形式】
種別：[土地/建物/区分建物]
所在：[所在地または「なし」]
地番等：[地番・家屋番号・部屋番号等または「なし」]

【抽出ルール】
- 表題部の「不動産の表示」欄から抽出
- 土地：所在地と地番を分けて記載
- 建物：所在地と家屋番号を分けて記載
- 区分建物：所在は「なし」、地番等に建物名+部屋番号のみ

【例】
種別：土地
所在：福岡市中央区清川一丁目
地番等：11号16番地

種別：区分建物
所在：なし
地番等：パークマンション801号"""

            def _call3(model):
                return self.claude_client.messages.create(
                    model=model,
                    max_tokens=150,
                    temperature=0,
                    messages=[{
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": image_data
                                }
                            }
                        ]
                    }],
                    timeout=30.0
                )
            model_primary = self.get_model()
            try:
                message = _call3(model_primary)
            except Exception as e:
                fb = 'claude-3-5-sonnet-20241022'
                self.log_message(f"⚠️ モデル '{model_primary}' 失敗。フォールバック {fb} を試行: {e}")
                message = _call3(fb)
            
            response = message.content[0].text.strip()
            print(f"不動産情報API生レスポンス: {response}")
            
            # レスポンスを解析
            property_info = {
                'type': None,
                'location': None,
                'address_number': None
            }
            
            for line in response.split('\n'):
                line = line.strip()
                if line.startswith('種別：'):
                    property_info['type'] = line.replace('種別：', '').strip()
                elif line.startswith('所在：'):
                    location = line.replace('所在：', '').strip()
                    if location.lower() not in ['なし', 'none', '']:
                        property_info['location'] = location
                elif line.startswith('地番等：'):
                    address_number = line.replace('地番等：', '').strip()
                    if address_number.lower() not in ['なし', 'none', '']:
                        property_info['address_number'] = address_number
            
            return property_info
            
        except Exception as e:
            print(f"不動産情報API エラー: {e}")
            self.log_message(f"❌ 不動産情報API エラー: {str(e)}")
            return None
    
    
    # システムトレイ・その他のメソッド（既存と同じ）
    def on_minimize(self, event):
        """ウィンドウ最小化時の処理（重複初期化防止版）"""
        try:
            if event.widget == self.window and self.minimize_to_tray.get():
                # 既に格納中なら何もしない
                if getattr(self, 'is_minimized_to_tray', False):
                    print("既に最小化済み - 処理をスキップ")
                    return
                
                # トレイアイコンの存在確認（デタッチ状態も考慮）
                if not self.tray_icon:
                    print("トレイアイコンが存在しないため初期化します")
                    self.init_tray()
                    if self.tray_icon:
                        try:
                            self.tray_icon.run_detached()
                            self._tray_detached = True
                            print("トレイアイコンをデタッチ実行")
                        except Exception as e:
                            print(f"デタッチ実行エラー: {e}")
                elif not self._tray_detached:
                    print("トレイアイコンは存在するがデタッチされていません")
                    try:
                        self.tray_icon.run_detached()
                        self._tray_detached = True
                        print("既存トレイアイコンをデタッチ実行")
                    except Exception as e:
                        print(f"デタッチ実行エラー: {e}")
                else:
                    print("既存のトレイアイコン（デタッチ済み）を使用")
                
                # ウィンドウを隠す
                self.hide_window()
        except Exception as e:
            print(f"最小化処理エラー: {e}")
    
    def hide_window(self):
        """ウィンドウをトレイに隠す"""
        try:
            self.window.withdraw()
            self.is_minimized_to_tray = True
            # トレイは常時稼働（run_detached）。ここでは明示操作しない。
        except Exception:
            pass
    
    def show_window(self):
        """ウィンドウを表示"""
        try:
            self.window.deiconify()
            self.window.lift()
            try:
                self.window.focus_force()
            except Exception:
                pass
            self.is_minimized_to_tray = False
            # トレイは常時表示（ここでは何もしない）
        except Exception:
            pass
    
    def on_window_close(self):
        """ウィンドウクローズ時の処理"""
        self.quit_application()
    
    def quit_application(self):
        """アプリケーション終了（トレイメニューからも安全に実行）"""
        def _do_quit():
            try:
                if self.is_watching:
                    self.stop_watching()
            except Exception:
                pass
            try:
                if self.tray_icon:
                    try:
                        self.tray_icon.stop()
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                # mainloopを抜けてから破棄
                self.window.quit()
            except Exception:
                pass
            try:
                self.window.destroy()
            except Exception:
                pass
            # 念のためプロセスを終了
            try:
                os._exit(0)
            except Exception:
                sys.exit(0)

        # Tk操作はメインスレッドにディスパッチ
        try:
            self.window.after(0, _do_quit)
        except Exception:
            _do_quit()

    def safe_notify(self, title: str, message: str):
        """トレイを使わず控えめなトーストで通知"""
        try:
            t = title if len(title) <= 60 else (title[:57] + '...')
            m = message if len(message) <= 200 else (message[:197] + '...')
            self.show_toast(t, m)
        except Exception as e:
            self.log_message(f"⚠️ 通知失敗: {e}")
    
    def update_startup_setting(self):
        """Windowsスタートアップ設定を更新"""
        try:
            app_name = "PDF Auto Watcher Advanced"
            app_path = os.path.abspath(__file__)
            
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            
            if self.auto_startup.get():
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{app_path}"')
                self.log_message("✅ スタートアップに登録しました")
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                    self.log_message("❌ スタートアップから削除しました")
                except FileNotFoundError:
                    pass
            
            winreg.CloseKey(key)
            
        except Exception as e:
            self.log_message(f"スタートアップ設定エラー: {e}")
    
    def show_api_setup_dialog(self):
        """初回起動時のAPI設定ダイアログ"""
        if self.claude_client:  # APIが既に設定されている場合はスキップ
            return
            
        result = messagebox.askyesno(
            "API設定が必要です",
            "PDF自動監視ツールを使用するにはClaude APIキーの設定が必要です。\n\n今すぐ設定しますか？",
            icon="question"
        )
        
        if result:
            self.setup_api()
        else:
            self.log_message("❌ APIキーが未設定のため、PDF処理機能は利用できません")
    
    def setup_api(self):
        """API設定ダイアログ（モデルはプルダウン：Claude 4 Sonnet 固定）"""
        dialog = tk.Toplevel(self.window)
        dialog.title("API設定")
        dialog.geometry("700x420")
        try:
            dialog.minsize(640, 360)
        except Exception:
            pass
        dialog.configure(bg="white")

        current_key = os.environ.get('ANTHROPIC_API_KEY', '') or (self._read_api_key_from_appdata() or '')
        status_text = "設定済み" if current_key else "未設定"

        tk.Label(dialog, text=f"現在の状況: {status_text}", bg="white").pack(pady=10)

        tk.Label(dialog, text="Claude API Key:", bg="white").pack()
        key_entry = tk.Entry(dialog, width=54, show="*")
        key_entry.pack(pady=5)
        if current_key:
            key_entry.insert(0, current_key)

        tk.Label(dialog, text="モデル（固定推奨）:", bg="white").pack(pady=(10, 4))
        model_combo = ttk.Combobox(dialog, state='readonly', width=44)
        model_combo['values'] = ["Claude 4 Sonnet"]
        model_combo.current(0)
        model_combo.pack()

        def save_api_key():
            new_key = key_entry.get().strip()
            # プルダウンは1択：Claude 4 Sonnet（API IDに変換して保存）
            new_model = 'claude-sonnet-4-20250514'
            if new_key:
                try:
                    test_client = anthropic.Anthropic(api_key=new_key)
                    # 簡易検証としてクライアントを生成（実コールは不要）
                    self._write_api_key_to_appdata(new_key)
                    os.environ['ANTHROPIC_API_KEY'] = new_key  # このセッションでも利用
                    self.claude_client = test_client
                    self.config['model'] = new_model
                    self.model_name = new_model
                    self.save_config()
                    messagebox.showinfo("成功", "APIキーを保存しました（AppData）／モデル: Claude 4 Sonnet")
                    self.log_message("✅ Claude APIキーをAppDataに保存し、モデルを設定しました（Claude 4 Sonnet）")
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("エラー", f"APIキーの検証に失敗: {e}")
            else:
                messagebox.showwarning("警告", "APIキーを入力してください")

        ttk.Button(dialog, text="保存", command=save_api_key).pack(pady=20)
    
    def run(self):
        """アプリケーション実行"""
        self.window.mainloop()

if __name__ == "__main__":
    # シングルトン確保（多層ロック：ファイルロック → TCP → Windows Mutex）
    running = False
    mutex_handle = None
    _lock_fd = None
    __lock_socket = None
    lock_path = None
    try:
        # 0) ファイルロック（最優先・プロセス間で最も確実）
        try:
            base_dir = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
            lock_dir = os.path.join(base_dir, 'AutoPDFWatcherAdvanced')
            os.makedirs(lock_dir, exist_ok=True)
            lock_path = os.path.join(lock_dir, 'app.lock')
            # O_EXCLで原子的に作成（既にあれば起動中と判断）
            try:
                _lock_fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                os.write(_lock_fd, str(os.getpid()).encode('utf-8'))
            except FileExistsError:
                # ステールの可能性があるためIPCへ接続を試す
                try:
                    c = socket.create_connection(("127.0.0.1", 57321), timeout=0.4)
                    c.sendall(b"SHOW\n")
                    c.close()
                    running = True
                except Exception:
                    # 接続できない = ステールロック。ロックファイルを削除して再取得
                    try:
                        os.unlink(lock_path)
                    except Exception:
                        pass
                    _lock_fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                    os.write(_lock_fd, str(os.getpid()).encode('utf-8'))
        except Exception:
            # 失敗しても他のロック方式でカバーする
            pass

        # 1) クロスプラットフォームでTCPロックを先に試みる（高速・確実）
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.bind(("127.0.0.1", 57321))
            s.listen(5)
            __lock_socket = s  # keep for IPC
        except OSError:
            running = True

        # 2) WindowsではNamed Mutexでも二重防止（同時起動レースの更なる保険）
        if platform.system() == 'Windows' and not running:
            kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            mutex_handle = kernel32.CreateMutexW(None, True, "Global\\PDF_Auto_Watcher_Advanced_Mutex")
            last_err = ctypes.get_last_error()
            if last_err == 183 or not mutex_handle:
                running = True
        if running:
            # 既存インスタンスへ『表示』IPCを送って終了
            try:
                c = socket.create_connection(("127.0.0.1", 57321), timeout=0.5)
                c.sendall(b"SHOW\n")
                c.close()
            except Exception:
                pass
            sys.exit(0)
        app = AutoPDFWatcherAdvanced()
        # IPCサーバ（既存インスタンスを前面表示）
        try:
            if __lock_socket:
                def _ipc_server():
                    while True:
                        try:
                            conn, _ = __lock_socket.accept()
                            data = conn.recv(64)
                            if data and data.strip().upper().startswith(b"SHOW"):
                                try:
                                    app.window.after(0, app.show_window)
                                except Exception:
                                    pass
                            conn.close()
                        except Exception:
                            time.sleep(0.1)
                threading.Thread(target=_ipc_server, daemon=True).start()
        except Exception:
            pass
        app.run()
    finally:
        try:
            if platform.system() == 'Windows' and mutex_handle:
                ctypes.windll.kernel32.CloseHandle(mutex_handle)
        except Exception:
            pass
        # ファイルロックはプロセス終了で自動解放されるが、明示クリーンアップ
        try:
            if _lock_fd is not None:
                os.close(_lock_fd)
                if lock_path and os.path.exists(lock_path):
                    os.unlink(lock_path)
        except Exception:
            pass
