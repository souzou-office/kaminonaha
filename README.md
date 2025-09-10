紙の名は。 

PDFを監視して、内容からわかりやすい日本語ファイル名に自動リネームするWindows用ツール。
複数フォルダ同時監視・フォルダごとの簡易ルール・トレイ常駐に対応。

例）「scan_2025-09-10_001.pdf」→「請求書_2025-09-10_山田商事.pdf」

主な機能

指定フォルダを常時監視（複数可）

PDFの1ページ目などを解析し、種別（請求書/契約書/議事録…）を推定して自動命名

日付・氏名/会社名の付与オプション

フォルダごとに自然文の指示で命名方針をカスタム

最小化でシステムトレイ常駐（開始/停止メニュー）

設定はAppDataに保存（PC再起動後も保持）

画面構成（ざっくり）

上部：開始/停止ボタン・ステータス

中央：監視フォルダ一覧（追加/設定/削除）

下部：処理ログ（保持期間は分単位で設定可）

動作環境

Windows 10/11

Python 3.10+（配布版EXEでも可）

ネット接続（AI命名にClaude APIを使う場合）

依存ライブラリ（主要）

watchdog, PyMuPDF (fitz), pystray, Pillow, tkinter, anthropic ほか。

auto_pdf_watcher_advanced_distr…

セットアップ（ソースから）
# 1) 仮想環境（任意）
python -m venv .venv
.venv\Scripts\activate

# 2) 依存関係
pip install watchdog pymupdf pystray pillow anthropic

# 3) 起動
python auto_pdf_watcher_advanced_distribution.py

初期設定

アプリを起動 → 「フォルダ追加」で監視ターゲットを登録

（任意）フォルダごとの「フォルダ設定」を開き、

「ファイル名に日付を付ける」

「名前（会社名/氏名）を付ける」

「AI命名の指示（自然文）」…例：

「1ページ目の見出しを短く使う。名詞句で8〜16文字。会社名は付けない。」

「開始」を押すと監視スタート。最小化するとトレイへ。

Claude API（任意・AI命名を使う場合）

環境変数 ANTHROPIC_API_KEY にAPIキーを設定
もしくはアプリ内の「API設定」から保存できます。

auto_pdf_watcher_advanced_distr…

設定ファイルの保存場所

AppData\AutoPDFWatcherAdvanced\config.json に自動保存／読み込み。
（エクスポート/インポートも対応）

auto_pdf_watcher_advanced_distr…

自動起動（任意）

設定の「Windows起動時に自動で開始」をONで、起動時に常駐します（レジストリ使用）。

auto_pdf_watcher_advanced_distr…

ビルド（配布用EXEを作る）

PyInstaller例：

pip install pyinstaller
pyinstaller -F -w ^
  --name Kaminonaha ^
  --icon kaminonaha_latest.ico ^
  --add-data "kaminonaha_latest.ico;." ^
  auto_pdf_watcher_advanced_distribution.py


タスクバー/トレイのアイコンは ICO埋め込み＋同梱 を推奨。ICOは16/24/32/48/64px程度を含むものが安定。

auto_pdf_watcher_advanced_distr…

トラブルシュート（よくある質問）

アイコンが出ない/変になる

ICOの場所解決は「PyInstaller同梱 → 環境変数 KAMINONAHA_ICON → 実行ファイルと同じフォルダ → AppData」の順で探索します。ICOをそのいずれかに配置し、サイズを複数含めてください。

auto_pdf_watcher_advanced_distr…

AI命名が動かない

ANTHROPIC_API_KEYが未設定、または通信不可の可能性。ログ欄の警告を確認。

auto_pdf_watcher_advanced_distr…

PDFの分類が意図と違う

フォルダ設定の「AI命名の指示」を短く具体的に。
例：

「文書の種類名だけ（請求書/契約書など）。会社名は付けない。名詞句で10〜14文字。」
