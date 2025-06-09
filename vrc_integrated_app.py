import tkinter as tk
from tkinter import ttk, messagebox
import gspread
from google.oauth2.service_account import Credentials
import vrchatapi
from vrchatapi.api import authentication_api
from vrchatapi.api import avatars_api
from vrchatapi.api import worlds_api
from vrchatapi.exceptions import UnauthorizedException, ApiException
from datetime import datetime
import time
import threading
from pythonosc import dispatcher
from pythonosc import osc_server
import requests
import os
import http.cookiejar
import base64
import json
import copy

# Pillowライブラリのインポート
from PIL import Image, ImageTk

# --- 設定 ---
# Google Sheets 設定
# ダウンロードしたサービスアカウントのJSONキーファイルへのパス
SERVICE_ACCOUNT_KEY_FILE = 'your-service-account-key.json' # <-- ★ここを書き換えてください★
# スプレッドシートの名前 (履歴保存用とGUI読み込み用を兼ねる)
SPREADSHEET_NAME = 'VRCAPI' # <-- ★使用するスプレッドシート名に書き換えてください★
# ワークシートの名前 (履歴を記録・読み込みするシート)
WORKSHEET_NAME = 'API' # <-- ★使用するワークsheet名に書き換えてください★
# 履歴を記録する列のヘッダー名 (スプレッドシートの1行目にこれらのヘッダーが必要です)
HISTORY_HEADERS = ['Type', 'ID', 'Name', 'author', 'Timestamp', 'Image URL']

# VRChat API エンドポイント (認証用)
VRC_API_AUTH_URL = "https://api.vrchat.cloud/api/1/auth/user"
VRC_API_2FA_VERIFY_URL = "https://api.vrchat.cloud/api/1/auth/user/verify"
VRC_API_2FA_EMAIL_VERIFY_URL = "https://api.vrchat.cloud/api/1/auth/user/verify2faemail"

# VRChat アカウント設定 - 実行時に入力します
VRC_USERNAME = None # 実行時に入力
VRC_PASSWORD = None # 実行時に入力

# OSC 設定
OSC_LISTEN_IP = "127.0.0.1" # OSCメッセージを待ち受けるIPアドレス (通常はローカルホスト)
OSC_LISTEN_PORT = 9001 # VRChatが送信するデフォルトポート 9001

# VRChatがアバター/ワールド変更時に送信するOSCアドレス (コミュニティ情報に基づく例)
OSC_AVATAR_CHANGE_ADDRESS = "/avatar/change"
OSC_WORLD_CHANGE_ADDRESS = "/world/change"

# API呼び出し間の遅延 (秒) - レートリミット回避のため
API_CALL_DELAY = 1 # GUI操作も考慮して少し短く設定

# クッキー保存ファイル名
AUTH_COOKIE_FILE = 'vrc_auth_cookie.lwp' # LWPCookieJar形式で保存

# 画像キャッシュディレクトリ
IMAGE_CACHE_DIR = 'image_cache' # 画像キャッシュを保存するディレクトリ名

# --- グローバル変数 ---
current_user = None
gc = None
worksheet = None
recorded_ids = set()
history_data_list = []
cookie_jar = http.cookiejar.LWPCookieJar(AUTH_COOKIE_FILE)
osc_server_instance = None
osc_server_thread = None
is_osc_server_running = False
root = None
username_entry = None
password_entry = None
twofactor_entry = None
login_button = None
auth_status_label = None
login_frame = None
history_frame = None
history_treeview = None
image_display_label = None # 画像表示用ラベル
current_photo_image = None # PhotoImageオブジェクトを保持するための変数

# --- 関数: Google Sheets 認証とワークsheet準備 ---
def setup_google_sheets():
    """Google Sheets APIの認証を行い、ワークsheetを準備する"""
    global gc, worksheet, recorded_ids
    print("Google スプレッドシートに接続中...")
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive' # スプレッドシートを開くために必要
        ]
        # サービスアカウント認証情報のロード
        if not os.path.exists(SERVICE_ACCOUNT_KEY_FILE):
             messagebox.showerror("設定エラー", f"サービスアカウントキーファイル '{SERVICE_ACCOUNT_KEY_FILE}' が見つかりません。")
             print(f"設定エラー: サービスアカウントキーファイル '{SERVICE_ACCOUNT_KEY_FILE}' が見つかりません。")
             return False

        credentials = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_KEY_FILE,
            scopes=scopes
        )
        # gspread クライアントの認証
        gc = gspread.authorize(credentials)
        print("Google スプレッドシート認証成功。")

        # スプレッドシートとワークsheetを開く
        try:
            spreadsheet = gc.open(SPREADSHEET_NAME)
            print(f"スプレッドシート '{SPREADSHEET_NAME}' を開きました。")
        except gspread.SpreadsheetNotFound:
            messagebox.showerror("設定エラー", f"スプレッドシート '{SPREADSHEET_NAME}' が見つかりません。存在することを確認してください。")
            print(f"設定エラー: スプレッドシート '{SPREADSHEET_NAME}' が見つかりません。存在することを確認してください。")
            gc = None # 無効なクライアントをクリア
            return False

        try:
            worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
            print(f"ワークsheet '{WORKSHEET_NAME}' を開きました。")
        except gspread.WorksheetNotFound:
            # ワークsheetが存在しない場合は新規作成
            print(f"ワークsheet '{WORKSHEET_NAME}' が見つかりませんでした。新規作成します。")
            worksheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows="100", cols="10")
            # ヘッダー行を書き込む
            worksheet.append_row(HISTORY_HEADERS)
            print(f"ワークsheet '{WORKSHEET_NAME}' を作成し、ヘッダーを書き込みました。")

        # 既存の履歴IDを読み込み、セットに格納 (重複防止用)
        print("既存の履歴を読み込み中...")
        all_records = worksheet.get_all_values()
        if len(all_records) > 1: # ヘッダー行以外にデータがある場合
             # ヘッダー行のID列のインデックスを取得
             try:
                 # ヘッダー行から各列のインデックスを取得
                 header_indices = {header: all_records[0].index(header) for header in HISTORY_HEADERS if header in all_records[0]}
                 id_col_index = header_indices.get('ID')

                 if id_col_index is not None:
                     for row in all_records[1:]: # データ行のみ処理
                         if len(row) > id_col_index and row[id_col_index]:
                              recorded_ids.add(row[id_col_index])
                     print(f"既存の履歴から {len(recorded_ids)} 件のIDを読み込みました。")
                 else:
                      print("警告: スプレッドシートに 'ID' ヘッダーが見つかりません。重複チェックが正しく機能しない可能性があります。")

             except ValueError as e:
                  print(f"警告: スプレッドシートのヘッダー解析エラー: {e}。重複チェックが正しく機能しない可能性があります。")
             except Exception as e:
                  print(f"警告: 既存履歴の読み込み中にエラーが発生しました: {e}")

        return True
    except Exception as e:
        messagebox.showerror("認証エラー", f"Google スプレッドシート認証または設定エラー: {e}")
        print(f"Google スプレッドシート認証または設定エラー: {e}")
        return False

# --- 関数: 認証クッキーをファイルに保存 ---
def save_auth_cookie():
    """グローバルな cookie_jar オブジェクトをファイルに保存する"""
    global cookie_jar
    print(f"DEBUG: save_auth_cookie() called. Cookie jar has {len(cookie_jar)} cookies.")
    if cookie_jar:
        try:
            # ディレクトリが存在しない場合は作成
            cookie_dir = os.path.dirname(AUTH_COOKIE_FILE)
            if cookie_dir and not os.path.exists(cookie_dir):
                os.makedirs(cookie_dir)
            cookie_jar.save(AUTH_COOKIE_FILE, ignore_discard=True, ignore_expires=True)
            print(f"認証クッキーを '{AUTH_COOKIE_FILE}' に保存しました。")
        except Exception as e:
            print(f"認証クッキーの保存中にエラーが発生しました: {e}")
    else:
        print("DEBUG: cookie_jar is None. Cannot save.")

# --- 関数: 認証クッキーをファイルから読み込み ---
def load_auth_cookie():
    """ファイルから認証クッキーを読み込み、グローバルな cookie_jar に設定する"""
    global cookie_jar
    print(f"DEBUG: load_auth_cookie() called. Looking for '{AUTH_COOKIE_FILE}'.")
    try:
        if os.path.exists(AUTH_COOKIE_FILE):
            cookie_jar.load(AUTH_COOKIE_FILE, ignore_discard=True, ignore_expires=True)
            print(f"認証クッキーを '{AUTH_COOKIE_FILE}' から読み込みました。")
            print(f"DEBUG: Loaded cookie jar has {len(cookie_jar)} cookies.")
            return True # 読み込み成功
        else:
            print(f"認証クッキーファイル '{AUTH_COOKIE_FILE}' が見つかりませんでした。")
            return False # ファイルなし
    except Exception as e:
        print(f"認証クッキーの読み込み中にエラーが発生しました: {e}")
        return False # 読み込み失敗


# --- 関数: VRChat API 認証 (requests を使用) ---
def authenticate_vrchat_with_requests(username, password, two_factor_code=None):
    """requests ライブラリを使用して VRChat API の認証を行う"""
    global cookie_jar, current_user
    print("requests を使用して VRChat API にログイン中...")
    root.after(0, update_auth_status_label, "ログイン中...")

    # 認証用のrequestsセッションを作成
    auth_session = requests.Session()
    auth_session.cookies = cookie_jar # 既存のクッキーがあればロード
    auth_session.headers.update({'User-Agent': f'MyVRChatIntegratedApp/1.0 ({username})'})

    # Basic認証ヘッダーを作成
    auth_string = f"{username}:{password}"
    base64_auth_string = base64.b64encode(auth_string.encode()).decode()
    headers = {
        "Authorization": f"Basic {base64_auth_string}"
    }

    try:
        # 認証リクエストを送信
        response = auth_session.get(VRC_API_AUTH_URL, headers=headers)
        response.raise_for_status() # HTTPエラーがあれば例外を発生させる

        user_data = response.json()

        if response.status_code == 200:
            # 認証成功
            print("requests 認証成功。")
            # ユーザー情報を取得 (簡易的なオブジェクトとして保持)
            current_user = type('CurrentUser', (object,), user_data)
            print(f"DEBUG: User data keys: {user_data.keys()}")
            display_name = getattr(current_user, 'displayName', '不明なユーザー名')
            print(f"VRChatログイン成功: {display_name}")

            # 認証成功後、auth_session のクッキーをグローバルな cookie_jar にコピーして保存
            cookie_jar.clear() # 既存のクッキーをクリア
            for cookie in auth_session.cookies:
                 cookie_jar.set_cookie(cookie)
            save_auth_cookie()

            root.after(0, update_auth_status_label, f"ログイン済み: {display_name}", "green") # GUI更新
            root.after(0, hide_manual_login_ui) # 手動ログインUIを隠す
            # 認証成功後、OSCサーバーを自動起動
            start_osc_server_threaded()
            return True

        elif response.status_code == 401:
            # 認証失敗または2FAが必要
            error_data = response.json()
            if error_data.get('requiresTwoFactorAuth'):
                print("ログイン情報が確認されましたが、二段階認証が必要です。")
                if two_factor_code:
                    print("二段階認証コードを検証中...")
                    root.after(0, update_auth_status_label, "2FA検証中...")
                    # 2FA検証リクエストを送信 (認証に使ったセッションを再利用)
                    verify_url = VRC_API_2FA_VERIFY_URL # TOTPの場合
                    if error_data.get('twoFactorAuthType') == 'email':
                         verify_url = VRC_API_2FA_EMAIL_VERIFY_URL # Emailの場合

                    verify_payload = {'code': two_factor_code}
                    verify_response = auth_session.post(verify_url, json=verify_payload)

                    try:
                        verify_response.raise_for_status() # 2FA検証のHTTPエラーをチェック
                        print("二段階認証検証成功。")

                        # 2FA検証成功後、再度ユーザー情報を取得してログイン完了とする
                        final_response = auth_session.get(VRC_API_AUTH_URL)
                        final_response.raise_for_status()
                        user_data_final = final_response.json()
                        current_user = type('CurrentUser', (object,), user_data_final) # 簡易的なオブジェクトとして保持
                        print(f"DEBUG: Final user data keys: {user_data_final.keys()}")
                        display_name_final = getattr(current_user, 'displayName', '不明なユーザー名')
                        print(f"VRChatログイン成功 (2FA完了): {display_name_final}")

                        # 認証成功後、auth_session のクッキーをグローバルな cookie_jar にコピーして保存
                        cookie_jar.clear()
                        for cookie in auth_session.cookies:
                             cookie_jar.set_cookie(cookie)
                        save_auth_cookie()

                        root.after(0, update_auth_status_label, f"ログイン済み: {display_name_final}", "green") # GUI更新
                        root.after(0, hide_manual_login_ui) # 手動ログインUIを隠す
                        # 認証成功後、OSCサーバーを自動起動
                        start_osc_server_threaded()
                        return True

                    except requests.exceptions.RequestException as verify_e:
                         error_details = verify_e.response.json() if hasattr(verify_e.response, 'json') else str(verify_e)
                         messagebox.showerror("認証エラー", f"二段階認証検証に失敗しました: {error_details}")
                         print(f"二段階認証検証に失敗しました: {error_details}")
                         current_user = None
                         root.after(0, update_auth_status_label, "認証失敗", "red")
                         root.after(0, show_manual_login_ui)
                         return False

                    except Exception as e:
                         messagebox.showerror("認証エラー", f"二段階認証検証中に予期しないエラーが発生しました: {e}")
                         print(f"二段階認証検証中に予期しないエラーが発生しました: {e}")
                         current_user = None
                         root.after(0, update_auth_status_label, "認証失敗", "red")
                         root.after(0, show_manual_login_ui)
                         return False

                else:
                    # 2FAが必要だがコードが提供されていない場合
                    messagebox.showwarning("認証情報", "二段階認証コードを入力してください。")
                    print("二段階認証コードを入力してください。")
                    current_user = None
                    root.after(0, update_auth_status_label, "二段階認証コード入力待ち", "orange")
                    root.after(0, show_manual_login_ui, requires_2fa=True)
                    return False

            else:
                # その他の認証失敗
                error_data = response.json() if hasattr(response, 'json') else {'error': {'message': f'HTTP Status {response.status_code}'}}
                error_message = error_data.get('error', {}).get('message', '不明なエラー')
                messagebox.showerror("認証エラー", f"VRChatログインに失敗しました: {error_message}")
                print(f"VRChatログインに失敗しました: {error_message}")
                current_user = None
                root.after(0, update_auth_status_label, "認証失敗", "red")
                root.after(0, show_manual_login_ui)
                return False

        else:
            # その他のHTTPステータスコード
            messagebox.showerror("APIエラー", f"VRChatログイン時に予期しないHTTPエラーが発生しました: Status {response.status_code}")
            print(f"VRChatログイン時に予期しないHTTPエラーが発生しました: Status {response.status_code}")
            current_user = None
            root.after(0, update_auth_status_label, "認証失敗", "red")
            root.after(0, show_manual_login_ui)
            return False

    except requests.exceptions.RequestException as e:
        # ネットワークエラーなど
        messagebox.showerror("ネットワークエラー", f"VRChat API への接続中にエラーが発生しました: {e}")
        print(f"ネットワークエラー: VRChat API への接続中にエラーが発生しました: {e}")
        current_user = None
        root.after(0, update_auth_status_label, "接続エラー", "red")
        root.after(0, show_manual_login_ui)
        return False

    except Exception as e:
        # 予期しないその他のエラー
        messagebox.showerror("エラー", f"VRChatログイン中に予期しないエラーが発生しました: {e}")
        print(f"予期しないエラー: VRChatログイン中に予期しないエラーが発生しました: {e}")
        current_user = None
        root.after(0, update_auth_status_label, "認証エラー", "red")
        root.after(0, show_manual_login_ui)
        return False
    finally:
        # 認証に使ったセッションを閉じる
        auth_session.close()
        print("DEBUG: Authentication session closed.")


# --- 関数: クッキーによる自動ログインを試行 (requests を使用) ---
def attempt_auto_login_with_requests():
    """保存されたクッキーを使用して自動ログインを試みる (requests を使用)"""
    global cookie_jar, current_user
    print("クッキーによる自動ログインを試行中 (requests を使用)...")
    root.after(0, update_auth_status_label, "クッキーで自動ログイン試行中...")

    # 保存されたクッキーを読み込み
    cookie_loaded = load_auth_cookie()

    auto_login_session = None
    try:
        if cookie_loaded and len(cookie_jar) > 0:
            print("クッキーが読み込まれました。自動ログインを試行します。")
            auto_login_session = requests.Session()
            auto_login_session.cookies.update(cookie_jar)
            auto_login_session.headers.update({'User-Agent': 'MyVRChatIntegratedApp/1.0 (AutoLogin)'})

            try:
                # 認証状態を確認するためにユーザー情報を取得
                response = auto_login_session.get(VRC_API_AUTH_URL)
                response.raise_for_status() # HTTPエラーがあれば例外を発生させる

                user_data = response.json()

                if response.status_code == 200:
                    # 認証成功
                    print("クッキー認証成功。")
                    current_user = type('CurrentUser', (object,), user_data)
                    print(f"DEBUG: Auto-login user data keys: {user_data.keys()}")
                    display_name = getattr(current_user, 'displayName', '不明なユーザー名')
                    print(f"VRChatログイン成功: {display_name}")

                    root.after(0, update_auth_status_label, f"ログイン済み: {display_name}", "green")
                    root.after(0, hide_manual_login_ui)
                    start_osc_server_threaded()

                elif response.status_code == 401:
                    # クッキーが無効または期限切れ
                    print("クッキー認証失敗: クッキーが無効または期限切れです。")
                    messagebox.showinfo("認証失敗", "自動ログインに失敗しました。再度ログインしてください。")
                    current_user = None
                    root.after(0, update_auth_status_label, "未ログイン", "red")
                    root.after(0, show_manual_login_ui)

                else:
                    # その他のHTTPステータスコード
                    messagebox.showerror("APIエラー", f"クッキー認証時に予期しないHTTPエラーが発生しました: Status {response.status_code}")
                    print(f"APIエラー: クッキー認証時に予期しないHTTPエラーが発生しました: Status {response.status_code}")
                    current_user = None
                    root.after(0, update_auth_status_label, "認証失敗", "red")
                    root.after(0, show_manual_login_ui)

            except requests.exceptions.RequestException as e:
                # ネットワークエラーなど
                messagebox.showerror("ネットワークエラー", f"VRChat API への接続中にエラーが発生しました: {e}")
                print(f"ネットワークエラー: VRChat API への接続中にエラーが発生しました: {e}")
                current_user = None
                root.after(0, update_auth_status_label, "接続エラー", "red")
                root.after(0, show_manual_login_ui)

            except Exception as e:
                # 予期しないその他のエラー
                messagebox.showerror("エラー", f"クッキー認証中に予期しないエラーが発生しました: {e}")
                print(f"予期しないエラー: クッキー認証中に予期しないエラーが発生しました: {e}")
                current_user = None
                root.after(0, update_auth_status_label, "認証エラー", "red")
                root.after(0, show_manual_login_ui)

        else:
            print("保存されたクッキーが見つからなかったか、クッキーが空でした。手動ログインが必要です。")
            current_user = None
            root.after(0, update_auth_status_label, "未ログイン", "red")
            root.after(0, show_manual_login_ui)
    finally:
        if auto_login_session:
             auto_login_session.close()
             print("DEBUG: Auto-login session closed.")


# --- 関数: アバター変更OSCメッセージハンドラー ---
def avatar_change_handler(address, *args):
    """アバター変更OSCメッセージを受信した際に呼び出される"""
    print(f"OSCメッセージ受信: {address}, args: {args}")
    if args and isinstance(args[0], str):
        avatar_id = args[0]
        print(f"アバター変更を検出: ID = {avatar_id}")
        if cookie_jar and len(cookie_jar) > 0:
             record_thread = threading.Thread(target=record_item_history, args=('Avatar', avatar_id))
             record_thread.start()
        else:
            print("エラー: 認証クッキーがありません。アバター詳細を取得できません。")


# --- 関数: ワールド変更OSCメッセージハンドラー ---
def world_change_handler(address, *args):
    """ワールド変更OSCメッセージを受信した際に呼び出される"""
    print(f"OSCメッセージ受信: {address}, args: {args}")
    if args and isinstance(args[0], str):
        world_id = args[0]
        print(f"ワールド変更を検出: ID = {world_id}")
        if cookie_jar and len(cookie_jar) > 0:
             record_thread = threading.Thread(target=record_item_history, args=('World', world_id))
             record_thread.start()
        else:
            print("エラー: 認証クッキーがありません。ワールド詳細を取得できません。")


# --- 関数: アイテム（アバターまたはワールド）の履歴を記録 ---
def record_item_history(item_type, item_id):
    """アイテムの詳細情報を取得し、スプレッドシートに記録する"""
    print(f"DEBUG: Entering record_item_history for {item_type} ID: {item_id}")

    item_name = f"詳細取得中 ({item_id})"
    item_author = "取得中..." # 作者名初期値
    image_url = ""
    api_call_successful = False

    try:
        sheets_available = (gc is not None and worksheet is not None)
        if not sheets_available:
             print("エラー: Google スプレッドシートが設定されていません。記録はスキップされます。")

        if current_user is None or not cookie_jar or len(cookie_jar) == 0:
            print("エラー: VRChat APIにログインしていません、または認証クッキーがありません。詳細を取得できません。")
            item_name = f"未ログイン/クッキーなし ({item_id})"
            item_author = "未ログイン" # 作者名も更新
        else:
            print(f"{item_type} ID '{item_id}' の詳細情報を取得中...")

            try:
                # vrchatapi.ApiClient を正しく初期化し、クッキーを設定する
                config = vrchatapi.Configuration()
                api_client = vrchatapi.ApiClient(config)
                api_client.user_agent = f"MyVRChatIntegratedApp/1.0 (Record_{item_type})"

                # cookie_jar からクッキーをヘッダー文字列に変換
                cookie_string = "; ".join([f"{cookie.name}={cookie.value}" for cookie in cookie_jar])
                if cookie_string:
                    api_client.set_default_header("Cookie", cookie_string)
                else:
                    raise Exception("Auth cookie not found for recording history.")


                if item_type == 'Avatar':
                    avatars_api_instance = avatars_api.AvatarsApi(api_client)
                    print(f"DEBUG: Calling get_avatar for ID: {item_id}")
                    avatar_info = avatars_api_instance.get_avatar(avatar_id=item_id)
                    item_name = avatar_info.name
                    item_author =getattr(avatar_info, 'authorName', 'a不明')
                    if not item_author.strip():
                        item_author = "名無し"
                    image_url = avatar_info.image_url or avatar_info.thumbnail_image_url
                    api_call_successful = True
                    print(f"DEBUG: Successfully got avatar info: {item_name}, Author: {item_author}")

                elif item_type == 'World':
                    worlds_api_instance = worlds_api.WorldsApi(api_client)
                    print(f"DEBUG: Calling get_world for ID: {item_id}")
                    world_info = worlds_api_instance.get_world(world_id=item_id)
                    item_name = world_info.name
                    item_author = getattr(world_info, 'authorName', 'w不明')
                    if not item_author.strip():
                        item_author = "名無し"
                    image_url = world_info.image_url or world_info.thumbnail_image_url
                    api_call_successful = True
                    print(f"DEBUG: Successfully got world info: {item_name}, Author: {item_author}")

            except ApiException as e:
                print(f"APIエラー during get_{item_type} for ID {item_id}: (Status: {e.status}) Reason: {e.reason}")
                error_details = f"Status: {e.status}"
                if e.body:
                    try:
                        error_body_json = json.loads(e.body)
                        if 'error' in error_body_json and 'message' in error_body_json['error']:
                            error_details += f", Message: {error_body_json['error']['message']}"
                        elif 'message' in error_body_json:
                             error_details += f", Message: {error_body_json['message']}"
                    except json.JSONDecodeError:
                        error_details += f", Body: {e.body[:100]}..."
                item_name = f"APIエラー ({error_details}) ({item_id})"
                item_author = "エラー"

            except Exception as e:
                print(f"予期しないエラー during API call or session setup for ID {item_id}: {e}")
                item_name = f"エラー取得 ({type(e).__name__}) ({item_id})"
                item_author = "エラー"

        timestamp = datetime.now().isoformat()
        image_formula = f"{image_url}"

        row_data = [item_type, item_id, item_name,timestamp,  image_formula , item_author]

        if sheets_available and item_id not in recorded_ids:
            print(f"DEBUG: Attempting to append row to worksheet: {row_data}")
            try:
                if worksheet:
                    worksheet.append_row(row_data)
                    print(f"DEBUG: Successfully appended row for ID: {item_id}")
                    recorded_ids.add(item_id)
                    print(f"DEBUG: Added ID {item_id} to recorded_ids. Current recorded_ids count: {len(recorded_ids)}")

            except Exception as sheet_e:
                print(f"エラー: スプレッドシート記録中にエラーが発生しました: {sheet_e}")
                root.after(0, messagebox.showerror, "スプレッドシート記録エラー", f"履歴記録中にエラーが発生しました: {sheet_e}")

        elif item_id in recorded_ids:
             print(f"ID '{item_id}' は既に記録済みです。シート記録はスキップされます。")

        elif not sheets_available:
             print(f"Google Sheets が利用できないため、ID '{item_id}' のシート記録はスキップされます。")

        root.after(0, load_history_from_sheet)

    except Exception as outer_e:
        print(f"予期しないエラー in record_item_history outer try block for ID {item_id}: {outer_e}")
        pass

    finally:
        print(f"DEBUG: Sleeping for {API_CALL_DELAY} seconds.")
        time.sleep(API_CALL_DELAY)
        print(f"DEBUG: Finished record_item_history for {item_id}")

# --- 関数: VRChat内でアバターを変更 ---
def change_avatar_in_vrchat(avatar_id):
    """
    指定されたアバターIDのアバターにVRChat内で変更する。
    """
    if not current_user or not cookie_jar or len(cookie_jar) == 0:
        root.after(0, messagebox.showerror, "アバター変更エラー", "VRChatにログインしていません。")
        print("エラー: VRChatにログインしていません。アバター変更をスキップします。")
        return

    print(f"DEBUG: VRChat内でアバター {avatar_id} に変更を試みています...")
    root.after(0, update_auth_status_label, f"アバター変更中: {avatar_id}...", "blue")

    try:
        # vrchatapi.ApiClient を正しく初期化し、クッキーを設定する
        config = vrchatapi.Configuration()
        api_client = vrchatapi.ApiClient(config)
        api_client.user_agent = f"MyVRChatIntegratedApp/1.0 (ChangeAvatar)"

        # cookie_jar からクッキーをヘッダー文字列に変換
        cookie_string = "; ".join([f"{cookie.name}={cookie.value}" for cookie in cookie_jar])
        if cookie_string:
            api_client.set_default_header("Cookie", cookie_string)
        else:
            root.after(0, messagebox.showerror, "アバター変更エラー", "認証クッキーが見つかりません。")
            print("エラー: 認証クッキーが見つかりません。アバター変更をスキップします。")
            root.after(0, update_auth_status_label, auth_status_label.cget("text").replace("アバター変更中", "ログイン済み"), "red")
            return
        
        avatars_api_instance = avatars_api.AvatarsApi(api_client)


        # PUT /avatars/{avatarId}/select を呼び出す
        avatars_api_instance.select_avatar(avatar_id=avatar_id)

        print(f"DEBUG: アバター {avatar_id} への変更リクエストを送信しました。")
        # アバター変更が成功した場合、履歴に記録
        root.after(0, update_auth_status_label, auth_status_label.cget("text").replace("アバター変更中...", "ログイン済み"), "green")


    except ApiException as e:
        error_message = f"VRChat APIエラー: アバター変更に失敗しました。(Status: {e.status}) Reason: {e.reason}"
        if e.body:
            try:
                error_body_json = json.loads(e.body)
                if 'error' in error_body_json and 'message' in error_body_json['error']:
                    error_message += f", Message: {error_body_json['error']['message']}"
            except json.JSONDecodeError:
                pass
        root.after(0, messagebox.showerror, "アバター変更エラー", error_message)
        print(f"エラー: {error_message}")
        root.after(0, update_auth_status_label, auth_status_label.cget("text").replace("アバター変更中", "ログイン済み"), "red") # エラー時は赤にするか、元に戻す
    except requests.exceptions.RequestException as e:
        root.after(0, messagebox.showerror, "ネットワークエラー", f"アバター変更リクエスト中にネットワークエラーが発生しました: {e}")
        print(f"ネットワークエラー: アバター変更リクエスト中にネットワークエラーが発生しました: {e}")
        root.after(0, update_auth_status_label, auth_status_label.cget("text").replace("アバター変更中", "ログイン済み"), "red")
    except Exception as e:
        root.after(0, messagebox.showerror, "アバター変更エラー", f"アバター変更中に予期しないエラーが発生しました: {e}")
        print(f"予期しないエラー: アバター変更中にエラーが発生しました: {e}")
        root.after(0, update_auth_status_label, auth_status_label.cget("text").replace("アバター変更中", "ログイン済み"), "red")
    finally:
        print("DEBUG: Change avatar logic finished.")
        time.sleep(API_CALL_DELAY) # API呼び出し間の遅延

# --- 関数: スプレッドシートから履歴全体を読み込み、GUIリストを更新 ---
def load_history_from_sheet():
    """スプレッドシートから履歴全体を読み込み、GUIのリストボックスを更新する"""
    global history_data_list, recorded_ids

    if gc is None or worksheet is None:
        print("Google スプレッドシートが設定されていないため履歴を読み込めません。")
        clear_gui_history_list()
        history_data_list = []
        recorded_ids.clear()
        print("DEBUG: history_data_list and recorded_ids cleared due to no Sheets config.")
        return

    print("スプレッドシートから履歴全体を読み込み中...")
    try:
        print(f"ワークsheet '{WORKSHEET_NAME}' から履歴を読み込み。")

        all_data = worksheet.get_all_values()
        if not all_data or len(all_data) < 2:
            print("スプレッドシートに履歴データが見つかりません。")
            clear_gui_history_list()
            history_data_list = []
            recorded_ids.clear()
            print("DEBUG: history_data_list and recorded_ids cleared due to no data in sheet.")
            return

        headers = all_data[0]
        header_indices = {}
        missing_headers = []

        for header in HISTORY_HEADERS:
            try:
                header_indices[header] = headers.index(header)
            except ValueError:
                missing_headers.append(header)

        if missing_headers:
             messagebox.showerror("スプレッドシートエラー", f"スプレッドシートに必須のヘッダーが見つかりません: {missing_headers}\n期待されるヘッダー: {HISTORY_HEADERS}")
             print(f"スプレッドシートに必須のヘッダーが見つかりません: {missing_headers}")
             clear_gui_history_list()
             history_data_list = []
             recorded_ids.clear()
             print("DEBUG: history_data_list and recorded_ids cleared due to missing headers.")
             return

        # 必須ヘッダーのインデックスを取得
        type_col_index = header_indices['Type']
        id_col_index = header_indices['ID']
        name_col_index = header_indices['Name']
        author_col_index = header_indices['author']
        timestamp_col_index = header_indices['Timestamp']
        image_url_col_index = header_indices['Image URL']

        history_data_list = []
        recorded_ids.clear()

        records = all_data[1:] # ヘッダーを除いたデータ行
        for i, row_data in enumerate(records):
            max_index = max(header_indices.values())
            if len(row_data) > max_index:
                 item_type = row_data[type_col_index].strip()
                 item_id = row_data[id_col_index].strip()
                 item_name = row_data[name_col_index].strip()
                 item_author = row_data[author_col_index].strip()
                 if not item_author:
                     item_author = "名無し"

                 timestamp_str = row_data[timestamp_col_index].strip()
                 image_url_cell_value = row_data[image_url_col_index].strip()
                 original_image_url = image_url_cell_value
                 if image_url_cell_value.startswith('=IMAGE('):
                     try:
                          start_quote = image_url_cell_value.find('"')
                          end_quote = image_url_cell_value.rfind('"')
                          if start_quote != -1 and end_quote != -1 and start_quote < end_quote:
                                 original_image_url = image_url_cell_value[start_quote + 1:end_quote]
                          else:
                                 print(f"警告: IMAGE関数のURL抽出に失敗しました: {image_url_cell_value}")
                     except Exception as e:
                          print(f"警告: IMAGE関数からのURL抽出中にエラーが発生しました: {e}")

                 if not item_id:
                      continue

                 item_data = {
                      "row_index": i + 2,
                      "type": item_type,
                      "id": item_id,
                      "name": item_name,
                      "author": item_author,
                      "timestamp": timestamp_str,
                      "image_url": original_image_url
                 }
                 history_data_list.append(item_data)
                 recorded_ids.add(item_id)

            else:
                 print(f"警告: 行 {i+2} のデータが不足しています。スキップします。")

        print(f"スプレッドシートから {len(history_data_list)} 件の履歴を読み込みました。")
        print(f"DEBUG: Current recorded_ids count after loading: {len(recorded_ids)}")

        root.after(0, update_gui_history_list)

    except gspread.exceptions.APIError as e:
         messagebox.showerror("Google Sheets APIエラー", f"スプレッドシートの読み込み中にAPIエラーが発生しました: {e}")
         print(f"Google Sheets APIエラー: スプレッドシートの読み込み中にAPIエラーが発生しました: {e}")
         clear_gui_history_list()
         history_data_list = []
         recorded_ids.clear()
    except Exception as e:
        messagebox.showerror("エラー", f"履歴の読み込み中に予期しないエラーが発生しました: {e}")
        print(f"予期しないエラー: 履歴の読み込み中に予期しないエラーが発生しました: {e}")
        clear_gui_history_list()
        history_data_list = []
        recorded_ids.clear()


# --- 関数: OSCサーバーを開始 ---
def start_osc_server():
    """OSCサーバーを開始する"""
    global osc_server_instance, is_osc_server_running
    if is_osc_server_running:
        print("OSCサーバーは既に実行中です。")
        return

    try:
        dispatcher_instance = dispatcher.Dispatcher()
        dispatcher_instance.map(OSC_AVATAR_CHANGE_ADDRESS, avatar_change_handler)
        dispatcher_instance.map(OSC_WORLD_CHANGE_ADDRESS, world_change_handler)

        server = osc_server.ThreadingOSCUDPServer(
            (OSC_LISTEN_IP, OSC_LISTEN_PORT), dispatcher_instance)
        osc_server_instance = server
        print(f"OSCサーバーを開始しました。 {OSC_LISTEN_IP}:{OSC_LISTEN_PORT} で待機中...")
        is_osc_server_running = True
        root.after(0, update_auth_status_label, auth_status_label.cget("text").replace(" | OSCエラー", "") + " | OSC待機中", "green")

        server.serve_forever()

    except OSError as e:
        if "address already in use" in str(e).lower():
             messagebox.showerror("OSCエラー", f"OSCポート {OSC_LISTEN_PORT} が既に使用されています。VRChatまたは他のアプリケーションがポートを使用している可能性があります。VRChat側のOSC設定を確認するか、他のポートを試してください。")
             print(f"OSCエラー: ポート {OSC_LISTEN_PORT} が既に使用されています。")
        else:
             messagebox.showerror("OSCエラー", f"OSCサーバーの開始に失敗しました: {e}")
             print(f"OSCサーバーの開始に失敗しました: {e}")
        is_osc_server_running = False
        if root:
             root.after(0, update_auth_status_label, auth_status_label.cget("text").replace(" | OSC待機中", "").replace(" | OSCエラー", "") + " | OSCエラー", "red")
    except Exception as e:
        messagebox.showerror("OSCエラー", f"OSCサーバーの開始中に予期しないエラーが発生しました: {e}")
        print(f"OSCサーバーの開始中に予期しないエラーが発生しました: {e}")
        is_osc_server_running = False
        if root:
             root.after(0, update_auth_status_label, auth_status_label.cget("text").replace(" | OSC待機中", "").replace(" | OSCエラー", "") + " | OSCエラー", "red")


# --- 関数: OSCサーバーを別スレッドで開始 ---
def start_osc_server_threaded():
    """OSCサーバーをバックグラウンドスレッドで開始する"""
    global osc_server_thread
    if osc_server_thread is None or not osc_server_thread.is_alive():
        print("OSCサーバー用スレッドを開始中...")
        osc_server_thread = threading.Thread(target=start_osc_server, daemon=True)
        osc_server_thread.start()
        print("OSCサーバー用スレッドを開始しました。")
    else:
        print("OSCサーバー用スレッドは既に実行中です。")

# --- 関数: OSCサーバーを停止 ---
def stop_osc_server():
    """OSCサーバーを停止する"""
    global osc_server_instance, is_osc_server_running
    if osc_server_instance:
        print("OSCサーバーをシャットダウン中...")
        osc_server_instance.shutdown()
        osc_server_instance.server_close()
        osc_server_instance = None
        is_osc_server_running = False
        print("OSCサーバーを停止しました。")
        if auth_status_label and root:
             root.after(0, lambda: update_auth_status_label(auth_status_label.cget("text").replace(" | OSC待機中", "").replace(" | OSCエラー", ""), "black"))


# --- 画像キャッシュ関連関数 ---
def get_image_cache_path(image_url):
    """画像URLからキャッシュファイルのパスを生成する"""
    if not image_url:
        return None
    import hashlib
    url_hash = hashlib.md5(image_url.encode('utf-8')).hexdigest()
    ext = os.path.splitext(image_url.split('?')[0])[-1]
    if not ext or len(ext) > 5: # 拡張子が不適切に長い場合やない場合
        ext = '.jpg' # デフォルトの拡張子
    return os.path.join(IMAGE_CACHE_DIR, f"{url_hash}{ext}")

def download_and_cache_image(image_url):
    """画像をダウンロードし、キャッシュディレクトリに保存する。成功した場合はキャッシュファイルのパスを返す。"""
    global cookie_jar

    if not image_url:
        print("警告: 画像URLが空です。")
        return None

    os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)
    cache_path = get_image_cache_path(image_url)

    if os.path.exists(cache_path):
        print(f"DEBUG: 画像は既にキャッシュされています: {cache_path}")
        return cache_path

    print(f"DEBUG: 画像をダウンロード中: {image_url} -> {cache_path}")
    
    download_session = requests.Session()
    
    # --- ★修正点★ ---
    # User-Agentヘッダーを追加
    download_session.headers.update({'User-Agent': 'MyVRChatIntegratedApp/1.0 (ImageDownloader)'})
    # --- ★修正ここまで★ ---
    
    # アプリケーションの認証済みクッキーをセッションにロード
    if cookie_jar:
        download_session.cookies.update(cookie_jar)
    
    try:
        response = download_session.get(image_url, stream=True, timeout=10)
        response.raise_for_status()

        with open(cache_path, 'wb') as out_file:
            for chunk in response.iter_content(chunk_size=8192):
                out_file.write(chunk)
        print(f"DEBUG: 画像をキャッシュしました: {cache_path}")
        return cache_path
    except requests.exceptions.RequestException as e:
        print(f"エラー: 画像のダウンロードに失敗しました '{image_url}': {e}")
        return None
    except Exception as e:
        print(f"エラー: 画像のキャッシュ中に予期しないエラーが発生しました '{image_url}': {e}")
        return None
    finally:
        download_session.close()

def load_image_for_display(image_path, max_size=(300, 300)):
    """指定されたパスから画像を読み込み、GUI表示用にリサイズしてImageTk.PhotoImageオブジェクトを返す。"""
    if not image_path or not os.path.exists(image_path):
        return None
    try:
        original_image = Image.open(image_path)
        original_image.thumbnail(max_size, Image.Resampling.LANCZOS) # 高品質なリサイズ
        return ImageTk.PhotoImage(original_image)
    except Exception as e:
        print(f"エラー: 画像の読み込みまたはリサイズに失敗しました '{image_path}': {e}")
        return None


def display_selected_image(event):
    """Treeviewで選択された項目の画像を表示する。"""
    global current_photo_image # PhotoImageオブジェクトへの参照を保持

    if not history_treeview or not image_display_label:
        return

    selected_items = history_treeview.selection()
    if not selected_items:
        image_display_label.config(image='', text="画像なし", compound="center")
        current_photo_image = None
        return

    selected_item = selected_items[0]
    # 'image_url' 列のデータを取得
    # Treeviewのvaluesは作成時に設定したcolumnsの順序に対応
    # columns = ('type', 'id', 'name', 'author', 'timestamp', 'image_url')
    # なので、image_urlはインデックス5
    try:
        item_values = history_treeview.item(selected_item, 'values')
        if len(item_values) > 5: # image_urlが5番目のインデックスにあることを確認
            image_url = item_values[5]
        else:
            image_url = ""
            print("警告: 選択されたアイテムに画像URLがありません。")

        if image_url and image_url != "不明":
            # 別スレッドで画像をダウンロードし、GUIスレッドで表示を更新
            def _load_and_display_in_thread():
                cached_path = download_and_cache_image(image_url)
                if cached_path:
                    photo_image = load_image_for_display(cached_path)
                    if photo_image:
                        # GUIスレッドで更新
                        root.after(0, lambda: set_image_on_label(photo_image, image_url))
                    else:
                        root.after(0, lambda: image_display_label.config(image='', text=f"画像の表示に失敗: {os.path.basename(cached_path)}", compound="center"))
                else:
                    root.after(0, lambda: image_display_label.config(image='', text="画像のダウンロードに失敗", compound="center"))
            
            # ロード中はメッセージを表示
            image_display_label.config(image='', text="画像読み込み中...", compound="center")
            threading.Thread(target=_load_and_display_in_thread).start()
        else:
            image_display_label.config(image='', text="画像URLがありません", compound="center")
            current_photo_image = None
    except Exception as e:
        print(f"画像表示中にエラーが発生しました: {e}")
        image_display_label.config(image='', text=f"エラー: {e}", compound="center")
        current_photo_image = None

def set_image_on_label(photo_image, url_for_debug=None):
    """PhotoImageをラベルに設定し、参照を保持するヘルパー関数。"""
    global current_photo_image
    current_photo_image = photo_image # GCされないように参照を保持
    if photo_image:
        image_display_label.config(image=photo_image, text="", compound="none")
    else:
        # 画像がNoneの場合、テキスト表示に戻す
        image_display_label.config(image='', text=f"画像表示エラー: {url_for_debug}", compound="center")


# --- GUI関連関数 ---
def create_gui():
    """Tkinter GUIを作成する"""
    global root, username_entry, password_entry, twofactor_entry, login_button, auth_status_label
    global login_frame, history_frame, history_treeview, image_display_label

    root = tk.Tk()
    root.title("VRChat アバター/ワールド 履歴トラッカー")
    root.geometry("1200x700") # ウィンドウサイズを調整
    root.minsize(900, 600) # 最小サイズを設定

    # メインフレーム
    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    main_frame.columnconfigure(0, weight=2) # 履歴リスト側を広く
    main_frame.columnconfigure(1, weight=1) # 画像表示側を狭く
    main_frame.rowconfigure(1, weight=1)

    # 認証情報入力フレーム
    login_frame = ttk.LabelFrame(main_frame, text="VRChat ログイン", padding="10")
    login_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady="5") # 2列にまたがる
    login_frame.columnconfigure(1, weight=1)

    ttk.Label(login_frame, text="ユーザー名:").grid(row=0, column=0, sticky=tk.W, padx="5", pady="2")
    username_entry = ttk.Entry(login_frame, width=40)
    username_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx="5", pady="2")

    ttk.Label(login_frame, text="パスワード:").grid(row=1, column=0, sticky=tk.W, padx="5", pady="2")
    password_entry = ttk.Entry(login_frame, width=40, show="*")
    password_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx="5", pady="2")

    ttk.Label(login_frame, text="2FAコード:").grid(row=2, column=0, sticky=tk.W, padx="5", pady="2")
    twofactor_entry = ttk.Entry(login_frame, width=40)
    twofactor_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx="5", pady="2")
    twofactor_entry.grid_remove()

    login_button = ttk.Button(login_frame, text="ログイン", command=on_login_button_click)
    login_button.grid(row=3, column=0, columnspan=2, pady="10")

    auth_status_label = ttk.Label(login_frame, text="初期化中...", foreground="black")
    auth_status_label.grid(row=4, column=0, columnspan=2)

    # 履歴表示フレーム (左側)
    history_frame = ttk.LabelFrame(main_frame, text="履歴", padding="10")
    history_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5)) # 右側に少しパディング
    history_frame.columnconfigure(0, weight=1)
    history_frame.rowconfigure(0, weight=1)

    # Treeview の設定 (画像URL列を追加)
    columns = ('type', 'id', 'name', 'author', 'timestamp', 'image_url')
    history_treeview = ttk.Treeview(history_frame, columns=columns, show='headings')

    # 各列のヘッダーを設定
    history_treeview.heading('type', text='種類', anchor=tk.W)
    history_treeview.heading('id', text='ID', anchor=tk.W)
    history_treeview.heading('name', text='名前', anchor=tk.W)
    history_treeview.heading('author', text='作者', anchor=tk.W)
    history_treeview.heading('timestamp', text='タイムスタンプ', anchor=tk.W)
    history_treeview.heading('image_url', text='画像URL', anchor=tk.W) # 画像URL列のヘッダー

    # 各列の幅を設定 (image_urlはデフォルトでは非表示)
    history_treeview.column('type', width=80, stretch=tk.NO)
    history_treeview.column('id', width=150, stretch=tk.NO)
    history_treeview.column('name', width=200, stretch=tk.YES)
    history_treeview.column('author', width=120, stretch=tk.NO)
    history_treeview.column('timestamp', width=150, stretch=tk.NO)
    history_treeview.column('image_url', width=0, stretch=tk.NO) # 初期状態では幅0で非表示

    history_treeview.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

    # スクロールバー
    vsb = ttk.Scrollbar(history_frame, orient="vertical", command=history_treeview.yview)
    vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
    history_treeview.configure(yscrollcommand=vsb.set)

    # Treeview項目選択時のイベントバインディング
    history_treeview.bind('<<TreeviewSelect>>', display_selected_image)

    # 画像表示フレーム (右側)
    image_display_frame = ttk.LabelFrame(main_frame, text="選択画像", padding="10")
    image_display_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
    image_display_frame.columnconfigure(0, weight=1)
    image_display_frame.rowconfigure(0, weight=1) # 画像ラベルが伸縮するように設定

    # 画像表示用ラベル
    image_display_label = ttk.Label(image_display_frame, text="アイテムを選択してください",
                                    relief="solid", borderwidth=1, anchor="center")
    image_display_label.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(0, 5)) # 少し下にパディングを追加してボタンとの間隔を確保

    # --- ここから新しいボタンを追加 ---
    change_avatar_button = ttk.Button(image_display_frame, text="このアバターに変更", command=on_change_avatar_button_click)
    change_avatar_button.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0)) # 画像ラベルの下に配置
    # --- ここまで新しいボタンを追加 ---

    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.after(100, load_history_from_sheet)


# --- GUI更新関数 ---
def update_auth_status_label(text, color="black"):
    """認証ステータスラベルを更新する"""
    if auth_status_label:
        auth_status_label.config(text=text, foreground=color)
    if login_button and ("ログイン中" not in text and "検証中" not in text):
         login_button.config(state=tk.NORMAL)


def show_manual_login_ui(requires_2fa=False):
    """手動ログインUIを表示する"""
    global login_frame, username_entry, password_entry, twofactor_entry, login_button
    if login_frame:
        login_frame.grid()
        if requires_2fa:
            if twofactor_entry: twofactor_entry.grid()
            if login_button: login_button.config(text="検証")
        else:
            if twofactor_entry: twofactor_entry.grid_remove()
            if login_button: login_button.config(text="ログイン")
        if login_button:
             login_button.config(state=tk.NORMAL)


def hide_manual_login_ui():
    """手動ログインUIを隠す"""
    global login_frame
    if login_frame:
        login_frame.grid_remove()

def on_login_button_click():
    """ログインボタンがクリックされた時の処理"""
    global VRC_USERNAME, VRC_PASSWORD
    if username_entry and password_entry and twofactor_entry and login_button:
        VRC_USERNAME = username_entry.get()
        VRC_PASSWORD = password_entry.get()
        two_factor_code = twofactor_entry.get()

        if not VRC_USERNAME or not VRC_PASSWORD:
            messagebox.showwarning("入力不足", "ユーザー名とパスワードを入力してください。")
            return

        login_button.config(state=tk.DISABLED)
        root.after(0, update_auth_status_label, "ログイン処理中...", "blue")

        login_thread = threading.Thread(target=authenticate_vrchat_with_requests, args=(VRC_USERNAME, VRC_PASSWORD, two_factor_code))
        login_thread.start()


    else:
        print("GUI widgets not found for login.")
# --- 関数: アバター変更ボタンがクリックされた時の処理 ---
def on_change_avatar_button_click():
    """
    「このアバターに変更」ボタンがクリックされた時の処理。
    Treeviewで選択されているアバターにVRChat内で変更する。
    """
    selected_items = history_treeview.selection()
    if not selected_items:
        messagebox.showwarning("選択なし", "アバターを変更するには、履歴リストからアバターを選択してください。")
        return

    selected_item = selected_items[0]
    item_values = history_treeview.item(selected_item, 'values')

    # columns = ('type', 'id', 'name', 'author', 'timestamp', 'image_url')
    item_type = item_values[0]
    item_id = item_values[1]
    item_name = item_values[2] # 表示用に名前も取得

    if item_type != 'Avatar':
        messagebox.showwarning("非アバター", f"'{item_name}' はアバターではありません。アバターのみ変更可能です。")
        return

    if not item_id or item_id == "不明":
        messagebox.showwarning("無効なID", "選択されたアイテムに有効なアバターIDがありません。")
        return
    
    # ユーザーに確認
    #if not messagebox.askyesno("アバター変更確認", f"VRChat内でアバターを '{item_name}' に変更しますか？"):
        return

    # 別スレッドでAPI呼び出しを実行
    change_thread = threading.Thread(target=change_avatar_in_vrchat, args=(item_id,))
    change_thread.start()

def clear_gui_history_list():
    """GUIの履歴表示エリア (Treeview) をクリアする"""
    if history_treeview:
        for item in history_treeview.get_children():
            history_treeview.delete(item)

def update_gui_history_list():
    """グローバルな history_data_list に基づいてGUIの履歴表示リストを更新する"""
    if history_treeview is None:
         print("history_treeview not initialized. Skipping GUI update.")
         return

    print(f"GUI履歴リストを更新中。アイテム数: {len(history_data_list)}")

    clear_gui_history_list()

    # アイテムを逆順に表示 (新しいものが上にくるように)
    for item_data in reversed(history_data_list):
        item_type = item_data.get("type", "不明")
        item_id = item_data.get("id", "不明")
        item_name = item_data.get("name", "不明")
        item_author = item_data.get("author", "不明")
        timestamp_str = item_data.get("timestamp", "不明")
        image_url = item_data.get("image_url", "") # 画像URLを取得

        try:
            dt_object = datetime.fromisoformat(timestamp_str)
            formatted_timestamp = dt_object.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            formatted_timestamp = timestamp_str

        # Treeview に行を挿入
        # image_url を values に追加
        history_treeview.insert('', 'end', values=(item_type, item_id, item_name, item_author, formatted_timestamp, image_url))


# --- アプリケーション終了時の処理 ---
def on_closing():
    """ウィンドウが閉じられるときの処理"""
    if messagebox.askokcancel("終了", "アプリケーションを終了しますか？"):
        print("アプリケーションを終了します...")
        stop_osc_server()
        time.sleep(0.1)

        if root:
            root.destroy()
        print("アプリケーション終了完了。")


# --- メイン処理 ---
if __name__ == "__main__":
    sheets_ready = setup_google_sheets()

    create_gui()

    root.after(100, attempt_auto_login_with_requests)

    root.mainloop()

    print("アプリケーションが終了しました。")
