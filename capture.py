import sys
import time
import os
import subprocess
import ctypes
import winsound
from datetime import datetime
from dotenv import load_dotenv
from PIL import Image
import pytesseract

from google import genai
from google.genai import types
import wave
import json
import array
import re
import glob

# .env ファイルを読み込んで API キーを環境変数にセットするよ！
load_dotenv()

from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pyautogui
import pyperclip
import win32clipboard

# --- OCRのパス設定 ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# --------------------

from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QSizePolicy, QInputDialog, QMessageBox, QScrollArea, QSizeGrip, QTextBrowser, QGraphicsOpacityEffect
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSize
from PyQt6.QtGui import QPixmap, QFont, QPainter, QColor, QPainterPath, QIcon

# コマンドプロンプトで絵文字を出すとエラーになることがあるので、UTF-8に強制するよ！
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

try:
    import win32gui
    import win32ui
except ImportError:
    print("pywin32がないみたい！ 'pip install pywin32' でインストールしてね！")
    exit(1)

# クリップボードに画像をコピーする関数（PyAutoGUIで貼り付ける準備）
def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def image_to_clipboard(img):
    import io
    output = io.BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:] # BMPヘッダを削除
    send_to_clipboard(win32clipboard.CF_DIB, data)

SAVE_DIR = "captures"
os.makedirs(SAVE_DIR, exist_ok=True)
MAX_IMAGES = 5
WINDOW_TITLE = "Winning Post 10 2025"

# 🧠 記憶ファイルのパス
HISTORY_FILE = os.path.join(SAVE_DIR, "history_log.json")
ACHIEVEMENT_FILE = os.path.join(SAVE_DIR, "achievements.json")

def background_capture(hwnd):
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    width = right - left
    height = bottom - top
    
    if width <= 0 or height <= 0: return None

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)

    result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

    img = None
    if result == 1:
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    return img
    
# ==========================================
# 🌐 Selenium ブラウザ管理クラス
# ==========================================
class GeminiBrowser:
    def __init__(self):
        self.driver = None # type: ignore
        self.setup_browser()

    def setup_browser(self):
        print("\n⚠️ 【超重要】Edgeが裏で動いているとエラーになるよ！起動中のEdgeは『全て』閉じておいてね！")
        edge_options = Options()
        
        # 💡 うまくお忍びで動かすためのマスター特製・偽装オプション！
        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-dev-shm-usage")
        
        # マスターのログイン済みプロファイルをガッチリ指定！
        edge_options.add_argument(r"user-data-dir=C:\Users\zebel\AppData\Local\Microsoft\Edge\User Data")
        edge_options.add_argument(r"profile-directory=Profile 1")
        
        # 自動操作（ロボット）であることをGoogleにバレないようにする魔法！
        edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        edge_options.add_experimental_option('useAutomationExtension', False)
        edge_options.add_argument('--disable-blink-features=AutomationControlled')
        
        # 仮想デスクトップ2で開くことを想定し、全画面表示
        edge_options.add_argument("--start-maximized")

        try:
            # 💡 マスターが直接ダウンロードした msedgedriver.exe の場所を指定するよ！
            # まずは同じフォルダ（winpost_capture内）を探すようにしておくね！
            driver_path = r"msedgedriver.exe"
            
            if not os.path.exists(driver_path):
                print(f"💦 {driver_path} が見つからないよ！")
                print("💡 https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/ から今のEdgeのバージョンに合うものをダウンロードして、")
                print("💡 この capture.py と同じフォルダに「msedgedriver.exe」という名前で置いてね！")
                return

            service = EdgeService(executable_path=driver_path)
            self.driver = webdriver.Edge(service=service, options=edge_options)
            self.driver.get("https://gemini.google.com/")
        except Exception as e:
            import traceback
            print("\n=======================================================")
            print("💦 ブラウザ起動に大失敗しちゃった！詳しいエラー原因はこちら！")
            traceback.print_exc()
            print("=======================================================\n")
            print("💡 Edgeがすでに起動しているとエラーになるよ！全部閉じてから試してみてね！")

    def get_response_count(self):
        try:
            elements = self.driver.find_elements(By.CSS_SELECTOR, "message-content, div.model-response-text, div[data-message-author-role='model']")
            return len(elements)
        except:
            return 0

    def get_latest_response_text(self):
        # 最新の返答エリアから直接テキストを抽出する
        try:
            elements = self.driver.find_elements(By.CSS_SELECTOR, "message-content, div.model-response-text, div[data-message-author-role='model']")
            
            if elements:
                latest_element = elements[-1] # 一番最後のメッセージ
                
                try:
                    # JavaScriptを使って、要素の中にあるテキスト（innerText）を直接取得する！
                    extracted_text = self.driver.execute_script("return arguments[0].innerText;", latest_element)
                    
                    if not extracted_text:
                        return ""
                        
                    # 日本語テキストとしてきれいに掃除
                    cleaned_text = "".join([c for c in extracted_text if c.isprintable() or c in ['\n', '\r']]).strip()
                    return cleaned_text
                except Exception as e:
                    print(f"📝 テキスト抽出エラー: {e}")
                    return ""
                    
            return ""
        except Exception as e:
            print(f"Browser Text Extraction Error: {e}")
            return ""

    def close(self):
        if self.driver:
            self.driver.quit()


# ==========================================
# 🧵 キャプチャとAI通信を裏で行うスレッド
# ==========================================
class CaptureThread(QThread):
    new_message = pyqtSignal(str)   # あいちゃんのセリフ用
    status_update = pyqtSignal(str) # 状況報告用

    def __init__(self):
        super().__init__()
        self.running = True
        self.user_message = ""
        self.force_capture = False # ユーザーがチャットを送った時にすぐ反応するためのフラグ
        self.browser = None # type: ignore
        
        # 💡 スマート検出用の記憶（前回見た画面と、前回のセリフ、前回の自動送信時刻）
        self.last_img_hash = None
        self.last_comment = None
        self.last_successful_auto_send = 0
        self.genai_client = None # TTS用クライアント

        # 🧠 記憶システム
        self.history_buffer = [] # 直近のやり取り (5-10件)
        self.achievements = []   # 重要な成果
        self.load_memory()       # 永続ファイルから読み込み

    def load_memory(self):
        """PCに保存されたJSONファイルから思い出を読み込むよ！"""
        try:
            if os.path.exists(HISTORY_FILE):
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    full_history = json.load(f)
                    # 直近10件だけバッファに入れる
                    self.history_buffer = full_history[-10:]
            
            if os.path.exists(ACHIEVEMENT_FILE):
                with open(ACHIEVEMENT_FILE, "r", encoding="utf-8") as f:
                    self.achievements = json.load(f)
            print(f"📖 [Memory] 記憶をロードしたよ！(履歴:{len(self.history_buffer)}件, 成果:{len(self.achievements)}件)")
        except Exception as e:
            print(f"💦 [Memory] 記憶のロードに失敗: {e}")

    def save_history(self, scene, user_msg, ai_res):
        """新しい出来事をJSONファイルに刻むよ！"""
        try:
            entry = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "scene": scene,
                "user": user_msg,
                "ai": ai_res
            }
            
            # --- 全履歴の保存 ---
            full_history = []
            if os.path.exists(HISTORY_FILE):
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    full_history = json.load(f)
            
            full_history.append(entry)
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(full_history, f, ensure_ascii=False, indent=2)
            
            # バッファも更新
            self.history_buffer = full_history[-10:]

            # --- 重要成果（Achievements）の抽出 ---
            # 「1着」「勝利」「優勝」「確定」などのキーワードがあれば追加
            # AIの返答か、場面判定の結果に含まれていれば成果とするよ！
            important_keywords = ["1着", "勝利", "優勝", "制覇", "確定", "能力S", "素質すごすぎ"]
            if any(kw in scene for kw in important_keywords) or any(kw in ai_res for kw in important_keywords):
                if entry not in self.achievements:
                    self.achievements.append(entry)
                    # 成果も永続保存
                    with open(ACHIEVEMENT_FILE, "w", encoding="utf-8") as f:
                        json.dump(self.achievements, f, ensure_ascii=False, indent=2)
                    print(f"🏆 [Memory] 重要な成果を記録したよ！: {scene}")

        except Exception as e:
            print(f"💦 [Memory] 記憶の保存に失敗: {e}")

    def calc_image_hash(self, img):
        """画像を8x8サイズの白黒に縮小して、ピクセルの明暗から64桁のハッシュを作るよ！"""
        # 単純な平均ハッシュ（Average Hash）で画面の全体的な変化だけをざっくり捉える！
        small_img = img.resize((8, 8)).convert('L')
        pixels = list(small_img.getdata())
        avg = sum(pixels) / len(pixels)
        hash_val = 0
        for p in pixels:
            hash_val <<= 1
            if p >= avg:
                hash_val |= 1
        return hash_val

    def hamming_distance(self, h1, h2):
        """2つのハッシュの差分（違うビットの数）を計算するよ！"""
        return bin(h1 ^ h2).count('1')

    def detect_scene_and_get_prompt(self, img):
        """画像から文字を読んで、今の画面「場面」を推理するよ！🔍"""
        try:
            # 画面全体だと時間とリソースがかかるので、上から40%（タイトルやヘッダがある部分）だけ切り抜いてOCRする高速化！
            width, height = img.size
            top_half = img.crop((0, 0, width, int(height * 0.4)))
            
            # 画像からテキストを抽出（日本語設定）
            extracted_text = pytesseract.image_to_string(top_half, lang='jpn')
            
            # 🎙️ コンテキスト文字列の構築（記憶の注入）
            context_str = "\n\n### [あいの記憶（重要成果と最近の出来事）]\n"
            if self.achievements:
                context_str += "【これまでの重要成果】:\n"
                # 最新の3件の成果を表示
                for a in self.achievements[-3:]:
                    context_str += f"- {a['timestamp']}: {a['scene']} (結果: {a['ai'][:30]}...)\n"
            
            if self.history_buffer:
                context_str += "\n【直近の出来事】:\n"
                for h in self.history_buffer[-5:]: # 最新5件
                    if h['user']: context_str += f"- マスター: {h['user']}\n"
                    context_str += f"- あい: {h['ai'][:50]}...\n"
            
            context_str += "\n※上記の記憶をふまえ、文脈に沿った一言を添えて実況してね！🎙️✨\n"

            # 💡 ルール1：レース結果
            if any(keyword in extracted_text for keyword in ["着順", "結果", "確定"]):
                print("🎯 [Scene] レース結果画面を検出しました！")
                return (
                    "画像は競馬ゲームのレース結果画面です。"
                    "以下の構成で【超ハイテンション＆絶好調な明るいトーン】で喋って！♪\n"
                    "1行目： [喜][悲][驚][怒] の中から1つだけ選び、その後に最高に明るい感情たっぷりの短い叫びを（例：[喜]やったぁぁ！1着だよマスター！♪）\n"
                    "2行目以降： レース結果の具体的な内容や励ましの言葉（2文程度）\n"
                    "※語尾に『♪』や『！』を多用して、あいの元気を爆発させてね！"
                    + context_str
                )
                
            # 💡 ルール2：個別能力
            elif any(keyword in extracted_text for keyword in ["競走馬詳細", "能力", "適性"]):
                print("🎯 [Scene] 馬の個別能力画面を検出しました！")
                return (
                    "画像は競馬ゲームの競走馬の能力画面です。"
                    "以下の構成で【最高に明るい分析トーン】で喋って！♪\n"
                    "1行目： [驚][喜] などの感情タグから始め、能力を見た第一印象を明るく叫んで！（例：[驚]うわぁ！この子の素質すごすぎだよぉ！♪）\n"
                    "2行目以降： その子の強みや弱みをズバッと分析し、育成のヒントになるアドバイス（1〜2文程度）\n"
                    + context_str
                )
                
            # 💡 ルール3：幼駒・セリ
            elif any(keyword in extracted_text for keyword in ["幼駒", "セリ", "種付け"]):
                print("🎯 [Scene] 幼駒／セリ／種付け画面を検出しました！")
                return (
                    "画像は競馬ゲームの幼駒やセリの画面です。"
                    "以下の構成で【ワクワク感いっぱいの明るいトーン】でお買い物アドバイスをして！♪\n"
                    "1行目： [喜][驚] などの感情タグから始め、直感的な明るい叫びを！（例：[喜]この子、絶対将来化けるよマスター！わくわくしちゃう♪）\n"
                    "2行目以降： マスターが買うべきか育てるべきか、背中を押す期待の一言（1〜2文程度）\n"
                    + context_str
                )
                
            # 💡 ルール4：通常・上記以外
            else:
                print("🎯 [Scene] 通常画面（メニュー等）と判定しました")
                return (
                    "画像は競馬ゲームのプレイ画面です。"
                    "今の状況を見て、親しみやすく元気な口調で解説して！\n"
                    "1行目： 状況に合わせた感情タグ（例：[喜]など）から書き始め、今の気分を短く一言！\n"
                    "2行目以降： 今の画面で何が起きているか、1〜2文で簡潔に解説してね。"
                    + context_str
                )
                
        except Exception as e:
            print(f"💦 OCRでの画面判定に失敗しました: {e}")
            # エラー時は汎用プロンプト
            return (
                "画像は競馬ゲームのプレイ画面です。"
                "今の状況を見て、親しみやすく元気な口調で解説して！\n"
                "1行目： [喜][悲][驚][怒] のいずれか1つの感情タグから書き始め、その後に感情たっぷりの短い叫び（例：[喜]やったぁぁ！1着だよ！）\n"
                "2行目以降： 今の画面で何が起きているか、1〜2文で簡潔に解説してね。"
            )

    def cleanup_voice_files(self):
        """最新の3つだけ残して、古い音声ファイルを削除するよ！🧹"""
        try:
            files = glob.glob(os.path.join(SAVE_DIR, "voice_*.wav"))
            # 更新日時順にソート（新しい順）
            files.sort(key=os.path.getmtime, reverse=True)
            
            # 3個以上ある場合は古いものを削除
            if len(files) > 3:
                for old_file in files[3:]:
                    try:
                        os.remove(old_file)
                        print(f"🧹 [TTS] 古い音声ファイルを削除したよ: {os.path.basename(old_file)}")
                    except Exception as e:
                        print(f"💦 [TTS] ファイル削除に失敗（再生中かも？）: {e}")
        except Exception as e:
            print(f"💦 [TTS] クリーンアップ中にエラー: {e}")

    def generate_voice(self, text):
        """Gemini Voice TTS を使って音声を生成・再生するよ！"""
        if not self.genai_client: return
        
        try:
            # 💡 [期待] など、カッコで囲まれたタグを正規表現で綺麗に取り除く
            clean_text = re.sub(r'\[.*?\]', '', text).strip()
            # 文末の絵文字なども除去しておくとより自然（TTSが読み上げられない場合がある）
            clean_text = "".join([c for c in clean_text if not (0x1F600 <= ord(c) <= 0x1F64F or 0x1F300 <= ord(c) <= 0x1F5FF)])
            
            if not clean_text:
                print("📝 [TTS] 読み上げるテキストが空だったのでスキップしたよ！（タグのみだった可能性あり）")
                return
            
            print(f"🎙️ [TTS] 生成開始: 「{clean_text}」")

            # 🎙️ 早口＆高音（ピッチ上げ）の指示タグを付与！
            voiced_text = f"[extremely fast][bright] {clean_text}"

            print(f"🎙️ [TTS] あいの叫びを生成中: 「{clean_text}」 (タグ付与: {voiced_text})")
            response = self.genai_client.models.generate_content(
                model="gemini-2.5-flash-preview-tts",
                contents=voiced_text,
                config=types.GenerateContentConfig(
                    response_modalities=["AUDIO"],
                    speech_config=types.SpeechConfig(
                        voice_config=types.VoiceConfig(
                            prebuilt_voice_config=types.PrebuiltVoiceConfig(
                                voice_name="Sulafat"
                            )
                        )
                    )
                )
            )
            
            if response.candidates and response.candidates[0].content.parts:
                print(f"✅ [TTS] レスポンス取得成功！パーツ数: {len(response.candidates[0].content.parts)}")
                for part in response.candidates[0].content.parts:
                    if part.inline_data:
                        pcm_data = part.inline_data.data
                        print(f"✅ [TTS] 音声データ受信完了！ ({len(pcm_data)} bytes)")

                        # 🔊 音量調整 (極限まで落としてみるよ！ 20%設定)
                        try:
                            # 'h' は 16-bit signed short. Geminiの音声が2バイト単位であることを確認して処理するよ！
                            if len(pcm_data) % 2 == 0:
                                audio_array = array.array('h', pcm_data)
                                for i in range(len(audio_array)):
                                    # 0.2倍にスケーリング。念のため範囲内に収めるよ！
                                    val = int(audio_array[i] * 0.2)
                                    audio_array[i] = max(-32768, min(32767, val))
                                pcm_data = audio_array.tobytes()
                                print("🔉 [TTS] 音量を約20%に調整して保存したよ！")
                            else:
                                print("⚠️ [TTS] データ長が2の倍数じゃないので音量調整をスキップしたよ。")
                        except Exception as ve:
                            print(f"⚠️ [TTS] 音量調整中にエラーが発生したよ: {ve}")
                        
                        # 💾 タイムスタンプ付きのユニークなファイル名にするよ！
                        unique_id = int(time.time())
                        temp_file = os.path.join(SAVE_DIR, f"voice_{unique_id}.wav")
                        
                        # Gemini TTS preview は生のPCM(16bit LE, 24kHz, Mono)で出すことが多いのでWAVにラップするよ！
                        with wave.open(temp_file, 'wb') as wav_file:
                            wav_file.setnchannels(1)   # モノラル
                            wav_file.setsampwidth(2)   # 16-bit (2 bytes)
                            wav_file.setframerate(24000) # 24kHz
                            wav_file.writeframes(pcm_data)
                        
                        # SND_ASYNC で非同期再生（実況を止めない！）
                        winsound.PlaySound(temp_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
                        
                        # 🧹 古いファイルはお掃除！
                        self.cleanup_voice_files()
                        break
        except Exception as e:
            print(f"💦 [TTS] 音声生成エラー: {e}")
            import traceback
            traceback.print_exc()

    def send_chat(self, text):
        """マスターからのチャットを受け取るよ！"""
        self.user_message = text
        self.force_capture = True

    def run(self):
        self.status_update.emit("🌟 準備中…まずは裏で残っているブラウザをお掃除するよ！")
        
        # Gemini Voice TTS の初期化
        api_key = os.environ.get("GEMINI_API_KEY")
        if api_key:
            self.genai_client = genai.Client(api_key=api_key)
        else:
            print("⚠️ GEMINI_API_KEYが見つからないため、音声機能は無効になります。")

        # 起動前に既存のEdgeプロセスをすべて強制終了する（プロファイルロック回避のため）
        try:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] 🧹 既存の msedge.exe プロセスを終了しています...")
            # /F (強制終了), /IM (イメージ名), /T (子プロセスも終了)
            subprocess.run(["taskkill", "/F", "/IM", "msedge.exe", "/T"], capture_output=True, text=True)
            time.sleep(2) # プロセスが完全に落ちるまで少し待機
        except Exception as e:
            print(f"プロセスのキルに失敗しました: {e}")

        self.status_update.emit("🌟 お掃除完了！ブラウザを起動しているよ！")
        
        # 別スレッドからSeleniumを起動
        self.browser = GeminiBrowser()
        if not self.browser.driver:
            self.status_update.emit("💦 ブラウザが開けなかったよ…プロセス終了してね！")
            return

        self.status_update.emit("⏳ Geminiの読み込みを待っているよ…")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ⏳ 起動時：入力欄が表示されるのを待機しています (最大30秒)...")
        
        # 🚀 高速化: 無駄な10秒固定待機を廃止し、入力欄が見つかり次第すぐ次に進む！
        ready = False
        selectors = ["rich-textarea p", "rich-textarea", "div.ql-editor", "div[role='textbox']", "div[contenteditable='true']"]
        
        for _ in range(30):
            if not self.running: return
            try:
                found = False
                for s in selectors:
                    if self.browser.driver.find_elements(By.CSS_SELECTOR, s):
                        found = True
                        break
                if found:
                    ready = True
                    break
            except:
                pass
            time.sleep(1)
            
        if ready:
            self.status_update.emit("🌟 準備完了！ゲームの画面が見えたら自動で実況を開始するね！")
            print(f"[{datetime.now().strftime('%H:%M:%S')}] ⚡ Geminiの入力欄を発見しました！")
            time.sleep(1) # ちょっとだけ余裕を持たせる
        else:
            self.status_update.emit("💦 待機タイムアウト：入力欄が見つからなかったよ！手動で確認してね！")
            print("💦 待機タイムアウト")
            
        while self.running:
            # ユーザー入力による割り込みがなければ、定期的に少し待ってからチェック（CPU負荷軽減と高頻度ポーリングの両立！）
            for _ in range(3):
                if not self.running or self.force_capture: break
                time.sleep(1)
                
            self.force_capture = False
            error_429_occurred = False
            skipped_by_diff = False

            try:
                hwnd = win32gui.FindWindow(None, WINDOW_TITLE)
                if hwnd:
                    if win32gui.IsIconic(hwnd):
                        self.status_update.emit("ゲームが最小化されてるみたい！")
                    else:
                        img = background_capture(hwnd)
                        if img:
                            # --- 💡 スマート画面変化検出＆クールダウン ---
                            current_hash = self.calc_image_hash(img)
                            is_auto_mode = not bool(self.user_message) # マスターの質問がない＝自動実況モード
                            
                            if is_auto_mode:
                                now_ts = time.time()
                                # 1. クールダウンチェック（連続実況を防ぐため、送信後は最低45秒空ける）
                                if now_ts - self.last_successful_auto_send < 45:
                                    continue # 上部の3秒待機に戻る
                                
                                # 2. 画面変化の検出
                                if self.last_img_hash is not None:
                                    dist = self.hamming_distance(self.last_img_hash, current_hash)
                                    if dist < 8: # ハッシュ距離が8未満なら「同じ画面」とみなす
                                        continue # 画面が変わってないなら何もしないで次へ
                                        
                            # 画面が変わった、またはマスターからの強制指示なので処理を進める！
                            self.last_img_hash = current_hash
                            # ---------------------------------------------
                            
                            # 🖼️ 画像をリサイズして軽くする！（横幅1024px）
                            target_width = 1024
                            if img.width > target_width:
                                ratio = target_width / img.width
                                new_height = int(img.height * ratio)
                                # 💡 ANTIALIAS は古くて非推奨なので LANCZOS を使うのがより良いよ！
                                img = img.resize((target_width, new_height), Image.Resampling.LANCZOS)
                                
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            # 💡 拡張子を png から jpg に変更！
                            filepath = os.path.join(SAVE_DIR, f"{timestamp}.jpg")
                            # 💡 JPGで保存（品質85で圧縮してさらに軽く！）
                            img.save(filepath, format="JPEG", quality=85)
                            
                            # 古い画像のお掃除 (jpg対象に変更)
                            saved_files = [f for f in os.listdir(SAVE_DIR) if f.endswith(".jpg")]
                            if len(saved_files) > MAX_IMAGES:
                                saved_files.sort()
                                for old_file in saved_files[:len(saved_files) - MAX_IMAGES]:
                                    try: os.remove(os.path.join(SAVE_DIR, old_file))
                                    except: pass
                            
                            self.status_update.emit("🤖 ブラウザに画像を貼り付けてるよ…")
                            
                            # クリップボードに画像をコピー
                            image_to_clipboard(img)
                            
                             # ブラウザに入力フォーカスを合わせる
                            time.sleep(1)
                            
                            # ✨ 超重要: Geminiの入力欄を意地でも見つけてフォーカスを当てる！
                            try:
                                print("🔍 入力欄を探しています...")
                                selectors = [
                                    "rich-textarea p",
                                    "rich-textarea",
                                    "div.ql-editor",
                                    "div[role='textbox']",
                                    "div[contenteditable='true']"
                                ]
                                input_area = None
                                # 複数の書き込み可能なエリアの候補を順番に試して、一番最初に見つかったものを採用する
                                for selector in selectors:
                                    try:
                                        input_area = WebDriverWait(self.browser.driver, 3).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                                        )
                                        if input_area:
                                            print(f"🎯 入力欄を発見しました！(selector: {selector})")
                                            break
                                    except:
                                        pass
                                
                                if input_area:
                                    # JavaScriptを使って強制的にフォーカスを奪う！（これがないとアドレスバーから動かないことがある）
                                    self.browser.driver.execute_script("arguments[0].focus();", input_area)
                                    time.sleep(0.5)
                                    # 念のため物理クリックもしておく
                                    input_area.click()
                                    time.sleep(0.5)
                                    from selenium.webdriver.common.keys import Keys
                                    from selenium.webdriver.common.action_chains import ActionChains

                                    # --- 💡 場面判定とプロンプトの決定 ---
                                    # ユーザーが明示的にチャットした場合はユーザーの言葉を優先、それ以外はOCRによる自動プロンプト
                                    if self.user_message:
                                        send_text = self.user_message
                                        self.status_update.emit("🗣️ マスターのメッセージを手打ち中…")
                                        self.user_message = "" # 使用したらクリア
                                    else:
                                        self.status_update.emit("👁️ 画面を読んで状況を分析中…")
                                        prompt_to_send = self.detect_scene_and_get_prompt(img)
                                        send_text = prompt_to_send # detect_scene_and_get_promptの実装に合わせて調整
                                    # ---------------------------------------------
                                    
                                    # 画像をペースト (OS共通のCtrl+Vをブラウザに直接送る)
                                    ActionChains(self.browser.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                                    time.sleep(1)
                                    
                                    # プロンプトテキストを準備してクリップボードへコピー
                                    if self.user_message:
                                        # 手動で質問された時も、メリハリお喋りルールを適用！
                                        prompt_text = (
                                            f"マスターからの質問：「{self.user_message}」\n"
                                            "以下の構成で【最高に明るい笑顔のトーン】に答えてね！♪\n"
                                            "1行目： [喜][驚] などの感情タグから始め、明るい第一声！語尾に『♪』を付けてね！\n"
                                            "2行目以降： 質問への丁寧な回答（文字数制限なし）"
                                        )
                                        self.user_message = ""
                                    else:
                                        # 自動実況の時は、detect_scene で作った特製プロンプトを使うよ！
                                        prompt_text = send_text
                                        
                                    pyperclip.copy(prompt_text)
                                    
                                    # テキストをペースト
                                    ActionChains(self.browser.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                                    time.sleep(1)
                                    
                                    # --- 送信前のAIの返答数をカウントしておく ---
                                    prev_count = self.browser.get_response_count()
                                    
                                    # Enterで送信！（Geminiの仕様でShift+Enterじゃなくて単なるEnterで飛ぶはず！）
                                    self.status_update.emit("🚀 送信したよ！お返事待ち…")
                                    print(f"[{datetime.now().strftime('%H:%M:%S')}] 🚀 ブラウザ経由でリクエストを送信したよ！")
                                    ActionChains(self.browser.driver).send_keys(Keys.ENTER).perform()
                                    
                                else:
                                    print("💦 フォーカス用の入力欄がどうしても見つからなかったよ！誤爆を防ぐため今回はスキップします！")
                                    continue # 誤爆防止のため、スキップして次のループへ
                                    
                            except Exception as e:
                                print(f"入力欄のフォーカス処理でエラー: {e}")
                                continue # エラー時も誤爆防止のためループをスキップ
                            
                            # 🤖 返答のストリーミング監視（スマート待機）
                            self.status_update.emit("🤖 考え中だよ…")
                            
                            # 1. まずは「新しい返答の枠」が作られるのを待つ（サーバー混雑時は長くなるので最大60秒に変更！）
                            new_response_started = False
                            for _ in range(60):
                                if not self.running: break
                                if self.browser.get_response_count() > prev_count:
                                    new_response_started = True
                                    break
                                time.sleep(1)
                                
                            if new_response_started:
                                # 2. 返答の文字が「ピタッと止まる（完成する）」まで監視する（最大120秒）
                                last_text = ""
                                unchanged_count = 0
                                
                                for _ in range(120):
                                    if not self.running: break
                                    
                                    current_text = self.browser.get_latest_response_text()
                                    
                                    if current_text:
                                        if current_text == last_text:
                                            # 文字が変わりない時間をカウント
                                            unchanged_count += 1
                                        else:
                                            # 文字が増えた！（まだ書いてる）
                                            unchanged_count = 0
                                            last_text = current_text
                                            self.status_update.emit(f"🤖 カキカキ中… ({len(current_text)}文字)")
                                            
                                            # 🏎️ 【爆速化】1行目が書き終わったら、全文完成を待たずに声を出し始めるよ！
                                            if '\n' in current_text and not getattr(self, "line_one_voiced", False):
                                                first_part = current_text.split('\n')[0].strip()
                                                tag_match = re.search(r'\[(.*?)\]', first_part)
                                                if tag_match:
                                                    print(f"[{datetime.now().strftime('%H:%M:%S')}] 🏎️ 1行目ゲット！先行して音声生成を開始するよ！")
                                                    self.status_update.emit("🎙️ 叫びを生成中…")
                                                    # タグを消して叫ぶ！
                                                    clean_shout = re.sub(r'\[.*?\]', '', first_part).strip()
                                                    if clean_shout:
                                                        self.generate_voice(clean_shout)
                                                    self.line_one_voiced = True # 二重生成防止
                                        
                                        # 3秒連続で文字が増えなければ完成！（ただし空っぽは除く）
                                        if unchanged_count >= 3 and len(current_text) > 0:
                                            # フラグリセット
                                            self.line_one_voiced = False
                                            
                                            # 前と全く同じセリフなら言わないようにする（ChatGPT提案の連投防止！）
                                            if current_text == self.last_comment:
                                                print(f"[{datetime.now().strftime('%H:%M:%S')}] 💦 同じセリフなのでスキップ！")
                                                self.status_update.emit("（前と同じセリフなので考え直すね…）")
                                            else:
                                                print(f"[{datetime.now().strftime('%H:%M:%S')}] 💬 取得した返信: {current_text[:30]}...")
                                                self.new_message.emit(current_text)
                                                self.status_update.emit("✨ お返事ゲット！")
                                                self.last_comment = current_text
                                                
                                                # 🧠 記憶に刻む！
                                                current_scene_name = "マスターへの回答" if self.user_message else "自動実況"
                                                self.save_history(current_scene_name, prompt_text[:100], current_text)
                                                
                                                # (先行TTSで処理済みなのでここでは何もしないよ！)
                                                # ------------------------------------
                                                
                                            if is_auto_mode:
                                                self.last_successful_auto_send = time.time() # 送信成功時刻を記録
                                            break
                                    time.sleep(1)
                            else:
                                print(f"[{datetime.now().strftime('%H:%M:%S')}] 💦 レスポンスの枠が新しく作られなかったよ…")
                                self.status_update.emit("💦 お返事が来なかったみたい…")
                                
                        else:
                            self.status_update.emit("💦 キャプチャ失敗しちゃった…")
                else:
                    self.status_update.emit("🥺 ウイポが起動してないみたい…")
            except Exception as e:
                 self.status_update.emit(f"💦 エラー: {str(e)}")
                 print(f"Unexpected Error: {e}")

            # 安定動作のための少しの待機（429エラー時は長めにペナルティ待機）
            if self.running and error_429_occurred:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] ⏳ 429エラーが発生したため、60秒待機します…")
                for _ in range(60):
                    if not self.running: break
                    time.sleep(1)

    def stop(self):
        self.running = False
        if self.browser:
            self.browser.close()


# ==========================================
# 🎀 可愛くて透明な秘書風オーバーレイUI
# ==========================================
class MessageWindow(QWidget):
    """ウイポのメッセージウィンドウ風のおしゃれなパネルだよ！"""
    
    # シグナルたち
    close_clicked = pyqtSignal()
    chat_sent = pyqtSignal(str)
    ghost_toggled = pyqtSignal(bool) # 透過モードが切り替わったことを親に知らせるシグナル

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        # 👑 右上の閉じるボタン（×）とステータスのレイアウト
        # 💡 ボタン位置が変わらないように、左右を分けたダミー配置にするよ！
        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)
        
        # 左側：ステータスエリア（ここが消えても右側をいじらせない！）
        self.status_label = QLabel("起動中…")
        self.status_label.setStyleSheet("color: white; background-color: rgba(0,0,0,150); border-radius: 5px; padding: 4px; font-weight:bold;")
        self.status_label.setFont(QFont("Meiryo", 10))
        top_layout.addWidget(self.status_label)
        
        top_layout.addStretch() # 真ん中を広げる
        
        # 右側：操作ボタンエリア（ここを固定幅にしてズレを防ぐ！）
        self.button_container = QWidget()
        self.button_container.setFixedWidth(130) # 90(透過btn) + 30(×btn) + 余白
        btn_layout = QHBoxLayout(self.button_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(5)
        
        # 👻 クリック透過を切り替える魔法のボタン
        self.ghost_btn = QPushButton("👻 透過(F12)")
        self.ghost_btn.setFixedSize(90, 30)
        self.ghost_btn.setStyleSheet("""
            QPushButton {
                background-color: rgba(100, 149, 237, 200); /* コーンフラワーブルー */
                color: white;
                border-radius: 15px;
                font-weight: bold;
                font-family: 'Meiryo';
            }
            QPushButton:hover {
                background-color: rgba(65, 105, 225, 255); /* ロイヤルブルー */
            }
        """)
        self.ghost_btn.clicked.connect(self.toggle_ghost_mode)
        self.is_ghost_mode = False # 現在の透過状態
        
        self.close_btn = QPushButton("✖")
        self.close_btn.setFixedSize(30, 30)
        self.close_btn.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 105, 180, 200);
                color: white;
                border-radius: 15px;
                font-weight: bold;
                font-family: 'Arial';
            }
            QPushButton:hover {
                background-color: rgba(255, 20, 147, 255);
            }
        """)
        self.close_btn.clicked.connect(self.close_clicked.emit)
        
        btn_layout.addWidget(self.ghost_btn)
        btn_layout.addWidget(self.close_btn)
        
        top_layout.addWidget(self.button_container)

        # クリックできなくなった時のために、F12キーでどこからでも戻せるように裏で監視するタイマー
        self.hotkey_timer = QTimer(self)
        self.hotkey_timer.timeout.connect(self.check_hotkey)
        self.hotkey_timer.start(100) # 0.1秒ごとにF12キーが押されたかチェック
        
        main_layout.addLayout(top_layout)

        # 💬 あいちゃんのセリフ部分（QTextBrowserで超軽量化！）
        self.chat_log = QTextBrowser()
        self.chat_log.setOpenExternalLinks(True)
        self.chat_log.setReadOnly(True)
        # メモリ肥大化を防ぐためのリミッターと、余計な機能の無効化（ChatGPT提案の安定化！）
        self.chat_log.setUndoRedoEnabled(False)
        self.chat_log.document().setMaximumBlockCount(500)
        self.chat_log.document().setDocumentMargin(2)
        # スクロールバーのデザインカスタマイズ（QTextBrowserの内蔵ScrollBarに適用）
        self.chat_log.verticalScrollBar().setStyleSheet("""
            QScrollBar:vertical {
                border: none;
                background: rgba(255, 182, 193, 50);
                width: 10px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: rgba(255, 105, 180, 150);
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: rgba(255, 20, 147, 200);
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)

        # QTextBrowser自体のスタイル（背景透明、枠線なし、文字設定）
        self.chat_log.setStyleSheet("""
            QTextBrowser {
                color: #4a4a4a;
                font-size: 18px;
                font-weight: bold;
                font-family: 'Meiryo', 'Yu Gothic', sans-serif;
                padding: 10px;
                background: transparent;
                border: none;
            }
        """)
        
        # 最初の挨拶をセット
        self.chat_log.setHtml("マスター、お疲れ様！<br>今はウイポの画面を探しているよー！")
        
        main_layout.addWidget(self.chat_log, 1) # 1で伸縮可能に
        
        # ⌨️ マスターからのチャット入力欄
        input_layout = QHBoxLayout()
        self.chat_input = QLineEdit()
        self.chat_input.setPlaceholderText("あいちゃんに話しかける... (Enterで送信)")
        self.chat_input.setStyleSheet("""
            QLineEdit {
                background-color: rgba(255, 255, 255, 200);
                border: 2px solid #ffb6c1;
                border-radius: 10px;
                padding: 5px 10px;
                font-size: 16px;
                font-family: 'Meiryo', 'Yu Gothic', sans-serif;
                color: #333333;
            }
            QLineEdit:focus {
                border: 2px solid #ff69b4;
                background-color: rgba(255, 255, 255, 255);
            }
        """)
        self.chat_input.returnPressed.connect(self.send_chat)
        
        self.send_btn = QPushButton("送信💌")
        self.send_btn.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 182, 193, 220);
                border: 2px solid #ff69b4;
                border-radius: 10px;
                padding: 5px 15px;
                font-size: 14px;
                font-weight: bold;
                color: #fff;
            }
            QPushButton:hover {
                background-color: rgba(255, 105, 180, 255);
            }
        """)
        self.send_btn.clicked.connect(self.send_chat)
        
        input_layout.addWidget(self.chat_input)
        input_layout.addWidget(self.send_btn)
        
        main_layout.addLayout(input_layout)
        
        self.setLayout(main_layout)
        
        # ウィンドウの横サイズを1.5〜2倍に拡大し、リサイズ可能にする
        self.setMinimumWidth(800)
        # 最大サイズの制限を外す
        self.setMaximumWidth(1600)
        
        # 内側のラベルが枠いっぱいに広がるようにポリシーを設定
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def setText(self, text):
        # 毎回全テキストを作り直すのをやめて、軽量な「追加(append)」に変更！
        now_time = datetime.now().strftime('%H:%M')
        
        # 改行をHTMLの <br> に変換
        formatted_new_text = text.replace('\n', '<br>')
        
        # 初回じゃない場合はセパレーターを入れる
        # QTextBrowserならdivのCSSが素直に解釈される！
        time_tag = f"<div style='color:#999;font-size:12px'>{now_time}</div>"
        line_tag = "<hr style='border:none; border-top:2px solid #FFC0CB; margin:10px 0;'>"
        append_html = f"{time_tag}{line_tag}{formatted_new_text}<br><br>"
            
        # スクロールバーが一番下にあるかチェック（ChatGPT提案のスマートスクロール！）
        scroll = self.chat_log.verticalScrollBar()
        at_bottom = scroll.value() >= scroll.maximum() - 4
        
        # 画面の一番下に追加！（余計な<p>タグの自動付与を避ける作戦！）
        cursor = self.chat_log.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.chat_log.setTextCursor(cursor)
        
        self.chat_log.insertHtml(append_html)
        
        # もともと一番下を見ていた場合だけ、追従してスクロールさせる
        if at_bottom:
            scroll.setValue(scroll.maximum())
            
    def toggle_ghost_mode(self):
        """👻 クリック透過モードを切り替えるよ！"""
        self.is_ghost_mode = not self.is_ghost_mode
        
        # 親ウィンドウ（SecretaryOverlay）にも知らせて、アバターなどを消してもらう！
        self.ghost_toggled.emit(self.is_ghost_mode)
        
        # 必要なパーツの表示・非表示を切り替え！
        visible = not self.is_ghost_mode
        self.chat_log.setVisible(visible)
        self.chat_input.setVisible(visible)
        self.send_btn.setVisible(visible)
        self.status_label.setVisible(visible)
        # ✖ボタンは透過中も位置が変わらないように「非表示」ではなく「無効化」か、そのままでもOK
        # 今回は「同じ位置に」という要望なので、close_btn は消さずに無効化するよ！
        self.close_btn.setEnabled(visible)
        self.set_close_btn_opacity(0.3 if self.is_ghost_mode else 1.0) # 透明度だけ変える

        if self.is_ghost_mode:
            # 💡 WindowTransparentForInput フラグを使うとボタンまでクリックできなくなるので、
            # WA_TranslucentBackground の「透明な部分はクリックが抜ける」性質を利用するよ！
            self.ghost_btn.setText("👆 復帰(F12)")
            self.ghost_btn.setStyleSheet("""
                QPushButton {
                    background-color: rgba(255, 140, 0, 220); /* ダークオレンジ (クリックしやすいように不透明度高め) */
                    color: white;
                    border-radius: 15px;
                    font-weight: bold;
                    font-family: 'Meiryo';
                }
                QPushButton:hover {
                    background-color: rgba(255, 69, 0, 255);
                }
            """)
            print("👻 透過モードON！背景が消えて、裏のゲームが操作できるようになりました！")
        else:
            self.ghost_btn.setText("👻 透過(F12)")
            self.ghost_btn.setStyleSheet("""
                QPushButton {
                    background-color: rgba(100, 149, 237, 200); /* コーンフラワーブルー */
                    color: white;
                    border-radius: 15px;
                    font-weight: bold;
                    font-family: 'Meiryo';
                }
                QPushButton:hover {
                    background-color: rgba(65, 105, 225, 255);
                }
            """)
            print("👆 透過モードOFF！チャット画面が復活しました！")
            
        # 背景の描画を更新させる
        self.update()

    # ボタンの透明度を操作するための便利なメソッドを追加
    def set_close_btn_opacity(self, opacity):
        self.close_btn.setGraphicsEffect(None) # リセット
        if opacity < 1.0:
            op = QGraphicsOpacityEffect(self.close_btn)
            op.setOpacity(opacity)
            self.close_btn.setGraphicsEffect(op)

    def check_hotkey(self):
        """裏でF12キーが押されたか監視して、いつでも透過をトグルできるようにするよ！"""
        import win32api
        import win32con
        # F12キーの状態を取得（最上位ビットが1なら現在押されている）
        state = win32api.GetAsyncKeyState(win32con.VK_F12)
        if state & 0x8000:
            if not getattr(self, "hotkey_pressed", False):
                self.hotkey_pressed = True
                self.toggle_ghost_mode() # トグル実行！
        else:
            self.hotkey_pressed = False

    # 上の toggle_ghost_mode を修正： close_btn の透明度設定用
    def update_ghost_ui(self):
        # 既に toggle_ghost_mode で処理しているので、ここでは特になし
        pass
        
    def auto_scroll(self, min_val, max_val):
        # 昔のQLabel時代の名残なのでパス（上の処理でスクロールできるため）
        pass
        
    def setStatus(self, text):
        self.status_label.setText(text)
        
    def send_chat(self):
        text = self.chat_input.text().strip()
        if text:
            self.chat_sent.emit(text)
            self.chat_input.clear()
            self.chat_input.setPlaceholderText("送信したよ！お返事を待っててね…")

    def paintEvent(self, event):
        """ここで可愛いフリル風のデザインを描くよ！🎨"""
        # 👻 透過モード中は背景を「完全透明」にして、クリックをすり抜けさせる！
        if hasattr(self, 'is_ghost_mode') and self.is_ghost_mode:
            return
            
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        rect = self.rect()
        margin = 3
        draw_rect = rect.adjusted(margin, margin, -margin, -margin)
        
        from PyQt6.QtCore import QRectF
        draw_rect_f = QRectF(draw_rect)
        
        path = QPainterPath()
        path.addRoundedRect(draw_rect_f, 20.0, 20.0)
        
        painter.fillPath(path, QColor(255, 255, 255, 220)) # 薄い白（半透明）
        
        pen = painter.pen()
        pen.setColor(QColor(255, 182, 193)) # ライトピンク
        pen.setWidth(4)
        painter.setPen(pen)
        painter.drawPath(path)
        
        inner_rect = draw_rect.adjusted(6, 6, -6, -6)
        inner_rect_f = QRectF(inner_rect)
        inner_path = QPainterPath()
        inner_path.addRoundedRect(inner_rect_f, 14.0, 14.0)
        pen.setWidth(1)
        pen.setColor(QColor(255, 105, 180, 150)) # 少し濃いピンク（半透明）
        painter.setPen(pen)
        painter.drawPath(inner_path)


class SecretaryOverlay(QWidget):
    def __init__(self):
        super().__init__()
        
        # 背景透明・枠なし・最前面 (Topmost)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint | 
            Qt.WindowType.WindowStaysOnTopHint | 
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        self.drag_position = None
        
        self.init_ui()
        
        # キャプチャスレッドの準備
        self.capture_thread = CaptureThread()
        self.capture_thread.new_message.connect(self.update_speech)
        self.capture_thread.status_update.connect(self.update_status)
        self.capture_thread.start()

    def init_ui(self):
        # 立ち絵 ＋ メッセージウィンドウの横並びレイアウト
        layout = QHBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(-15) # 💡 立ち絵と吹き出しを少し重ねて一体感を出すよ！
        self.setLayout(layout)
        
        # 高さを揃えるための基準サイズ（マスターの要望でより高く・見やすく）
        target_height = 360
        
        # 1. 立ち絵画像（ai_secretary.png を使ってバストアップに切り抜く！）
        self.avatar_label = QLabel()
        self.crop_rect = None
        avatar_path = "ai_secretary.png"
        
        if os.path.exists(avatar_path):
            original_pixmap = QPixmap(avatar_path)
            
            # 💡 もっと顔を大きく！真のバストアップ（腰から上）だけを切り抜くよ！
            width = original_pixmap.width()
            height = original_pixmap.height()
            
            # 上の余白（頭の上）を見切れないように少なくカットして、顔〜胸下あたりを狙い撃ち！
            crop_y = int(height * 0.01) # 頭が見切れないようにほぼ一番上から！
            crop_h = int(height * 0.40) # 顔から腰までの範囲
            self.crop_rect = original_pixmap.copy(0, crop_y, width, crop_h)
            
            # 吹き出しと同じ初期高さにスケールを合わせる
            final_pixmap = self.crop_rect.scaledToHeight(target_height, Qt.TransformationMode.SmoothTransformation)
            self.avatar_label.setPixmap(final_pixmap)
        else:
            self.avatar_label.setText(" 🥺\n画像がないよ\nai_secretary.png\nを保存してね！")
            self.avatar_label.setStyleSheet("background-color: rgba(255,255,255,200); padding: 20px; border-radius: 10px;")
            
        self.avatar_label.setAlignment(Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)
        self.avatar_label.setMinimumHeight(target_height)
        self.avatar_label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        
        layout.addWidget(self.avatar_label, alignment=Qt.AlignmentFlag.AlignBottom)
        
        # 2. メッセージウィンドウ（チャット欄・閉じるボタン入り）
        self.speech_bubble = MessageWindow()
        self.speech_bubble.setMinimumHeight(target_height) # 💡 固定高さを解除してリサイズを許可
        
        # シグナルの接続
        self.speech_bubble.close_clicked.connect(self.close) # ×ボタンで閉じる
        self.speech_bubble.chat_sent.connect(self.handle_chat) # チャット送信時
        self.speech_bubble.ghost_toggled.connect(self.handle_ghost_mode) # 透過切り替え時
        
        layout.addWidget(self.speech_bubble) # alignmentを外して高さを自由に広げる
        
        # 3. サイズ変更用のグリップ（右下）
        self.size_grip = QSizeGrip(self)
        self.size_grip.setFixedSize(20, 20)
        self.size_grip.setStyleSheet("background-color: rgba(255, 105, 180, 200); border-radius: 10px;")
        self.size_grip.setToolTip("ここを引っぱってサイズを変えてね！")
        
        grip_layout = QVBoxLayout()
        grip_layout.addStretch()
        grip_layout.addWidget(self.size_grip, 0, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)
        layout.addLayout(grip_layout)
        
        # 画面の右下に配置するよ！
        screen = QApplication.primaryScreen().geometry()
        
        # ウィンドウの初期サイズを計算
        self.adjustSize()
        total_width = self.width() + 100 # 少し余裕を持たせる
        
        x_pos = screen.width() - total_width - 80 # 右端から少し離す
        y_pos = screen.height() - target_height - 100 # 下から浮かせる
        self.move(x_pos, y_pos)

    def handle_ghost_mode(self, is_ghost):
        """透過モードに合わせてアバターとサイズグリップの表示を切り替えるよ！"""
        if is_ghost:
            self.avatar_label.hide()
            self.size_grip.hide()
        else:
            self.avatar_label.show()
            self.size_grip.show()

    # 🖱️ マウスでドラッグして動かせる魔法！
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            # 入力欄などをクリックした時はドラッグ判定しないようにする工夫
            child = self.childAt(event.position().toPoint())
            if isinstance(child, QLineEdit) or isinstance(child, QPushButton):
                return
            self.drag_position = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self.drag_position is not None:
            self.move(event.globalPosition().toPoint() - self.drag_position)
            event.accept()
            
    def mouseReleaseEvent(self, event):
        self.drag_position = None

    def resizeEvent(self, event):
        """ウィンドウサイズが変更されたら、立ち絵の高さも動的に追従させる！"""
        super().resizeEvent(event)
        if hasattr(self, 'crop_rect') and self.crop_rect is not None:
            # 吹き出しの現在の実高さに合わせてリサイズ
            new_height = self.speech_bubble.height()
            if new_height > 10:
                self.avatar_label.setPixmap(self.crop_rect.scaledToHeight(new_height, Qt.TransformationMode.SmoothTransformation))

    def handle_chat(self, text):
        """マスターからのチャットを裏方スレッドに渡すよ！"""
        self.update_speech(f"💬 マスター: {text}\n(考え中だよ…✨)")
        self.capture_thread.send_chat(text)

    def update_speech(self, text):
        self.speech_bubble.setText(text)
        
    def update_status(self, text):
        self.speech_bubble.setStatus(text)
        print(f"👉 状況: {text}")

    def closeEvent(self, event):
        print("👋 アプリを終了するね！お疲れ様でした！")
        self.capture_thread.stop()
        # API通信などでスレッドが止まっている場合の対策として、1秒待ってダメなら強制終了！
        if not self.capture_thread.wait(1000):
            self.capture_thread.terminate()
        event.accept()
        QApplication.instance().quit() # ターミナルの入力待ちに返すため、完全にアプリを終わらせるよ！

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    print("🎀 API完全不要・超自動ブラウザ実況システムを起動するよ！")
    print("💡 Edgeブラウザが自動で立ち上がるのを確認してね！")
    print("💡 ウィンドウはマウスでドラッグして好きな場所に移動できるよ！")
    print("💡 終了したいときは、右上の「✖」ボタンを押してね！")
    
    overlay = SecretaryOverlay()
    overlay.show()
    
    try:
        sys.exit(app.exec())
    except KeyboardInterrupt:
        pass
