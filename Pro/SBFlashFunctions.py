# -*- coding: utf-8 -*-
"""
SBFlashFunctions.py (Ver1.00)

SBFlash Pro / Lite の機能差分と、アプリ共通設定をまとめて外だしする定義ファイル。
ini は「利用者が実行環境ごとに変えたい値の上書き先」として残しつつ、
標準値・製品設定・版数はこのファイル側に寄せる。
"""

# =========================================================
# 製品情報 / リリース情報
# =========================================================
PRODUCT_NAME = "SBFlashPro"
PRODUCT_EDITION = "Pro"
APP_TITLE = "SBFlash Pro"
APP_VERSION = "Ver1.00"
DEFAULT_INI_FILENAME = "SBFlashPro.ini"

# =========================================================
# データ / 起動時既定値
# =========================================================
DEFAULT_EXCEL_PATH = "FlashCards.xlsx"
INITIAL_SHEET = "sheet0"
WRONG_SHEET = "回答シート"

DATA_START_ROW_DEFAULT = 2
WRONG_START_ROW = 3

WRONG_ONLY = False
WORST_FIRST = False
ALL_SUBJECTS = False

# =========================================================
# UI既定値
# =========================================================
WINDOW_WIDTH = 0
WINDOW_HEIGHT = 0
START_MAXIMIZED = False

AUTO_RATIO = 0.90
MIN_WIDTH = 820
MIN_HEIGHT = 620
MAX_WIDTH = 2200
MAX_HEIGHT = 1400

THUMB_SIZE = 400
ZOOM_MAX = 1200

REVERSE_LABEL_NORMAL = "解答⇔設問"
REVERSE_LABEL_REVERSED = "設問⇔解答"

# =========================================================
# 機能制御
# False の場合は、表示・実行・ショートカットをすべて無効化
# =========================================================
USE_SELF_GRADE_CORRECT = True
USE_SELF_GRADE_INCORRECT = True

USE_OX_BUTTONS = True
USE_SAVE_ANSWER = False
USE_BOOKMARK_TOGGLE = True
USE_BOOKMARK_CLEAR_ALL = True
USE_BOOKMARK_NEXT = True
USE_TOPIC_REVIEW = True
