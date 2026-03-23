# -*- coding: utf-8 -*-
"""
SBFlashFunctions.py (Ver1.00)

SBFlash Pro / Lite の機能制御と、iniファイル関連の定義をまとめる。
データ / 起動時既定値、および UI既定値は ini ファイル側で管理する。
"""

# アプリ種別
PRODUCT_NAME = "SBFlashPro"
PRODUCT_EDITION = "Pro"

# iniファイル設定
DEFAULT_INI_FILENAME = "SBFlashPro.ini"

# -------------------------------------
# 自己採点機能の一括制御
# False の場合は、表示・実行・ショートカットをすべて無効化
# -------------------------------------
USE_SELF_GRADE_CORRECT = True
USE_SELF_GRADE_INCORRECT = True

# -------------------------------------
# Lite向け外だし対象
# False の場合は、表示・実行・ショートカットをすべて無効化
# -------------------------------------
USE_OX_BUTTONS = True
USE_SAVE_ANSWER = False
USE_BOOKMARK_TOGGLE = True
USE_BOOKMARK_CLEAR_ALL = True
USE_BOOKMARK_NEXT = True
USE_TOPIC_REVIEW = True
