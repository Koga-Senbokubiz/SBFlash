# -*- coding: utf-8 -*-
"""
SBFlashFunctions.py (v0.21)

SBFlash Pro / Lite の機能差分を外だしするための定義ファイル。
False にした場合は、その機能をまとめて無効にする。
- ボタンを表示しない
- 実行できない
- ショートカットも効かない
"""

# アプリ種別
PRODUCT_NAME = "SBFlashPro"
PRODUCT_EDITION = "Pro"

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
