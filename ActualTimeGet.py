import datetime
import pyautogui as pgui
import time
# from ast import IsNot
# from operator import is_not
# import pyperclip as clip
# import pandas as pd
# import subprocess
# import sys
# from tkinter import messagebox


# ------------------------------------------------------
# 関数名     get_first_date
# 用途       月初を取得
# 引数       dt = 日付
# ------------------------------------------------------
def get_first_date(dt):
    return dt.replace(day=1)

FirstDate = get_first_date(datetime.date.today())
strFirstDate = FirstDate.strftime("%Y/%m/%d")

# 製造実績集計ツールを開く
pgui.press("win")

# 製造実績集計ツールが開くまで待つ
while pgui.locateOnScreen(r".\img\SeizoTool.png" , confidence=0.9) is None:
    time.sleep(1)

# 現在位置を取得
loc = pgui.position()
offset=(0, 0)

# 製造実績集計ツールボタンの位置を特定
img_x, img_y = pgui.locateCenterOnScreen(r".\img\SeizoTool.png", grayscale=True, confidence=0.9)

# オフセット分を計算
img_x = img_x + offset[0]
img_y = img_y + offset[1]

# 対象位置へ移動
pgui.moveTo(img_x,img_y,1)
pgui.click(img_x,img_y)


# 製造実績集計ツールが開くまで待つ
while pgui.locateOnScreen(r".\img\OutButton.png" , confidence=0.9) is None:
    time.sleep(1)

# 現在位置を取得
loc = pgui.position()
offset=(0, 0)

# 製造実績集計ツールボタンの位置を特定
img_x, img_y = pgui.locateCenterOnScreen(r".\img\OutButton.png", grayscale=True, confidence=0.9)

# オフセット分を計算
img_x = img_x + offset[0]
img_y = img_y + offset[1]

# 対象位置へ移動
pgui.moveTo(img_x,img_y,1)
pgui.click(img_x,img_y)

# 製造実績集計ツールが開くまで待つ
while pgui.locateOnScreen(r".\img\SeizoToo2.png" , confidence=0.9) is None:
    time.sleep(1)

# 開始日付
pgui.press("tab",3,interval=0.3)
pgui.typewrite(strFirstDate)

# 作業部門
pgui.press("tab",5,interval=0.3)
pgui.typewrite("GMMS11")

# 完了
pgui.press("tab",4,interval=0.3)
time.sleep(1)
pgui.press("space")

# 製造実績集計ツールボタンの位置を特定
img_x, img_y = pgui.locateCenterOnScreen(r".\img\StartButton.png", grayscale=True, confidence=0.9)

# オフセット分を計算
img_x = img_x + offset[0]
img_y = img_y + offset[1]

# 対象位置へ移動
pgui.moveTo(img_x,img_y,1)
pgui.click(img_x,img_y)

# 製造実績集計ツールが開くまで待つ
while pgui.locateOnScreen(r".\img\Executing.png" , confidence=0.9) is not None:
    time.sleep(1)

# 製造実績集計ツールボタンの位置を特定
img_x, img_y = pgui.locateCenterOnScreen(r".\img\ExcelButton.png", grayscale=True, confidence=0.9)

# オフセット分を計算
img_x = img_x + offset[0]
img_y = img_y + offset[1]

# 対象位置へ移動
pgui.moveTo(img_x,img_y,1)
pgui.click(img_x,img_y)

# 製造実績集計ツールが開くまで待つ
while pgui.locateOnScreen(r".\img\PathIn.png" , confidence=0.9) is None:
    time.sleep(1)

# 製造実績集計ツールボタンの位置を特定
img_x, img_y = pgui.locateCenterOnScreen(r".\img\PathIn.png", grayscale=True, confidence=0.9)

# オフセット分を計算
img_x = img_x + offset[0]
img_y = img_y + offset[1]

# 対象位置へ移動
pgui.moveTo(img_x,img_y,1)
pgui.click(img_x,img_y)

pgui.typewrite("\\fks-file-005\Disk22\005\17_一般管理業務\機種別（旧MC）★\MS2830A\その他\Okumura\113_TMI\2021_H2\2月")
time.sleep(1)
pgui.press("enter")
time.sleep(1)
pgui.hotkey("alt","s")