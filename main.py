#!/usr/bin/python3
# coding:utf-8

"""
Riosanatea - はがき宛名印刷ソフト
メインエントリーポイント

使い方:
1. ライブラリのインストール:
   pip install -r requirements.txt

2. 実行:
   python3 main.py
   または
   python3 main.py <CSVファイルのパス>

3. モジュール構成:
   - main.py: このファイル（エントリーポイントとGUI）
   - image_utils.py: 画像処理ユーティリティ
   - csv_utils.py: CSV処理と印刷機能
   - requirements.txt: 必要なライブラリ一覧

互換性:
   Ubuntu 22.04 / 24.04 対応
"""

# Riosanatea.pyを読み込む（メインの処理は元のファイルに残す）
# 大規模な分割を避け、段階的にリファクタリングできるようにする
from Riosanatea import *

if __name__ == "__main__":
    import sys
    import os
    import wx
    from csv_utils import csv_to_list
    
    # 引数の処理
    csvfile_path = ""
    for com_line_arg in sys.argv:
        # 引数の拡張子がcsvで、ファイルとして存在していればパスを登録する
        # 起動時に開くのは1つだけなので、複数来てもパスを上書きしていくことになる
        if os.path.splitext(com_line_arg)[1].lower() == ".csv" and os.path.isfile(com_line_arg):
            csvfile_path = os.path.abspath(com_line_arg)
    
    application = wx.App()
    frame = frame_plus(None, wx.ID_ANY, "宛名印刷ソフト Riosanatea")
    
    # CSVファイルが起動時に引数から渡されていれば、開く
    if csvfile_path != "":
        frame.open_csv_file(csvfile_path)
    
    frame.Centre()
    frame.Show()
    
    application.MainLoop()
