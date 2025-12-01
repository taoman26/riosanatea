#!/usr/bin/python3
# coding:utf-8

"""
CSV処理とファイル入出力ユーティリティモジュール
"""

import csv
import codecs
import subprocess
from io import BytesIO


def csv_to_list(csv_path):
    """何種類かの文字コードに対応させた、CSV読み込みリスト化関数"""
    code_list = ["euc_jp", "shift_jis", "cp932", "iso2022_jp", "utf-8"]
    csv_code = ""
    csv_description_list = []
    temporary_lines = ""
    
    # まず、ファイルの文字コードを調べる
    for code in code_list:
        try:
            # Python3以降だと、2までと違って読み込みの段階で弾かれるようなので、読み込みをtryにかける
            f1 = codecs.open(csv_path, "r", code)
            for line in f1:
                temporary_lines += line
        except:
            continue
        
        csv_code = code
        break
    
    # 調べた文字コードを元に、読み込んでリスト化していく
    f2 = open(csv_path, "r", encoding=csv_code)
    reader = csv.reader(f2, quotechar='"')
    
    # readerはイテレータで内容を直接表さないのでリストに格納しなおす
    for row in reader:
        csv_description_list.append(row)
    
    f2.close()
    return csv_description_list


def pil_printing(pil_image, paper_size="", upside_down=False):
    """
    画像をlpに渡して印刷する
    Ubuntu 22.04 / 24.04 対応
    """
    
    # 上下反転印刷モードであれば。宛名の画像を180°回転させる
    if upside_down is True:
        pil_image = pil_image.rotate(180)
    
    # 画像をIOオブジェクトに変換して印刷する。
    # Ubuntu 22.04と24.04の両方で動作するようにPNG形式を使用。
    # （Ubuntu 24.04でPNM形式が非対応になったため、より広くサポートされるPNG形式を使用）
    io_object = BytesIO()
    pil_image.save(io_object, "PNG")
    
    # 用紙サイズの指定がなければ、「lpoptions -l」を実行して使用可能なものを検出する
    if paper_size == "":
        lpoptions_stdout = ""
        paper_size = "A6"
        # ↑プリンターがCustom.WIDTHxHEIGHTやPostcardに対応していなくても
        # A6での印刷は可能と想定して（検出できなかった場合に使う）初期値にしておく
        
        try:
            subproc = subprocess.Popen(["lpoptions", "-l"], stdout=subprocess.PIPE)
            lpoptions_stdout = [x.decode("utf-8") for x in subproc.stdout.readlines()]
        except:
            paper_size = "A6"
        
        for line in lpoptions_stdout:
            if line.startswith("PageSize") is True:
                if "Custom.WIDTHxHEIGHT" in line:
                    paper_size = "Custom.100x148mm"
                elif "Postcard" in line:
                    paper_size = "Postcard"
    
    try:
        p = subprocess.Popen(['lp', '-o', 'PageSize=' + paper_size], 
                            stdin=subprocess.PIPE, 
                            stderr=subprocess.PIPE)
        stdout, stderr = p.communicate(input=io_object.getvalue())
        if p.returncode != 0:
            print("印刷エラー: " + stderr.decode("utf-8"))
    except Exception as e:
        print("印刷処理でエラーが発生しました: " + str(e))
