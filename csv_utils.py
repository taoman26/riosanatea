#!/usr/bin/python3
# coding:utf-8

"""
CSV処理とファイル入出力ユーティリティモジュール
"""

import csv
import codecs
import subprocess
import tempfile
import os
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm


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
    画像をPDFに変換してlpに渡して印刷する
    Ubuntu 22.04 / 24.04 対応
    
    PDFを使用することで、用紙サイズと画像サイズを正確に制御し、
    プリンタドライバによる自動スケーリングを防ぐ
    
    画像は8pixel/mmで生成されているため、画像サイズから用紙サイズを自動計算
    """
    
    # 上下反転印刷モードであれば。宛名の画像を180°回転させる
    if upside_down is True:
        pil_image = pil_image.rotate(180)
    
    # 画像サイズから用紙サイズを計算（8 pixel/mm）
    img_width_px, img_height_px = pil_image.size
    paper_width_mm = img_width_px / 8
    paper_height_mm = img_height_px / 8
    
    print(f"画像サイズ: {img_width_px}x{img_height_px} pixel")
    print(f"用紙サイズ: {paper_width_mm:.1f}x{paper_height_mm:.1f} mm")
    
    # 画像を一時ファイルに保存
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_img:
        img_path = tmp_img.name
        pil_image.save(img_path, format='PNG')
    
    # 一時ファイルにPDFを作成
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp_pdf:
        pdf_path = tmp_pdf.name
        
        # 実際の用紙サイズでPDFを作成
        pdf_canvas = canvas.Canvas(pdf_path, pagesize=(paper_width_mm*mm, paper_height_mm*mm))
        
        # PDFに画像を配置（用紙サイズぴったりに）
        # 座標は左下が原点、画像は用紙全体に配置
        pdf_canvas.drawImage(img_path, 0, 0, width=paper_width_mm*mm, height=paper_height_mm*mm, 
                            preserveAspectRatio=False)
        pdf_canvas.save()
    
    # デフォルトプリンタの確認
    default_printer = None
    try:
        result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            # "system default destination: プリンタ名" の形式から抽出
            output = result.stdout.strip()
            if ":" in output:
                default_printer = output.split(":")[-1].strip()
    except Exception as e:
        print(f"デフォルトプリンタの確認エラー: {e}")
    
    # 用紙サイズの指定がなければ、画像サイズに基づいてプリンタがサポートする最適なサイズを検出
    if paper_size == "":
        try:
            subproc = subprocess.Popen(["lpoptions", "-l"], stdout=subprocess.PIPE)
            lpoptions_stdout = [x.decode("utf-8") for x in subproc.stdout.readlines()]
        except:
            lpoptions_stdout = []
        
        # 画像サイズ（mm）を使って、プリンタがサポートする用紙サイズ名を検索
        # サイズ指定形式のパターン: "100x148", "w283h420" (Canon), "a4", "postcard" など
        width_int = int(round(paper_width_mm))
        height_int = int(round(paper_height_mm))
        
        # 検索パターン（優先順位順）
        search_patterns = [
            f"{width_int}x{height_int}",  # 直接サイズ指定: "100x148"
            f"w{width_int*283//100}h{height_int*420//148}",  # Canon形式（はがきの場合w283h420）
        ]
        
        # 一般的な名称も追加（サイズに応じて）
        if width_int == 100 and height_int == 148:
            search_patterns.extend(["hagaki", "postcard"])
        elif width_int == 105 and height_int == 148:
            search_patterns.append("a6")
        elif width_int == 148 and height_int == 210:
            search_patterns.append("a5")
        elif width_int == 210 and height_int == 297:
            search_patterns.append("a4")
        elif width_int == 297 and height_int == 420:
            search_patterns.append("a3")
        
        print(f"用紙サイズ検索パターン: {search_patterns}")
        
        # プリンタがサポートする用紙サイズを検索
        for pattern in search_patterns:
            for line in lpoptions_stdout:
                if line.startswith("PageSize") or line.startswith("media"):
                    if pattern.lower() in line.lower():
                        # 実際の用紙サイズ名を抽出
                        for word in line.split():
                            if pattern.lower() in word.lower() and not word.startswith("*"):
                                paper_size = word
                                break
                        if paper_size:
                            break
            if paper_size:
                break
        
        # 検出できない場合は、カスタムサイズまたは汎用デフォルト
        if not paper_size:
            # カスタムサイズを試す
            custom_size = f"Custom.{width_int}x{height_int}mm"
            for line in lpoptions_stdout:
                if "Custom" in line or "custom" in line.lower():
                    paper_size = custom_size
                    print(f"カスタムサイズを使用: {custom_size}")
                    break
            
            # それでもダメなら汎用デフォルト（サイズに応じて）
            if not paper_size:
                if paper_width_mm <= 110 and paper_height_mm <= 160:
                    paper_size = "Postcard"
                elif paper_width_mm <= 160 and paper_height_mm <= 230:
                    paper_size = "A5"
                else:
                    paper_size = "A4"
                print(f"警告: 用紙サイズを検出できませんでした。{paper_size}を使用します")
    
    print(f"使用する用紙サイズ設定: {paper_size}")
    
    try:
        # lpコマンドの構築 - PDFファイルを印刷
        # PDFはページサイズが埋め込まれているため、fit-to-pageで正確に印刷できる
        lp_cmd = ['lp', 
                  '-o', 'fit-to-page',              # PDFのページサイズに合わせる
                  '-o', 'media=' + paper_size,      # 用紙サイズ指定
                  pdf_path]                         # PDFファイルパス
        
        # デフォルトプリンタが設定されていない場合、利用可能なプリンタを確認
        if not default_printer:
            try:
                result = subprocess.run(['lpstat', '-p'], capture_output=True, text=True, timeout=5)
                if result.returncode == 0 and result.stdout.strip():
                    # 最初のプリンタを使用
                    first_line = result.stdout.strip().split('\n')[0]
                    if first_line.startswith('printer '):
                        printer_name = first_line.split()[1]
                        lp_cmd.extend(['-d', printer_name])
                        print(f"デフォルトプリンタが設定されていないため、'{printer_name}' を使用します")
                else:
                    raise Exception("プリンタが見つかりません。システム設定でプリンタを追加してください。")
            except subprocess.TimeoutExpired:
                raise Exception("プリンタの検索がタイムアウトしました")
        
        # PDFファイルを印刷
        p = subprocess.Popen(lp_cmd, 
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE)
        stdout, stderr = p.communicate()
        
        if p.returncode != 0:
            error_msg = stderr.decode("utf-8")
            if "No default destination" in error_msg:
                raise Exception("デフォルトプリンタが設定されていません。\n以下のコマンドでプリンタを確認・設定してください:\n  lpstat -p  (利用可能なプリンタ一覧)\n  lpoptions -d プリンタ名  (デフォルトプリンタの設定)")
            else:
                raise Exception(f"印刷エラー: {error_msg}")
    except Exception as e:
        raise Exception(f"印刷処理でエラーが発生しました: {str(e)}")
    finally:
        # 一時ファイルを削除（画像とPDF）
        try:
            if os.path.exists(pdf_path):
                os.unlink(pdf_path)
            if os.path.exists(img_path):
                os.unlink(img_path)
        except:
            pass

