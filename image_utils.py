#!/usr/bin/python3
# coding:utf-8

"""
画像処理ユーティリティモジュール
宛名画像生成に使用する汎用的な画像処理関数
"""

import os
import re
from PIL import Image, ImageDraw, ImageFont, ImageChops


def contraction(pil_image, size_xy):
    """与えた画像を、指定したサイズに収まるように縮小して返す。（拡大はしない）"""
    
    # 画層が幅、高さの両方で指定されたサイズより小さければ、そのまま返す。
    # （字を拡大してぼやけるのを避けたいので）
    if pil_image.size[0] < size_xy[0] and pil_image.size[1] < size_xy[1]:
        return pil_image
    
    # 横幅のみ指定サイズより大きければ、横方向に縮小。
    elif pil_image.size[0] > size_xy[0] and pil_image.size[1] <= size_xy[1]:
        resized_image = pil_image.resize(
            (size_xy[0], int(pil_image.size[1] * size_xy[0] / pil_image.size[0])), 
            Image.LANCZOS
        )
        return resized_image
    
    # 高さのみ指定サイズより大きければ、縦方向に縮小。
    elif pil_image.size[0] <= size_xy[0] and pil_image.size[1] > size_xy[1]:
        resized_image = pil_image.resize(
            (int(pil_image.size[0] * size_xy[1] / pil_image.size[1]), size_xy[1]), 
            Image.LANCZOS
        )
        return resized_image
    
    # 幅、高さともに指定サイスより大きいなら、縦横比しだいで幅か高さのどちらかを基準にして縮小する。
    else:
        if 1.0 * pil_image.size[0] / pil_image.size[1] > 1.0 * size_xy[0] / size_xy[1]:
            resized_image = pil_image.resize(
                (size_xy[0], int(pil_image.size[1] * size_xy[0] / pil_image.size[0])), 
                Image.LANCZOS
            )
            return resized_image
        else:
            resized_image = pil_image.resize(
                (int(pil_image.size[0] * size_xy[1] / pil_image.size[1]), size_xy[1]), 
                Image.LANCZOS
            )
            return resized_image


def pil_through_paste_greyscale(base_image, put_image, point_tuple, transparent_luminance):
    """
    特定の色を透明色扱いにして画像を重ねる関数（グレイスケール版）。
    透過色以外の点では、色が合成される。
    """
    if point_tuple[0] < 0:
        horizontal_min = point_tuple[0] * -1
        if point_tuple[0] + put_image.size[0] > base_image.size[0]:
            horizontal_max = point_tuple[0] * -1 + base_image.size[0]
        else:
            horizontal_max = put_image.size[0]
    else:
        horizontal_min = 0
        if point_tuple[0] + put_image.size[0] > base_image.size[0]:
            horizontal_max = base_image.size[0] - point_tuple[0]
        else:
            horizontal_max = put_image.size[0]
    
    if point_tuple[1] < 0:
        vertical_min = point_tuple[1] * -1
        if point_tuple[1] + put_image.size[1] > base_image.size[1]:
            vertical_max = point_tuple[1] * -1 + base_image.size[1]
        else:
            vertical_max = put_image.size[1]
    else:
        vertical_min = 0
        if point_tuple[1] + put_image.size[1] > base_image.size[1]:
            vertical_max = base_image.size[1] - point_tuple[1]
        else:
            vertical_max = put_image.size[1]
    
    part = put_image.crop((horizontal_min, vertical_min, horizontal_max, vertical_max))
    x_max = horizontal_max - horizontal_min
    x = 0
    y = 0
    
    part_data = part.getdata()
    
    for i in part_data:
        if i != transparent_luminance:
            current_pixel = (point_tuple[0] + x, point_tuple[1] + y)
            
            if current_pixel[0] >= 0 and current_pixel[1] >= 0:
                base_info = base_image.getpixel(current_pixel)
                
                luminance = i + base_info - 255
                if luminance < 0:
                    luminance = 0
                
                # 渡したbase_imageそのものを書き換えるので注意が必要。
                base_image.putpixel(current_pixel, luminance)
        
        x += 1
        if x == x_max:
            x = 0
            y += 1


def letter_to_pil_image(letter, fontpath, fontsize, max_square_width, rotation_degree=0):
    """文字と台紙イメージのサイズから、画像イメージを作成する関数。"""
    
    if not os.path.isfile(fontpath):
        print(f"文字を画像化する関数に渡されたフォントのパス「{fontpath}」が存在しません。\n真っ黒なイメージを返しておきます")
        return Image.new("L", (int(max_square_width / 2), int(max_square_width / 2)), 0)
    
    textimage_area = Image.new("L", (max_square_width, max_square_width), 255)
    fnt = ImageFont.truetype(fontpath, fontsize, encoding="unic")
    textimage_draw = ImageDraw.Draw(textimage_area)
    textimage_draw.text((0, 0), letter, font=fnt, fill="black")
    
    if rotation_degree != 0:
        textimage_area = textimage_area.rotate(rotation_degree)
    
    textimage_area = greyscale_autocrop(textimage_area)
    
    return textimage_area


def greyscale_autocrop(pil_image):
    """画像の余白を除去する"""
    empty_image = Image.new("L", pil_image.size, 255)
    diff = ImageChops.difference(pil_image, empty_image)
    bbox = diff.getbbox()
    if bbox:
        return pil_image.crop(bbox)
    else:
        print("greyscale_autocropが失敗したので、幅、高さが0のイメージを返しておきます。")
        print("原因としては、渡されたイメージが空白だった可能性があります。")
        return Image.new("L", (0, 0), 255)


def maybe_list_natsort(not_sorted_list):
    """
    ファイル名を「半角数字を大小で比較しながら」ソートする関数。
    Pythonでの自然順ソートの方法がわからなかったので、これを作った。
    """
    double_list = []
    sorted_list = []
    
    for not_sorted_element in not_sorted_list:
        expanded_name = ""
        zeropool = ""
        
        for i in range(len(not_sorted_element)):
            # 半角数字なら、半角数字が続く限り蓄積しておく
            if not_sorted_element[i] in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]:
                zeropool += not_sorted_element[i]
                # 最後の字なら、0を足して数字を結合して蓄積を初期化
                if i == len(not_sorted_element) - 1:
                    expanded_name += "0" * (255 - len(zeropool))
                    expanded_name += zeropool
                    zeropool = ""
            
            # 半角数字以外なら、そのまま結合
            # その前に、半角数字が蓄積していれば0を足して結合
            else:
                # 半角数字が蓄積していれば0を足して結合して蓄積を初期化
                if zeropool != "":
                    expanded_name += "0" * (255 - len(zeropool))
                    expanded_name += zeropool
                    zeropool = ""
                
                expanded_name += not_sorted_element[i]
        
        double_list.append([expanded_name, not_sorted_element])
    
    double_list.sort()
    sorted_list = [x[1] for x in double_list]
    
    return sorted_list
