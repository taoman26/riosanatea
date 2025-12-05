#!/usr/bin/python3
#coding:utf-8

# はがき宛名印刷ソフト「Riosanatea」：Version36


# Copyright 2017 天杉 善哉 (あますぎ ぜんざい Amasugi Zenzai) All rights reserved.

# Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

# 1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
# 2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

# THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

import wx
import sys
import re
import os
import copy
import json
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
from PIL import ImageChops
from PIL import ImageOps
import subprocess
import wx.lib.sheet
import wx.lib.scrolledpanel
import csv
import codecs
import configparser
import threading
import datetime
from io import BytesIO

# 分離したモジュールをインポート
from image_utils import (
    contraction, 
    pil_through_paste_greyscale, 
    letter_to_pil_image, 
    greyscale_autocrop,
    maybe_list_natsort
)
from csv_utils import csv_to_list, pil_printing



class atena_image_maker():

	def __init__( self, papersize_widthheight_millimetre = ( 100, 148 ), overwrite_settings = {} ):

		#画素数とミリメートルの変換比（pixel/mm）。用紙サイズや各パーツの配置の基準となる。
		self.mm_pixel_rate = 8

		#宛名の画像イメージの横幅と縦の長さを何ピクセルにするかを算出する。
		#横幅のピクセル値から、様々な寸法を連動して決定することになる。
		self.width = papersize_widthheight_millimetre[0] * self.mm_pixel_rate
		self.height = papersize_widthheight_millimetre[1] * self.mm_pixel_rate
		self.atena_baseimage = Image.new( "L", ( self.width, self.height ), 255 )

		self.A6_baseimage = Image.new( "L", ( 105 * self.mm_pixel_rate, 148 * self.mm_pixel_rate ), 255 )

		#各項目のフォントサイズ
		#レイアウトの〜-areasizeを大きくしても、このサイズ以上には字は大きくなりません。
		#万一、文字をものすごく大きくしたいという人がいましたら、ここを増やしてください。
		self.postalcode_fontsize = int( self.width / 7 )
		self.name_fontsize = int( self.width / 4 )
		self.company_fontsize = int( self.width / 9 )
		self.department_fontsize = int( self.width / 11 )
		self.address_fontsize = int( self.width / 10 )
		self.our_postalcode_fontsize = int( self.width / 11 )
		self.our_name_fontsize = int( self.width / 10 )
		self.our_address_fontsize = int( self.width / 15 )

		#フォントの画像を取得するための台紙画像の正方形の一辺の長さ
		#暫定的に宛名画像の幅にしておくが、あとでGUIオブジェクトから再決定関数を起動する
		self.postalcode_fontmat_size = self.width
		self.name_fontmat_size = self.width
		self.company_fontmat_size = self.width
		self.department_fontmat_size = self.width
		self.address_fontmat_size = self.width
		self.our_postalcode_fontmat_size = self.width
		self.our_name_fontmat_size = self.width
		self.our_address_fontmat_size = self.width

		#宛名レイアウトのデフォルト値
		#文字サイズをかなり大きくとってから縮小する方針のため
		#各フォントサイズはユーザーからは指定できないようにする。 GUIの入力欄は設置しない。
		self.parts_dict = { "postalcode-position" : [ 45, 14 ], "postalcode-letter-areasize" : [ 4, 5 ], "postalcode-fontsize" : self.postalcode_fontsize, "postalcode-placement" : [ 7, 14, 22, 29, 36, 43 ], "postalcode-direction" : [ "right", "down" ], "name-position" : [ 52, 63 ], "name-areasize" : [ 12, 84 ], "name-bind-space" : 2, "name-fontsize" : self.name_fontsize, "name-direction" : [ "center", "center" ], "honorific-space" : 4, "twoname-honorific-mode" : 1, "twoname-alignment-mode" : "bottom", "company-position" : [ 79, 24 ], "company-areasize" : [ 6, 80 ], "company-bind-space" : 1, "company-fontsize" : self.company_fontsize, "company-direction" : [ "left", "down" ], "department-position" : [ 72, 24 ], "department-areasize" : [ 6, 80 ], "department-bind-space" : 1, "department-fontsize" : self.department_fontsize, "department-direction" : [ "left", "down" ], "address-position" : [ 94, 24 ], "address-areasize" : [ 7, 100 ], "address-bind-space" : 1, "address-fontsize" : self.address_fontsize, "address-direction" : [ "left", "down" ], "our-postalcode-position" : [ 6, 123 ], "our-postalcode-letter-areasize" : [ 3, 3 ], "our-postalcode-fontsize" : self.our_postalcode_fontsize, "our-postalcode-placement" : [ 4, 8, 13, 17, 21, 25 ], "our-postalcode-direction" : [ "right", "down" ], "our-name-position" : [ 11, 118 ], "our-name-areasize" : [ 6, 60 ], "our-name-bind-space" : 1, "our-name-fontsize" : self.our_name_fontsize, "our-name-direction" : [ "center", "up" ], "our-address-position" : [ 19, 118 ], "our-address-areasize" : [ 4, 80 ], "our-address-bind-space" : 1, "our-address-fontsize" : self.our_address_fontsize, "our-address-direction" : [ "right", "up" ], "A6-adjust-mode" : "center", "A6-adjust-point" : [ 0, 0 ], "resize％" : [ 100, 100 ], "redline-width" : 2, "fontfile" : "NotoSansCJK-Regular.ttc" }

		#はがき用のデフォルト配置の控え
		#用紙サイズが変更された場合に、これをもとにレイアウト調整をするのでバックアップしておく
		self.standard_parts_dict = copy.deepcopy( self.parts_dict )

		#引数として与えられている設定上書き用の辞書で、設定の初期値を更新する
		self.parts_dict.update( overwrite_settings )


	#宛名レイアウトの値を書き換える関数
	def set_parts_data( self, dict_key, value, list_position = None ):

		#list_positionが渡されていない（None）なら、書き換え対象がリストではないか
		#リストであってもリスト全体をまるごと置き換える、という扱いでvalueを辞書に入れる。
		if list_position is None:
			self.parts_dict[ dict_key ] = value
			return True

		#書き換え対象がListで、list_positionが指定されているなら、リストの一部だけを変更する
		elif isinstance( self.parts_dict[ dict_key ], list ) and isinstance( list_position, int ):
			self.parts_dict[ dict_key ][ list_position ] = value

		else:
			return False


	#宛名レイアウトの値を取得する関数
	def get_parts_data( self, dict_key ):
		try:
			get_result = self.parts_dict[ dict_key ]
			return get_result

		except:
			print( "「get_parts_data」関数のエラー。指定された項目「" + dict_key + "」が、レイアウト設定を格納した辞書中に存在しません" )
			return "error at function [ get_parts_data ]"


	#宛名レイアウトの辞書に新しい辞書の内容を追加、上書きする関数
	def set_parts_dictionary( self, new_layout_dictionary ):
		self.parts_dict.update( new_layout_dictionary )


	#宛名レイアウトのレイアウト辞書全体（のコピー）を取得する関数
	def get_parts_dictionary( self ):
		return copy.deepcopy( self.parts_dict )


	#宛名レイアウトの標準配置の辞書（のコピー）を取得する関数
	def get_standard_parts_dictionary( self ):
		return copy.deepcopy( self.standard_parts_dict )


	#宛名、住所、差出人など、縦書き部分の画像を作成する関数。
	def vertical_text( self, text, font_path ,font_size, mat_size ):

		if not os.path.isfile( font_path ):
			print( "関数に渡されたフォントのパス「" + font_path + "」が存在しません。\nこれでは実行不可能なので終了します。" )
			sys.exit()

		image = Image.new( "L", ( int( self.width / 2 ), self.height * 4 ), 255 )
		font = ImageFont.truetype( font_path, font_size, encoding="unic" )
		vertical_position = 0 #それぞれの字を置く縦の位置

		#「|」の高さを半角・全角スペースの間隔を算出する基準とし、その1/5を文字の間隔（vertical_space）とする。
		letter_image = letter_to_pil_image( "|", font_path, font_size, max_square_width = mat_size )
		vertical_unit = letter_image.size[1]
		vertical_space = int( vertical_unit / 5 )

		#全角スペースや横線も、python2までは印刷文字列に合わせてutf-8化しないと比較できないので変数として確保しておく。
		fullsize_space = "　"
		horizontal_line1 = "─"
		horizontal_line2 = "━"
		horizontal_line3 = "＝"

		for i in range( len( text ) ):
			if text[i] == fullsize_space:
				vertical_position += int( vertical_unit * 0.8 )

			elif text[i] == " " :
				vertical_position += int( vertical_unit * 0.4 )

			else:
				if text[i] in [ "-", "=", horizontal_line1, horizontal_line2, horizontal_line3 ]:
					single_area = letter_to_pil_image( text[i], font_path, font_size, self.width, rotation_degree = 90 )

				elif text[i] in "[]()<>{}「」（）｛｝＜＞【】『』［］〈〉《》〔〕":
					single_area = letter_to_pil_image( text[i], font_path, font_size, self.width, rotation_degree = 270 )

				else:
					single_area = letter_to_pil_image( text[i], font_path, font_size, self.width )


				if single_area is None or single_area.size[0] == 0:
					vertical_position += int( vertical_unit * 0.4 )

				else:
					if single_area.size[1] > int( vertical_unit * 0.3 ):
						image.paste( single_area,( int( ( image.size[0] - single_area.size[0] ) / 2 ), vertical_position ) )
						vertical_position += single_area.size[1] + vertical_space
					else:
					#漢数字の一のように高さがあまりに小さい場合、前後の字との間隔を2倍にする。
						image.paste( single_area,( int( ( image.size[0] - single_area.size[0] ) / 2 ), vertical_position + vertical_space ) )
						vertical_position += single_area.size[1] + vertical_space * 3

		return greyscale_autocrop( image )


	#郵便番号のPILイメージを作成する
	#番号の間隔をあらかじめミリメートルで指定しているので、郵便番号については全体の縮小はしない。（1字ごとの縮小はする）
	def postal_code_build( self, code, center_list, font_path ,font_size, letter_size_list, mat_size ):

		letter_size_x = int( letter_size_list[0] * self.mm_pixel_rate )
		letter_size_y = int( letter_size_list[1] * self.mm_pixel_rate )
		pcode_image = Image.new( "L", ( self.width, int( self.width / 4 ) ), 255 )

		horizontal_zero_point = int( letter_size_x / 2 )

		for i in range( 0, 7 ):
			single_letter = letter_to_pil_image( code[i], font_path, font_size, max_square_width = mat_size )

			#まず、収めようとする範囲の高さに合わせて縮小（ここではアスペクト比は保持する）
			single_letter = single_letter.resize( ( int( single_letter.size[0] * letter_size_y / single_letter.size[1] ), letter_size_y ), Image.LANCZOS )

			#高さを揃えた上で、収めようとする範囲より横長なら、横幅だけを縮小する
			#※アスペクト比を捨てるので、収める範囲があまりに縦長だと見苦しくなる
			#しかし、アスペクト比を保つと数字の大きさがバラつきすぎる場合があるので、このほうがまだマシだと判断した
			if single_letter.size[0] > letter_size_x:
				single_letter = single_letter.resize( ( letter_size_x, single_letter.size[1] ), Image.LANCZOS )

			if i == 0:
				pcode_image.paste( single_letter,( int( ( letter_size_x - single_letter.size[0] ) / 2 ), 0 ) )
			else:
				pcode_image.paste( single_letter,( horizontal_zero_point + center_list[ i - 1 ] - int( single_letter.size[0] / 2 ), 0 ) )

		pcode_image = pcode_image.crop( ( 0, 0, center_list[5] + letter_size_x, letter_size_y ) )
		return pcode_image


	#与えた文字列から縦書き画像を作成して、指定位置の指定方向に来るように貼り付ける。
	def parts_setting( self, image, text1, text2, font, fontsize, position_xy, size_xy, mat_size, mm_space = 1, direction = ( "center", "center" ), alignment_mode = "address" ):

		if text1 == "" and text2 == "":
			return "渡された文字列（text1,text2）が両方とも空だったので、貼り付けしません。"

		#ミリメートルで指定された値をピクセルに変換する。
		parts_position = [ int( x * self.mm_pixel_rate ) for x in position_xy ]
		parts_size = [ int( x * self.mm_pixel_rate ) for x in size_xy ]
		space = mm_space * self.mm_pixel_rate

		#もしtext1が空だったら縮小する予定サイズ（parts_size）の空白画像を代わりに用意する
		if text1 == "":
			part_image1 = Image.new( "L", parts_size, 255 )
			resized_part1 = part_image1.copy()

		else:
			text1 = text1.replace( "ー", "|" )

			part_image1 = self.vertical_text( text1, font, fontsize, mat_size = mat_size )

		#もしtext1が半角か全角のスペースのみで、オートクロップが働かなかった場合
		#縮小する予定サイズ（parts_size）の画像を代わりに用意する
		#（失敗した場合にNoneを返すか空白イメージを返すか迷うので、両方に対応させておく）
		if part_image1 is None or part_image1.size[0] == 0:
			part_image1 = Image.new( "L", parts_size, 255 )
			resized_part1 = part_image1.copy()
		else:
			resized_part1 = contraction( part_image1, parts_size )

		oneline_areasize = [ resized_part1.size ]

		if text2 == "" or re.match( "^[ 　]+$", text2 ):
			united_parts = resized_part1

		#text2が指定されていれば、それも画像化して、text1の画像と結合する
		else:
			text2 = text2.replace( "ー", "|" )

			part_image2 = self.vertical_text( text2, font, fontsize, mat_size = mat_size )
			#もしtext2が半角か全角のスペースのみで、オートクロップが働かなかった場合
			#text1のイメージを空白化したものを代わりに用意する
			#（失敗した場合にNoneを返すか空白イメージを返すか迷うので、両方に対応させておく）
			if part_image2 is None or part_image2.size[0] == 0:
				part_image2 =part_image1.copy()
				if part_image2 is not None and part_image2.size[0] != 0:
					draw = ImageDraw.Draw( part_image2 )
					draw.rectangle( ( ( 0, 0 ), part_image2.size ), fill = "white" )

			#揃え方（alignment_mode）が住所なら1列目の幅に合わせて2列目をリサイズし、名前なら独立して大きさを決める
			if alignment_mode == "address":
				resized_part2 = contraction( part_image2, ( resized_part1.size[0], parts_size[1] ) )
			else:
				resized_part2 = contraction( part_image2, parts_size )

			oneline_areasize.append( resized_part2.size )

			#宛名2つの高い方の高さに合わせてunited_parts（宛名二列を連結する土台となる画像）を作成する
			if resized_part1.size[1] > resized_part2.size[1]:
				united_parts = Image.new( "L", ( int( resized_part1.size[0] + space + resized_part2.size[0] ), resized_part1.size[1] ), 255 )
			else:
				united_parts = Image.new( "L", ( int( resized_part1.size[0] + space + resized_part2.size[0] ), resized_part2.size[1] ), 255 )

			#まだunited_partsは空白の台紙なので、それぞれの宛名画像を貼り付ける
			#揃え方（alignment_mode）が住所モードであり、1行目のほうが短い場合は上端で揃える（両者とも上下方向の座標を0にして貼り付け）
			if alignment_mode == "address" and resized_part1.size[1] < resized_part2.size[1]:
				united_parts.paste( resized_part2, ( 0, 0 ) )
				united_parts.paste( resized_part1, ( int( resized_part2.size[0] + space ), 0 ) )

			#揃え方が名前モードで、かつ上寄せ設定の場合は上端で揃えて貼り付ける
			elif alignment_mode == "name" and self.parts_dict.get( "twoname-alignment-mode" ) == "top":
				united_parts.paste( resized_part2, ( 0, 0 ) )
				united_parts.paste( resized_part1, ( int( resized_part2.size[0] + space ), 0 ) )

			#それ以外、揃え方が名前モードで下寄せとか1行目が長いのなら下端で揃えて貼り付ける
			else:
				united_parts.paste( resized_part2, ( 0, united_parts.size[1] - resized_part2.size[1] ) )
				united_parts.paste( resized_part1, ( int( resized_part2.size[0] + space ), united_parts.size[1] - resized_part1.size[1] ) )

		#左右方向の指定に従い、貼り付け位置を決める
		if direction[0] == "right" :
			pastepoint_x = parts_position[0]
		elif direction[0] == "left" :
			pastepoint_x = parts_position[0] - united_parts.size[0]
		else:
			pastepoint_x = int( parts_position[0] - united_parts.size[0] / 2 )

		#上下方向の指定に従い、貼り付け位置を決める
		if direction[1] == "down" :
			pastepoint_y = parts_position[1]
		elif direction[1] == "up" :
			pastepoint_y = parts_position[1] - united_parts.size[1]
		else:
			pastepoint_y = int( parts_position[1] - united_parts.size[1] / 2 )

		pil_through_paste_greyscale( image, united_parts, ( pastepoint_x , pastepoint_y ), 255 )

		#貼り付けた範囲（始点、終点）と一列目（name1、address1など）の大きさを返す
		return { "start-point" : ( pastepoint_x , pastepoint_y ), "end-point" : ( pastepoint_x + united_parts.size[0], pastepoint_y + united_parts.size[1] ), "oneline-areasize" : oneline_areasize }


	#郵便番号から画像を作成して、指定位置の指定方向に来るように貼り付ける。
	def postalcode_setting( self, image, postal_code, font, fontsize, letter_size_xy, position_xy, center_mm_list, mat_size, direction = ( "center", "center" ) ):

		#郵便番号に全角数字があれば、半角に変換しておく
		postal_code = postal_code.replace( "０", "0" )
		postal_code = postal_code.replace( "１", "1" )
		postal_code = postal_code.replace( "２", "2" )
		postal_code = postal_code.replace( "３", "3" )
		postal_code = postal_code.replace( "４", "4" )
		postal_code = postal_code.replace( "５", "5" )
		postal_code = postal_code.replace( "６", "6" )
		postal_code = postal_code.replace( "７", "7" )
		postal_code = postal_code.replace( "８", "8" )
		postal_code = postal_code.replace( "９", "9" )
		#スペースやハイフンの類があれば除去
		postal_code = postal_code.replace( " ", "" )
		postal_code = postal_code.replace( "　", "" )
		postal_code = postal_code.replace( "-", "" )
		postal_code = postal_code.replace( "−", "" )
		postal_code = postal_code.replace( "ー", "" )
		postal_code = postal_code.replace( "―", "" )
		postal_code = postal_code.replace( "‐", "" )

		if postal_code == "" :
			#print( "渡された postal_code が空だったので、郵便番号は貼り付けしません。" )
			return "渡された postal_code が空だったので、郵便番号は貼り付けしません。"
		elif not re.match( "^[0-9]{7}$", postal_code ):
			print( "渡された postal_code が英数字7個ではないので、郵便番号は貼り付けしません。" )
			return "渡された postal_code が英数字7個ではないので、郵便番号は貼り付けしません。"

		#mmで指定された値をピクセルに変換する。
		#center_listは、先頭の数字の中心から何mm右に、以降の数字の中心を持ってくるかを示す。
		center_list = [ int( x * self.mm_pixel_rate ) for x in center_mm_list ]
		pc_position = [ int( x * self.mm_pixel_rate ) for x in position_xy ]

		pc_image = self.postal_code_build( postal_code, center_list, font, fontsize, letter_size_xy, mat_size )

		#左右方向の指定に従い、貼り付け位置を決める
		if direction[0] == "right" :
			pastepoint_x = pc_position[0]
		elif direction[0] == "left" :
			pastepoint_x = pc_position[0] - pc_image.size[0]
		else:
			pastepoint_x = int( pc_position[0] - pc_image.size[0] / 2 )

		#上下方向の指定に従い、貼り付け位置を決める
		if direction[1] == "down" :
			pastepoint_y = pc_position[1]
		elif direction[1] == "up" :
			pastepoint_y = pc_position[1] - pc_image.size[1]
		else:
			pastepoint_y = int( pc_position[1] - pc_image.size[1] / 2 )

		pil_through_paste_greyscale( image, pc_image, ( pastepoint_x , pastepoint_y ), 255 )


	#台紙画像の上に各部品を配置していき、宛名画像を作成する
	def get_atena_image( self, data_dict, area_frame = False ):

		atena_image = self.atena_baseimage.copy()

		# フォントサイズの倍率を取得
		resize_percent = self.parts_dict.get( "resize％", [ 100, 100 ] )
		font_scale = resize_percent[0] / 100.0

		#宛先の郵便番号
		self.postalcode_setting( image = atena_image, postal_code = data_dict.get( "postal-code", "" ), font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "postalcode-fontsize" ] * font_scale ), letter_size_xy = self.parts_dict[ "postalcode-letter-areasize" ], position_xy = self.parts_dict[ "postalcode-position" ], center_mm_list = self.parts_dict[ "postalcode-placement" ], direction = self.parts_dict[ "postalcode-direction" ], mat_size = self.postalcode_fontmat_size )

		#宛名
		name_result = self.parts_setting( image = atena_image, text1 = data_dict.get( "name1", "" ), text2 = data_dict.get( "name2", "" ), font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "name-fontsize" ] * font_scale ), position_xy = self.parts_dict[ "name-position" ], size_xy = self.parts_dict[ "name-areasize" ], mat_size = self.name_fontmat_size, mm_space = self.parts_dict[ "name-bind-space" ], direction = self.parts_dict[ "name-direction" ], alignment_mode = "name" )

		#名前が正しく貼り付けできていれば（辞書型が返ってくれば）
		#名前のフォントサイズなどを元に敬称の画像を作り、返値から位置を算出し貼り付ける。
		if isinstance( name_result, dict ) and data_dict.get( "honorific", "" ) != "" and not re.match( "^( |　)+$", data_dict.get( "honorific" ) ):

			#敬称画像の作成
			honorific_image_origine = self.vertical_text( data_dict.get( "honorific", "" ), self.parts_dict[ "fontfile" ], int( self.parts_dict[ "name-fontsize" ] * font_scale ), mat_size = self.name_fontmat_size )

			#名前が一列だけなら中央に一つ敬称を付ける。
			if len( name_result.get( "oneline-areasize" ) ) == 1:
				honorific_image = contraction( honorific_image_origine, ( name_result.get( "oneline-areasize" )[0][0], name_result.get( "oneline-areasize" )[0][1] * 3 ) )

				pil_through_paste_greyscale( atena_image, honorific_image, ( int( ( name_result.get( "start-point" )[0] + name_result.get( "end-point" )[0] ) / 2 - honorific_image.size[0] / 2 ), name_result.get( "end-point" )[1] + int( self.parts_dict[ "honorific-space" ] * self.mm_pixel_rate ) ), 255 )

			#宛名が二列あってtwoname-honorific-modeが1なら、二列の中間の大きさで中央に一つ敬称を付ける。
			elif self.parts_dict.get( "twoname-honorific-mode" ) == 1:
				honorific_image = contraction( honorific_image_origine, ( int( ( name_result.get( "oneline-areasize" )[0][0] + name_result.get( "oneline-areasize" )[1][0] ) / 2 ), int( ( name_result.get( "oneline-areasize" )[0][1] + name_result.get( "oneline-areasize" )[1][1] ) / 2 ) * 3 ) )

				pil_through_paste_greyscale( atena_image, honorific_image, ( int( ( name_result.get( "start-point" )[0] + name_result.get( "end-point" )[0] ) / 2 - honorific_image.size[0] / 2 ), name_result.get( "end-point" )[1] + int( self.parts_dict[ "honorific-space" ] * self.mm_pixel_rate ) ), 255 )

			#宛名が二列あってtwoname-honorific-modeが2なら、二列それぞれに敬称を付ける。
			elif self.parts_dict.get( "twoname-honorific-mode" ) == 2:
				honorific_image1 = contraction( honorific_image_origine, ( name_result.get( "oneline-areasize" )[0][0], name_result.get( "oneline-areasize" )[0][1] * 3 ) )
				honorific_image2 = contraction( honorific_image_origine, ( name_result.get( "oneline-areasize" )[1][0], name_result.get( "oneline-areasize" )[1][1] * 3 ) )

				pil_through_paste_greyscale( atena_image, honorific_image2, ( name_result.get( "start-point" )[0], name_result.get( "end-point" )[1] + int( self.parts_dict[ "honorific-space" ] * self.mm_pixel_rate ) ), 255 )

				pil_through_paste_greyscale( atena_image, honorific_image1, ( name_result.get( "end-point" )[0] - name_result.get( "oneline-areasize" )[0][0], name_result.get( "end-point" )[1] + int( self.parts_dict[ "honorific-space" ] * self.mm_pixel_rate ) ), 255 )

			#宛名が二列あってtwoname-honorific-modeが1,2以外（3を想定）なら、左側にのみ敬称を付ける。
			else:
				honorific_image2 = contraction( honorific_image_origine, ( name_result.get( "oneline-areasize" )[1][0], name_result.get( "oneline-areasize" )[1][1] * 3 ) )

				pil_through_paste_greyscale( atena_image, honorific_image2, ( name_result.get( "start-point" )[0], name_result.get( "end-point" )[1] + int( self.parts_dict[ "honorific-space" ] * self.mm_pixel_rate ) ), 255 )

		#宛先の会社名
		self.parts_setting( image = atena_image, text1 = data_dict.get( "company", "" ), text2 = "", font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "company-fontsize" ] * font_scale ), position_xy = self.parts_dict[ "company-position" ], size_xy = self.parts_dict[ "company-areasize" ], mat_size = self.company_fontmat_size, mm_space = self.parts_dict[ "company-bind-space" ], direction = self.parts_dict[ "company-direction" ], alignment_mode = "address" )

		#宛先の部署名
		self.parts_setting( image = atena_image, text1 = data_dict.get( "department", "" ), text2 = "", font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "department-fontsize" ] * font_scale ), position_xy = self.parts_dict[ "department-position" ], size_xy = self.parts_dict[ "department-areasize" ], mat_size = self.department_fontmat_size, mm_space = self.parts_dict[ "department-bind-space" ], direction = self.parts_dict[ "department-direction" ], alignment_mode = "address" )

		#宛先の住所
		self.parts_setting( image = atena_image, text1 = data_dict.get( "address1", "" ), text2 = data_dict.get( "address2", "" ), font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "address-fontsize" ] * font_scale ), position_xy = self.parts_dict[ "address-position" ], size_xy = self.parts_dict[ "address-areasize" ], mat_size = self.address_fontmat_size, mm_space = self.parts_dict[ "address-bind-space" ], direction = self.parts_dict[ "address-direction" ], alignment_mode = "address" )

		#差出人側の郵便番号
		self.postalcode_setting( image = atena_image, postal_code = data_dict.get( "our-postal-code", "" ), font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "our-postalcode-fontsize" ] * font_scale ), letter_size_xy = self.parts_dict[ "our-postalcode-letter-areasize" ], position_xy = self.parts_dict[ "our-postalcode-position" ], center_mm_list = self.parts_dict[ "our-postalcode-placement" ], direction = self.parts_dict[ "our-postalcode-direction" ], mat_size = self.our_postalcode_fontmat_size )

		#差出人の氏名
		self.parts_setting( image = atena_image, text1 = data_dict.get( "our-name1", "" ), text2 = data_dict.get( "our-name2", "" ), font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "our-name-fontsize" ] * font_scale ), position_xy = self.parts_dict[ "our-name-position" ], size_xy = self.parts_dict[ "our-name-areasize" ], mat_size = self.our_name_fontmat_size, mm_space = self.parts_dict[ "our-name-bind-space" ], direction = self.parts_dict[ "our-name-direction" ], alignment_mode = "name")

		#差出人の住所
		self.parts_setting( image = atena_image, text1 = data_dict.get( "our-address1", "" ), text2 = data_dict.get( "our-address2", "" ), font = self.parts_dict[ "fontfile" ], fontsize = int( self.parts_dict[ "our-address-fontsize" ] * font_scale ), position_xy = self.parts_dict[ "our-address-position" ], size_xy = self.parts_dict[ "our-address-areasize" ], mat_size = self.our_address_fontmat_size, mm_space = self.parts_dict[ "our-address-bind-space" ], direction = self.parts_dict[ "our-address-direction" ], alignment_mode = "address" )

		#郵便番号や住所や名前といった各パーツの最大範囲を示す枠を付加する
		if area_frame is True:
			frame_linewidth = int( self.width / 200 )

			#宛先の郵便番号の枠
			if data_dict.get( "postal-code", "" ) != "":
				last_postalcode_place = self.parts_dict[ "postalcode-placement" ][ len( self.parts_dict[ "postalcode-placement" ] ) - 1 ]
				postalcode_areasize_x = self.mm_pixel_rate * ( last_postalcode_place + self.parts_dict[ "postalcode-letter-areasize" ][0]  )
				postalcode_areasize_y = self.mm_pixel_rate * self.parts_dict[ "postalcode-letter-areasize" ][1]

				self.paste_area_frame( base_image = atena_image, area_size = ( postalcode_areasize_x, postalcode_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "postalcode-position" ] ], area_direction = self.parts_dict[ "postalcode-direction" ], line_width = frame_linewidth )

			#宛先の住所の枠
			#住所が1列か2列かで枠の幅を変える
			if data_dict.get( "address2", "" ) == "":
				destination_address_areasize_x = self.mm_pixel_rate * self.parts_dict[ "address-areasize" ][0]
			else:
				destination_address_areasize_x = self.mm_pixel_rate * ( self.parts_dict[ "address-areasize" ][0] * 2 + self.parts_dict[ "address-bind-space" ] )

			destination_address_areasize_y = self.mm_pixel_rate * self.parts_dict[ "address-areasize" ][1]

			self.paste_area_frame( base_image = atena_image, area_size = ( destination_address_areasize_x, destination_address_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "address-position" ] ], area_direction = self.parts_dict[ "address-direction" ], line_width = frame_linewidth )

			#宛先の氏名の枠
			namearea_width = self.parts_dict[ "name-areasize" ][0]
			namearea_height = self.parts_dict[ "name-areasize" ][1]

			#宛名が1列か2列かで枠の幅を変える
			if data_dict.get( "name2", "" ) == "":
				destination_name_areasize_x = self.mm_pixel_rate * namearea_width
			else:
				destination_name_areasize_x = self.mm_pixel_rate * ( namearea_width * 2 + self.parts_dict[ "name-bind-space" ] )

			destination_name_areasize_y = self.mm_pixel_rate * namearea_height

			#敬称の枠
			if data_dict.get( "honorific", "" ) == "":
				honorific_height = 0
			else:
				#宛名枠の縦横の小さいほうを敬称1字の最大高さだと仮定し、3字分を最大範囲とする
				if namearea_width < namearea_height:
					honorific_height = self.mm_pixel_rate * ( namearea_width * 3 + self.parts_dict[ "honorific-space" ] )
				else:
					honorific_height = self.mm_pixel_rate * ( namearea_height * 3 + self.parts_dict[ "honorific-space" ] )

			self.paste_area_frame( base_image = atena_image, area_size = ( destination_name_areasize_x, destination_name_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "name-position" ] ], area_direction = self.parts_dict[ "name-direction" ], line_width = frame_linewidth, additional_height = honorific_height )

			#宛先の会社名の枠
			if data_dict.get( "company", "" ) != "":
				destination_company_areasize_x = self.mm_pixel_rate * self.parts_dict[ "company-areasize" ][0]
				destination_company_areasize_y = self.mm_pixel_rate * self.parts_dict[ "company-areasize" ][1]
				self.paste_area_frame( base_image = atena_image, area_size = ( destination_company_areasize_x, destination_company_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "company-position" ] ], area_direction = self.parts_dict[ "company-direction" ], line_width = frame_linewidth )

			#宛先の部署名の枠
			if data_dict.get( "department", "" ) != "":
				destination_department_areasize_x = self.mm_pixel_rate * self.parts_dict[ "department-areasize" ][0]
				destination_department_areasize_y = self.mm_pixel_rate * self.parts_dict[ "department-areasize" ][1]
				self.paste_area_frame( base_image = atena_image, area_size = ( destination_department_areasize_x, destination_department_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "department-position" ] ], area_direction = self.parts_dict[ "department-direction" ], line_width = frame_linewidth )

			#差出人の住所の枠
			if data_dict.get( "our-address1", "" ) != "":

				if data_dict.get( "our-address2", "" ) == "":
					destination_our_address_areasize_x = self.mm_pixel_rate * self.parts_dict[ "our-address-areasize" ][0]
				else:
					destination_our_address_areasize_x = self.mm_pixel_rate * ( self.parts_dict[ "our-address-areasize" ][0] * 2 + self.parts_dict[ "our-address-bind-space" ] )

				destination_our_address_areasize_y = self.mm_pixel_rate * self.parts_dict[ "our-address-areasize" ][1]
				self.paste_area_frame( base_image = atena_image, area_size = ( destination_our_address_areasize_x, destination_our_address_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "our-address-position" ] ], area_direction = self.parts_dict[ "our-address-direction" ], line_width = frame_linewidth )

			#差出人の氏名の枠
			if data_dict.get( "our-name1", "" ) != "":
				if data_dict.get( "our-name2", "" ) == "":
					destination_our_name_areasize_x = self.mm_pixel_rate * self.parts_dict[ "our-name-areasize" ][0]
				else:
					destination_our_name_areasize_x = self.mm_pixel_rate * ( self.parts_dict[ "our-name-areasize" ][0] * 2 + self.parts_dict[ "our-name-bind-space" ] )

				destination_our_name_areasize_y = self.mm_pixel_rate * self.parts_dict[ "our-name-areasize" ][1]
				self.paste_area_frame( base_image = atena_image, area_size = ( destination_our_name_areasize_x, destination_our_name_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "our-name-position" ] ], area_direction = self.parts_dict[ "our-name-direction" ], line_width = frame_linewidth )

			#差出人の郵便番号の枠
			if data_dict.get( "our-postal-code", "" ) != "":
				last_our_postalcode_place = self.parts_dict[ "our-postalcode-placement" ][ len( self.parts_dict[ "our-postalcode-placement" ] ) - 1 ]
				our_postalcode_areasize_x = self.mm_pixel_rate * ( last_our_postalcode_place + self.parts_dict[ "our-postalcode-letter-areasize" ][0]  )
				our_postalcode_areasize_y = self.mm_pixel_rate * self.parts_dict[ "our-postalcode-letter-areasize" ][1]
				self.paste_area_frame( base_image = atena_image, area_size = ( our_postalcode_areasize_x, our_postalcode_areasize_y ), area_position = [ i * self.mm_pixel_rate for i in self.parts_dict[ "our-postalcode-position" ] ], area_direction = self.parts_dict[ "our-postalcode-direction" ], line_width = frame_linewidth )

		return atena_image


	#灰色の枠を作成し、ある位置から指定された方向にずらして台紙画像に貼り付ける
	def paste_area_frame( self, base_image, area_size, area_position, area_direction, line_width, additional_height = 0 ):
		#まず、引数にある線幅で枠の画像を作る
		frame_image = Image.new( "L", ( area_size[0], area_size[1] + additional_height ), 140 )
		frame_image.paste( Image.new( "L", ( area_size[0] - line_width * 2, area_size[1] - line_width * 2 + additional_height ), 255 ), ( line_width, line_width ) )

		#横方向の座標を方向指定に応じてずらす
		if area_direction[0] == "left":
			pastepoint_x = area_position[0] - area_size[0]
		elif area_direction[0] == "right":
			pastepoint_x = area_position[0]
		else:
			pastepoint_x = int( area_position[0] - area_size[0] / 2 )

		#縦方向の座標を方向指定に応じてずらす
		if area_direction[1] == "up":
			pastepoint_y = area_position[1] - area_size[1]
		elif area_direction[1] == "down":
			pastepoint_y = area_position[1]
		else:
			pastepoint_y = int( area_position[1] - area_size[1] / 2 )

		#台紙画像に貼り付ける
		#pasteで単純に貼り付けると既存の字を消してしまうし、枠と字が重なる場合に
		#どちらかだけにしたくないので、処理が重くなるが透過貼り付けの関数を使う
		pil_through_paste_greyscale( base_image, frame_image, ( pastepoint_x, pastepoint_y ), 255 )


	#上下左右の余白領域を消した宛名画像を取得する
	def get_cutted_atena_image( self, data_dict, space_tblr_mm_list = [ 0, 0, 0, 0 ], return_pasted_image = False, cutted_atena_image_upside_down = False ):

		#ミリメートルで指定された値をピクセルに変換する
		upper_space_pixel = int( space_tblr_mm_list[0] * self.mm_pixel_rate )
		down_space_pixel = int( space_tblr_mm_list[1] * self.mm_pixel_rate )
		left_space_pixel = int( space_tblr_mm_list[2] * self.mm_pixel_rate )
		right_space_pixel = int( space_tblr_mm_list[3] * self.mm_pixel_rate )

		origin_atena_image = self.get_atena_image( data_dict )

		if cutted_atena_image_upside_down is False:
			cutted_atena_image = origin_atena_image.crop( ( left_space_pixel, upper_space_pixel, origin_atena_image.size[0] - right_space_pixel, origin_atena_image.size[1] - down_space_pixel ) )

		else:
			cutted_atena_image = origin_atena_image.crop( ( right_space_pixel, down_space_pixel, origin_atena_image.size[0] - left_space_pixel, origin_atena_image.size[1] - upper_space_pixel ) )

		#印刷「プレビュー」用に、上下左右を切り取った画像を、元サイズの空白画像に貼り付けて返す
		#（プリンターで印刷できない部分が空白領域として存在している）
		if return_pasted_image is True:
			paste_image = Image.new( "L", origin_atena_image.size, 255 )

			if cutted_atena_image_upside_down is False:
				paste_image.paste( cutted_atena_image, ( left_space_pixel, upper_space_pixel ) )
			else:
				paste_image.paste( cutted_atena_image, ( right_space_pixel, down_space_pixel ) )

			return paste_image

		else:
			#印刷用に、「印刷不可能部分が削除されて存在しない」画像を返す
			return cutted_atena_image


	def make_A6_image( self, datadict_or_pilimage ):

		A6_image = self.A6_baseimage.copy()
		mode = self.parts_dict.get( "A6-adjust-mode" )
		paste_point = self.parts_dict.get( "A6-adjust-point" )
		resize_percent = self.parts_dict.get( "resize％" )

		if isinstance( datadict_or_pilimage, dict ):
			image = self.get_atena_image( datadict_or_pilimage )
		else:
			image = datadict_or_pilimage

		if resize_percent != [ 100, 100 ]:
			image = image.resize( ( int( image.size[0] * resize_percent[0] / 100 ), int( image.size[1] * resize_percent[1] / 100 ) ), Image.LANCZOS )

		if mode == "manual" :
			A6_adjust_pixel = ( int( paste_point[0] * self.mm_pixel_rate ), int( paste_point[1] * self.mm_pixel_rate ) )
			A6_image.paste( image, A6_adjust_pixel )

		elif mode == "center" :
			A6_image.paste( image, ( int( ( A6_image.size[0] - image.size[0] ) / 2 ), 0 ) ) #あくまで左右方向のcenter指定なので、上下の調整はしない。上下に動かしたいならmanualで。

		elif mode == "right" :
			A6_image.paste( image, ( A6_image.size[0] - image.size[0], 0 ) )

		else:
			A6_image.paste( image, ( 0, 0 ) ) #leftの場合を想定

		return A6_image


	# 実際の印刷サイズの検証用として、赤色で外縁部に枠が描かれた印刷用イメージを返す
	def get_red_frame_image( self ):
		line_width_pixel = self.parts_dict.get( "redline-width", 2 )

		#枠だけの画像を印刷しても方向がわからないので、上下左右を記述した画像を中央に貼る
		#上下左右の説明パーツをそれぞれ用意
		arrow_left1 = self.vertical_text( text = "←", font_path = self.parts_dict.get( "fontfile" ), font_size = int( self.width / 15 ), mat_size = int( self.width / 4 ) )
		arrow_left2 = self.vertical_text( text = "左余白", font_path = self.parts_dict.get( "fontfile" ), font_size = int( self.width / 15 ), mat_size = int( self.width / 4 ) )
		arrow_top_bottom = self.vertical_text( text = "↑ 上余白 ・ 下余白 ↓", font_path = self.parts_dict.get( "fontfile" ), font_size = int( self.width / 15 ), mat_size = int( self.width / 4 ) )
		arrow_right1 = self.vertical_text( text = "→", font_path = self.parts_dict.get( "fontfile" ), font_size = int( self.width / 15 ), mat_size = int( self.width / 4 ) )
		arrow_right2 = self.vertical_text( text = "右余白", font_path = self.parts_dict.get( "fontfile" ), font_size = int( self.width / 15 ), mat_size = int( self.width / 4 ) )
		arrow_space = Image.new( "RGB", ( arrow_top_bottom.size[0], arrow_top_bottom.size[1] ), ( 255, 255, 255 ) )

		#説明パーツを台紙に貼り付けて統合する
		paste_point_x = 0
		total_arrow_image = Image.new( "RGB", ( arrow_left1.size[0] + arrow_left2.size[0] + arrow_top_bottom.size[0] + arrow_right1.size[0] + arrow_right2.size[0] + arrow_space.size[0] * 2, arrow_top_bottom.size[1] ), ( 255, 255, 255 ) )
		total_arrow_image.paste( arrow_left1, ( 0, int( ( total_arrow_image.size[1] -  arrow_left1.size[1] ) / 2 ) ) )
		paste_point_x += arrow_left1.size[0]
		total_arrow_image.paste( arrow_left2, ( paste_point_x, int( ( total_arrow_image.size[1] -  arrow_left2.size[1] ) / 2 ) ) )
		paste_point_x += arrow_left2.size[0]
		total_arrow_image.paste( arrow_space, ( paste_point_x, 0 ) )
		paste_point_x += arrow_space.size[0]
		total_arrow_image.paste( arrow_top_bottom, ( paste_point_x, 0 ) )
		paste_point_x += arrow_top_bottom.size[0]
		total_arrow_image.paste( arrow_space, ( paste_point_x, 0 ) )
		paste_point_x += arrow_space.size[0]
		total_arrow_image.paste( arrow_right2, ( paste_point_x, int( ( total_arrow_image.size[1] -  arrow_right2.size[1] ) / 2 ) ) )
		paste_point_x += arrow_right2.size[0]
		total_arrow_image.paste( arrow_right1, ( paste_point_x, int( ( total_arrow_image.size[1] -  arrow_right1.size[1] ) / 2 ) ) )

		#台紙画像の横幅の何分の一かの大きさにリサイズする
		resize_width = int( self.width / 3 )
		total_arrow_image.resize( ( resize_width, int( total_arrow_image.size[1] * resize_width / total_arrow_image.size[0] ) ), Image.BILINEAR )

		#一面赤の画像に白画像を貼ることで枠画像を作り、中央に方向説明を貼る
		frame_image = Image.new( "RGB", ( self.atena_baseimage.size[0], self.atena_baseimage.size[1] ), ( 255, 0, 0 ) )
		frame_image.paste( Image.new( "RGB", ( frame_image.size[0] - line_width_pixel * 2, frame_image.size[1] - line_width_pixel * 2 ), ( 255, 255, 255 ) ), ( line_width_pixel, line_width_pixel ) )
		frame_image.paste( total_arrow_image, ( int( ( frame_image.size[0] - total_arrow_image.size[0] ) / 2 ), int( ( frame_image.size[1] - total_arrow_image.size[1] ) / 2 ) ) )

		return frame_image


	#宛名画像生成オブジェクト外での宛名画像の操作用に、ミリメートル単位の長さを
	#宛名画像のピクセル単位に変換して返す（主にサンプル画像への赤枠追加用）
	def convert_mm_to_pixel( self, millimeter_value ):
		if isinstance( millimeter_value, int ) or isinstance( millimeter_value, float ):
			return int( millimeter_value * self.mm_pixel_rate )

		elif isinstance( millimeter_value, list ):
			return [ int( x * self.mm_pixel_rate ) for x in millimeter_value ]

		else:
			return False


	#フォントを画像にする際の台紙となる正方形画像の一辺の長さを決める
	def determine_fontmat_size( self, font_path ):
		max_fontsize = max( ( self.parts_dict[ "postalcode-fontsize" ], self.parts_dict[ "name-fontsize" ], self.parts_dict[ "address-fontsize" ], self.parts_dict[ "our-postalcode-fontsize" ], self.parts_dict[ "our-name-fontsize" ], self.parts_dict[ "our-address-fontsize" ] ) )

		#これからフォント画像を取得していくための仮の台紙長さを決める
		temp_letter_image = letter_to_pil_image( "|", font_path, max_fontsize, self.width )
		temp_mat_size = int( temp_letter_image.size[1] * 3 )

		temp_letter_image = letter_to_pil_image( "|", font_path, self.parts_dict[ "postalcode-fontsize" ], temp_mat_size )
		self.postalcode_fontmat_size = int( temp_letter_image.size[1] * 1.4 )

		temp_letter_image = letter_to_pil_image( "|", font_path, self.parts_dict[ "name-fontsize" ], temp_mat_size )
		self.name_fontmat_size =  int( temp_letter_image.size[1] * 2 )

		temp_letter_image = letter_to_pil_image( "|", font_path, self.parts_dict[ "address-fontsize" ], temp_mat_size )
		self.address_fontmat_size =  int( temp_letter_image.size[1] * 2 )

		temp_letter_image = letter_to_pil_image( "|", font_path, self.parts_dict[ "our-postalcode-fontsize" ], temp_mat_size )
		self.our_postalcode_fontmat_size =  int( temp_letter_image.size[1] * 1.4 )

		temp_letter_image = letter_to_pil_image( "|", font_path, self.parts_dict[ "our-name-fontsize" ], temp_mat_size )
		self.our_name_fontmat_size =  int( temp_letter_image.size[1] * 2 )

		temp_letter_image = letter_to_pil_image( "|", font_path, self.parts_dict[ "our-address-fontsize" ], temp_mat_size )
		self.our_address_fontmat_size =  int( temp_letter_image.size[1] * 2 )


#-----宛名画像の生成クラスはここまで

# ユーティリティ関数は image_utils.py と csv_utils.py に移動しました
# - contraction()
# - pil_through_paste_greyscale()
# - letter_to_pil_image()
# - greyscale_autocrop()
# - pil_printing()
# - maybe_list_natsort()
# - csv_to_list()


#整数・小数入力用に小規模な改造をしたテキスト入力
#今回は差出人郵便番号のチェックに使用
class TextCtrlPlus( wx.TextCtrl ):

	def __init__( self, *args, **kwargs ):
		wx.TextCtrl.__init__( self, *args, **kwargs )

		#TextCtrlに想定外の値が入った場合に、空欄や直前の値に戻すことにしたが
		#そうすると、そのTextCtrlへの自動書き換え自体が入力と判定されてしまうようだ。
		#Bindした関数が再度呼びだされておかしなことになるため、管理用の目印を追加した。
		self.autoset_check = False

	def get_check( self ):
		return self.autoset_check

	def set_check( self, flag_bool ):
		self.autoset_check = flag_bool


# maybe_list_natsort() は image_utils.py に移動しました


#GUI部分の構築
class frame_plus( wx.Frame ):

	def __init__( self, *args, **kwargs ):
		wx.Frame.__init__( self, *args, **kwargs )

		## ********** 初期値、内部データ **********
		self.base_window_title = self.GetTitle()

		self.paper_size_data = { "category" : "はがき", "width" : 100, "height" : 148 }

		self.software_setting = { "window_maximize" : False, "window_size" : [ 1600, 760 ], "table-font" : "", "table-fontsize" : 0, "write-fileinfo-on-titlebar" : "filename"  }

		self.column_etc_dictionary = { "column-postalcode" : 1, "column-address1" : 2, "column-address2" : 3, "column-name1" : 4, "column-name2" : 5, "column-company" : 6, "column-department" : 7, "enable-default-honorific" : True, "default-honorific" : "様", "printer-space-top,bottom,left,right" : [ 0, 0, 0, 0 ], "print-control" : False, "print-control-column" : 0, "print-sign" : "×", "print-or-ignore" : "ignore", "enable-honorific-in-table" : True, "column-honorific" : 8, "sampleimage-areaframe" : True, "upside-down-print" : False }

		self.our_data = { "our-postalcode-data" : "", "our-name1-data" : "", "our-name2-data" : "", "our-address1-data" : "", "our-address2-data" : "" }

		#はがきと封筒の規格（名称とサイズ）
		self.post_card_envelope_standard = { "はがき,100x148", "往復はがき,200x148", "長形1号,142x332", "長形2号,119x277", "長形3号,120x235", "長形4号,90x205", "長形5号,90x185", "長形6号,110x220", "長形8号,119x197", "長形13号,105x235", "長形14号,95x217", "長形30号,92x235", "長形40号,90x225", "角形0号,287x382", "角形1号,270x382", "角形2号,240x332", "角形3号,216x277", "角形4号,197x267", "角形5号,190x240", "角形6号,162x229", "角形7号,142x205", "角形8号,119x197", "角形20号,229x324", "角形A3号,335x490", "角形A4号,228x312", "角形B3号,375x525", "角形0号マチ付,290x382", "角形2号マチ付,250x335", "角形3号マチ付,218x277", "角形ジャンボ,435x510", "洋形2号タテ形,114x162", "洋形4号タテ形,105x235", "洋形5号タテ形,95x217", "洋形6号タテ形,98x190" }

		#印刷を途中で中止するための変数
		self.print_stop_flag = False

		#宛名のイメージ例に使う架空の宛先
		self.destination_postalcode_example = "1234567"
		self.destination_name1_example = "宛名田 葉書"
		self.destination_name2_example = "菜野代"
		self.destination_address1_example = "架空県一応市地域の例1-23"
		self.destination_address2_example = "ナントナク456-7号室"

		#宛名のイメージ例を表示するためのもろもろ1
		self.sample_grayscale_image = Image.new( "L", ( 450, 500 ), 0 )
		self.sample_color_image = Image.new( "RGB", ( 450, 500 ), ( 0, 0, 0 ) )

		#宛名のイメージ例を表示するためのもろもろ2
		self.wx_sample_image = wx.Image( 200, 200 )
		self.wx_resized_sample = self.wx_sample_image.Copy()
		self.wx_bitmap_image = self.wx_resized_sample.ConvertToBitmap()
		self.display_position = [ 0, 0 ]

		#使用できるフォント
		self.fonts_data = self.get_fontlist()
		self.outer_font = "" #システムフォントの一覧が取得できなかった場合に使うシステム外（かもしれない）フォント
		self.table_font = ""
		self.table_fontsize = 0
		self.table_systemdefault_font = "" #フォントを何も指定しなかった時に使われているフォント

		#表の検索関連
		self.find_list = []
		self.current_find_number = 0

		#住所表の履歴
		self.table_history = [] #内容変更やCSVの読み込みによる変化を記録した、住所表の履歴
		self.current_history_position = 0 #アンドゥ、リドゥで、現在どこまで戻っているかという位置（csv_history[？]の？になる。0が最新の位置で、古いものほど数字が増える後の方に押し出されていく）
		self.history_stock_max = 20 #履歴の最大数

		#宛名画像作成インスタンスの作成時に引数として渡す、設定値上書き用辞書
		#もしINIファイルが読み込まれたら、保存されていた設定値でこれを上書きする
		self.overwrite_dict_for_image_generator = {}

		#INIファイルからの設定値読み込み
		#今回はコマンドライン引数は使わない（INIから読み込んだ値を引数で上書きしない）ので
		#GUI内部に収納しておく
		self.load_settings()

		#設定を読みこんだら、最新になった用紙サイズと設定値の辞書をもとに宛名画像作成インスタンスを作る
		self.image_generator = atena_image_maker( papersize_widthheight_millimetre =  ( self.paper_size_data[ "width" ], self.paper_size_data[ "height" ] ), overwrite_settings = self.overwrite_dict_for_image_generator )

		#ないとは思うが、読み込んだ用紙サイズがデフォルトの用紙サイズ候補の中にない場合のために
		#INIに記されている用紙サイズを候補リスト(重複しないようにsetにしてある)に追加しておく
		self.post_card_envelope_standard.add( self.paper_size_data[ "category" ] + "," + str( self.paper_size_data[ "width" ] ) + "x" + str( self.paper_size_data[ "height" ] ) )

		#設定を読みこんだら、それを元にウィンドウを設定する
		self.SetSize( self.software_setting[ "window_size" ][0], self.software_setting[ "window_size" ][1] )

		if self.software_setting[ "window_maximize" ] is True:
			self.Maximize( True )
		else:
			self.Maximize( False )

		#フォント一覧の取得に失敗し、INIの設定を読み込んでも使えるフォント情報がない場合は
		#ファイル選択ダイアログからフォントファイルを読み込む
		self.dict_font_data = self.image_generator.get_parts_data( "fontfile" )
		if self.fonts_data == [] and self.outer_font == "" and not os.path.isfile( self.image_generator.get_parts_data( "fontfile" ) ):

			fdialog = wx.FileDialog( self, "フォントの自動取得に失敗したのでフォントを選択してください", defaultDir = os.path.split( sys.argv[0] )[0], wildcard = "*.ttf;*.ttc;*.otf" )

			if fdialog.ShowModal() == wx.ID_OK:
				filename = fdialog.GetFilename()
				dirpath = fdialog.GetDirectory()

				self.outer_font = os.path.join ( dirpath, filename )
				self.image_generator.set_parts_data( "fontfile", self.outer_font )

			fdialog.Destroy()

		#フォント一覧はあるが、宛名画像生成インスタンス内のデフォルトフォントが
		#ファイルとしても一覧中のフォント名としてもないのなら、一覧からnotoフォントを探して入れ替える。
		#（Ubuntu18.04から日本語フォントがNotoシリーズになりそうという情報を見かけたので）
		elif not os.path.isfile( self.dict_font_data ) and not self.dict_font_data in [ x[0] for x in self.fonts_data ]:

			#Notoが見つからない場合の保険として、暫定的に一覧の最初のフォントで設定しておく
			self.image_generator.set_parts_data( "fontfile", self.fonts_data[0][1] )

			for fname in [ x[0] for x in self.fonts_data ]:
				if re.search( "Noto.*CJK.*Regular", fname ):
					self.image_generator.set_parts_data( "fontfile", self.fonts_data[ [ x[0] for x in self.fonts_data ].index( fname ) ][1] )
					break

		#郵便番号や住所、宛名の各画像を得るための台紙サイズを決定しておく
		self.image_generator.determine_fontmat_size( self.image_generator.get_parts_data( "fontfile" ) )

		#設定を読みこんでフォント設定も一段落したので、各設定値辞書の初期値をコピーして記録しておく
		#ソフトの終了時に、その時点での設定と比較して設定変更があったかチェックするためのもの
		#設定保存したら、これを上書きする
		self.savepoint_paper_size_data = copy.deepcopy( self.paper_size_data )
		self.savepoint_parts_dictionary = copy.deepcopy( self.image_generator.get_parts_dictionary() )
		self.savepoint_software_setting = copy.deepcopy( self.software_setting )
		self.savepoint_column_etc_dictionary = copy.deepcopy( self.column_etc_dictionary )
		self.savepoint_our_data = copy.deepcopy( self.our_data )


		#キー操作をバインド
		id_shift_f3 = wx.NewIdRef().GetId()
		id_nomal_f3 = wx.NewIdRef().GetId()
		id_ctrl_f = wx.NewIdRef().GetId()
		id_ctrl_shift_z = wx.NewIdRef().GetId()
		id_ctrl_z = wx.NewIdRef().GetId()

		self.Bind( wx.EVT_MENU, self.key_shift_f3, id = id_shift_f3 )
		self.Bind( wx.EVT_MENU, self.key_nomal_f3, id = id_nomal_f3 )
		self.Bind( wx.EVT_MENU, self.key_ctrl_f, id = id_ctrl_f )
		self.Bind( wx.EVT_MENU, self.key_ctrl_shift_z, id = id_ctrl_shift_z )
		self.Bind( wx.EVT_MENU, self.key_ctrl_z, id = id_ctrl_z )

		self.hotkey_list = []
		self.hotkey_list.append( ( wx.ACCEL_SHIFT, wx.WXK_F3, id_shift_f3 ) )
		self.hotkey_list.append( ( wx.ACCEL_NORMAL, wx.WXK_F3, id_nomal_f3 ) )
		self.hotkey_list.append( ( wx.ACCEL_CTRL, ord( "f" ), id_ctrl_f ) )
		self.hotkey_list.append( ( wx.ACCEL_CTRL | wx.ACCEL_SHIFT, ord( "z" ), id_ctrl_shift_z ) )
		self.hotkey_list.append( ( wx.ACCEL_CTRL, ord( "z" ), id_ctrl_z ) )

		self.SetAcceleratorTable( wx.AcceleratorTable( self.hotkey_list ) )


		## ********** ここからGUI構築 **********
		self.notebook = wx.Notebook( self, wx.ID_ANY )

		self.atena_tab_panel = wx.Panel( self.notebook, wx.ID_ANY )
		self.layout_tab_panel = wx.Panel( self.notebook, wx.ID_ANY )
		self.appropriate_tab_panel = wx.lib.scrolledpanel.ScrolledPanel( self.notebook, wx.ID_ANY )
		self.setting_tab_panel = wx.lib.scrolledpanel.ScrolledPanel( self.notebook, wx.ID_ANY )
		self.notebook.AddPage( self.atena_tab_panel, '住所表編集、宛名印刷' )
		self.notebook.AddPage( self.layout_tab_panel, '印刷レイアウト' )
		self.notebook.AddPage( self.appropriate_tab_panel, '差出人記述・敬称等' )
		self.notebook.AddPage( self.setting_tab_panel, 'ソフトウェア設定' )

		self.statusbar = self.CreateStatusBar()
		#タブの切り替えで、住所表のタブ以外ではステータスバーを隠すようにバインド
		self.notebook.Bind( wx.EVT_NOTEBOOK_PAGE_CHANGED, self.showhide_statusbar_with_tab )

		#■■■■■住所表、印刷パネル■■■■■
		self.csvopen_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "開く", size = ( 60, -1 ) )
		self.csvsave_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "保存", size = ( 60, -1 ) )
		self.history_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "履歴" )
		self.grep_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "検索・置換" )
		self.row_column_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "行列の加減" )
		self.pcode_grep_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "郵便番号検索" )

		self.csvopen_button.SetToolTip( "CSV住所録を開く。クリックするとファイル選択ウィンドウが開きます。" )
		self.csvsave_button.SetToolTip( "表をCSV形式で保存する" )
		self.history_button.SetToolTip( "もとに戻す（リドゥ）、やり直し（アンドゥ）のことです" )
		self.grep_button.SetToolTip( "指定した文字列を、表中から検索ないし置換します" )
		self.row_column_button.SetToolTip( "末尾に行や列を追加したり、最後の行や列を削除します" )
		self.pcode_grep_button.SetToolTip( "この検索には、郵政公社が配布しているデータが必要です" )

		self.print_start_line = wx.SpinCtrl( self.atena_tab_panel, wx.ID_ANY, value = "1", min = 1, max = 1000000, size = ( 150, 30 ) )
		self.print_start_line.SetMinSize( ( 150, 30 ) )
		self.print_end_line = wx.SpinCtrl( self.atena_tab_panel, wx.ID_ANY, value = "1", min = 1, max = 1000000, size = ( 150, 30 ) )
		self.print_end_line.SetMinSize( ( 150, 30 ) )
		self.print_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "宛名印刷する", size = ( 120, 30 ) )
		self.print_button.SetMinSize( ( 120, 30 ) )
		self.print_button.SetMaxSize( ( 120, 30 ) )
		self.print_button.SetToolTip( "ここをクリックすると印刷を開始します。印刷処理中は中止ボタンに変わります（すでに印刷ジョブに処理し終わった分までは取り消しません）" )
		self.preview_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "イメージ確認", size = ( 120, 30 ) )
		self.preview_button.SetMinSize( ( 120, 30 ) )
		self.preview_button.SetMaxSize( ( 120, 30 ) )
		self.preview_button.SetToolTip( "ここをクリックすると、指定された範囲の各行をどのような外観で印刷するかを、実際に印刷しないで画面上で確認できます" )

		self.upsidedown_print_checkbox = wx.CheckBox( self.atena_tab_panel, wx.ID_ANY, "反転印刷" )
		if self.column_etc_dictionary[ "upside-down-print" ] is True:
			self.upsidedown_print_checkbox.SetValue( True )
		else:
			self.upsidedown_print_checkbox.SetValue( False )

		self.close_button = wx.Button( self.atena_tab_panel, wx.ID_ANY, "終了", size = ( 60, 30 ) )
		self.close_button.SetMinSize( ( 60, 30 ) )
		self.close_button.SetMaxSize( ( 60, 30 ) )
		self.close_button.SetToolTip( "ウィンドウを閉じて、このソフトを終了します" )

		self.upsidedown_print_checkbox.SetToolTip( "封筒の印刷を安定させるために、下方向からプリンターに給紙して180°反転した印刷にしたい場合は、これにチェックを入れてください" )

		#CSVファイルのボタンと印刷入力欄を1行にまとめる
		self.csv_and_print_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.csv_and_print_sizer.Add( self.csvopen_button, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( self.csvsave_button, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( wx.StaticLine( self.atena_tab_panel, style = wx.LI_VERTICAL ), 0, wx.LEFT | wx.RIGHT, 4 )
		self.csv_and_print_sizer.Add( self.history_button, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( self.grep_button, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( self.row_column_button, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( self.pcode_grep_button, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( wx.StaticLine( self.atena_tab_panel, style = wx.LI_VERTICAL ), 0, wx.LEFT | wx.RIGHT, 6 )
		self.csv_and_print_sizer.Add( self.print_start_line, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( wx.StaticText( self.atena_tab_panel, wx.ID_ANY, "行から" ), 0, wx.FIXED_MINSIZE | wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT, 2 )
		self.csv_and_print_sizer.Add( self.print_end_line, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( wx.StaticText( self.atena_tab_panel, wx.ID_ANY, "行までを" ), 0, wx.FIXED_MINSIZE | wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT, 2 )
		self.csv_and_print_sizer.Add( self.print_button, 0, wx.FIXED_MINSIZE | wx.LEFT, 2 )
		self.csv_and_print_sizer.Add( wx.StaticText( self.atena_tab_panel, wx.ID_ANY, "ないし" ), 0, wx.FIXED_MINSIZE | wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT, 4 )
		self.csv_and_print_sizer.Add( self.preview_button, 0, wx.FIXED_MINSIZE )
		self.csv_and_print_sizer.Add( self.upsidedown_print_checkbox, 0, wx.FIXED_MINSIZE | wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8 )
		self.csv_and_print_sizer.Add( self.close_button, 0, wx.FIXED_MINSIZE | wx.LEFT, 8 )

		#SpinCtrlの初期値を設定
		self.print_start_line.SetValue( 1 )
		self.print_end_line.SetValue( 1 )

		#CSV関連のボタンと印刷行の入力欄をバインド
		self.csvopen_button.Bind( wx.EVT_BUTTON, self.fileselect_and_opencsv )
		self.csvsave_button.Bind( wx.EVT_BUTTON, self.save_csv_file )
		self.print_button.Bind( wx.EVT_BUTTON, self.atenaprint_thread_control )
		self.preview_button.Bind( wx.EVT_BUTTON, self.show_preview_dialog )

		self.history_button.Bind( wx.EVT_BUTTON, self.goto_history_point )
		self.grep_button.Bind( wx.EVT_BUTTON, self.table_search )
		self.row_column_button.Bind( wx.EVT_BUTTON, self.row_col_add_del )
		self.pcode_grep_button.Bind( wx.EVT_BUTTON, self.postalcode_search )
		self.close_button.Bind( wx.EVT_BUTTON, self.window_close )

		self.upsidedown_print_checkbox.Bind( wx.EVT_CHECKBOX, self.change_upsidedown_print )

		#終了ボタンのBindの直下に、終了時の設定変更チェックのBindもついでに書いておく
		self.Bind( wx.EVT_CLOSE, self.check_at_close )

		#住所表の作成と設定
		self.grid = wx.grid.Grid( self.atena_tab_panel )
		self.grid.CanDragCell() #これを作っている時点では、効果がないようだ
		self.grid.CanDragColMove() #これを作っている時点では、効果がないようだ
		self.grid.CreateGrid( 5, 7 )
		self.set_grid_labels()
		self.table_history.append( [ [ "" for x in range( self.grid.GetNumberCols() ) ] for y in range( self.grid.GetNumberRows() ) ] ) #開始時点の空白の表を、最初の履歴にしておく

		#表の内容の控えを作る
		#終了時やファイルを開く前に、これと比較して変更が保存されているかどうかチェックする
		#その性質上、ファイルを開いた直後や保存した後に更新しておく必要がある
		self.table_checkpoint = self.get_current_table_list()
		#開いたファイルのパスも今のうちに用意する
		self.opened_file_path = ""

		#印刷行範囲のデフォルト値を表の行数に合わせる
		self.print_start_line.SetValue( 1 )
		self.print_end_line.SetValue( self.grid.GetNumberRows() )

		#表のフォントとフォントサイズを設定する
		self.table_systemdefault_font = self.grid.GetDefaultCellFont().GetFaceName() #フォントを設定する前に、デフォルトで使われれるフォントを取得しておく

		self.table_font = self.software_setting[ "table-font" ]
		self.table_fontsize = self.software_setting[ "table-fontsize" ]
		self.table_fontlist = [ [ "", "" ] ] + self.fonts_data #宛名イメージのフォント一覧を流用

		if self.table_font != "" and self.software_setting[ "table-fontsize" ] > 0:

			self.table_fontdata = wx.Font( self.table_fontsize, family = wx.DEFAULT, style = wx.NORMAL, weight = wx.NORMAL, underline=False, faceName = self.table_font )
			self.grid.SetDefaultCellFont( self.table_fontdata )

			#フォントあるいはフォントサイズを変更したので、表の行・列のサイズを調整する
			self.grid.AutoSizeColumns()
			self.grid.AutoSizeRows()


		#住所表の変更を、履歴を蓄積する関数にバインドする
		self.grid.Bind( wx.grid.EVT_GRID_CELL_CHANGED, self.add_table_to_history )

		#ボタン、入力欄、表を立てに並べてパネルに貼る
		self.atena_tab_sizer = wx.BoxSizer( wx.VERTICAL )
		self.atena_tab_sizer.Add( self.csv_and_print_sizer, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM | wx.EXPAND, 6 )
		self.atena_tab_sizer.Add( self.grid, 1, wx.ALL | wx.EXPAND )

		self.atena_tab_panel.SetSizer( self.atena_tab_sizer )


		#■■■■■レイアウト設定のパネル■■■■■
		self.right_panels_color = "#CCCCCC"
		self.left_panel = wx.lib.scrolledpanel.ScrolledPanel( self.layout_tab_panel, wx.ID_ANY, size = self.GetClientSize() ) #ここにおくSpinCtrlやComboBoxの初期値設定はその場では行わず、関数にまとめてある
		self.right_panel = wx.lib.scrolledpanel.ScrolledPanel( self.layout_tab_panel, wx.ID_ANY, size = self.GetClientSize() )
		self.right_panel.SetBackgroundColour( self.right_panels_color )

		self.sample_image_panel = wx.Panel( self.right_panel, wx.ID_ANY )
		self.sample_image_panel.SetBackgroundColour( self.right_panels_color )
		#表示画像が更新されるようにバインド
		self.sample_image_panel.Bind( wx.EVT_PAINT, self.OnPaint )
		#パネルの大きさが変わったら画像がリサイズされるようにバインド
		self.sample_image_panel.Bind( wx.EVT_SIZE, self.adjust_sample_image_with_panel )

		#枠（StaticBoxSizer）に入れる
		self.sampleimage_sbox = wx.StaticBox( self.right_panel, wx.ID_ANY, "印刷の参考イメージ ( " + self.paper_size_data[ "category" ] + "、" + str( self.paper_size_data[ "width" ] ) + "mm x " + str( self.paper_size_data[ "height" ] ) + "mm )" )
		self.sampleimage_sb_sizer = wx.StaticBoxSizer( self.sampleimage_sbox, wx.VERTICAL )
		self.sampleimage_sb_sizer.Add( self.sample_image_panel, 1, wx.ALL | wx.EXPAND, 10 )


		#○○○○○現在の用紙サイズとレイアウトの保存ボタンと読み込みボタン○○○○○
		self.layoutfile_save_button = wx.Button( self.left_panel, wx.ID_ANY, "パーツのレイアウトと用紙サイズを保存" )
		self.layoutfile_open_button = wx.Button( self.left_panel, wx.ID_ANY, "レイアウトファイルを読み込む" )

		self.layoutfile_save_button.SetToolTip( "宛名に使用するフォント、郵便番号や住所氏名の配置、はがきや長形3号といった用紙の種別＆サイズ、敬称の数の現時点での設定値を、レイアウトファイルとして保存します" )
		self.layoutfile_open_button.SetToolTip( "レイアウトファイルを読み込んでフォント、各パーツの配置、用紙サイズ、敬称が一つか二つかを入植欄に反映します" )

		#バインド
		self.layoutfile_save_button.Bind( wx.EVT_BUTTON, self.layoutfile_save )
		self.layoutfile_open_button.Bind( wx.EVT_BUTTON, self.layoutfile_open )

		#2つのボタンを横一列にまとめる
		self.twin_buttons_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.twin_buttons_sizer.Add( self.layoutfile_save_button, 0, wx.ALL | wx.FIXED_MINSIZE, 4 )
		self.twin_buttons_sizer.Add( self.layoutfile_open_button, 0, wx.ALL | wx.FIXED_MINSIZE, 4 )

		#○○○○○フォントの選択○○○○○

		self.combobox_font = wx.ComboBox( self.left_panel, wx.ID_ANY, "", style = wx.CB_READONLY )
		for fontdata in self.fonts_data:
			self.combobox_font.Append( fontdata[0], fontdata[1] )

		#バインド
		self.combobox_font.Bind( wx.EVT_COMBOBOX, self.change_font )

		#枠（StaticBoxSizer）に入れる
		self.fontchoice_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●フォントの選択●" )
		self.fontchoice_sb_sizer = wx.StaticBoxSizer( self.fontchoice_sbox, wx.VERTICAL )
		self.fontchoice_sb_sizer.Add( self.combobox_font, 1 )
		self.fontchoice_sb_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "　↑ホイールスクロールなどは遅いので、PageDown、PageUpキーのスクロールを推奨します" ) )


		#○○○○○用紙サイズ（はがき、各種封筒）の選択○○○○○
		self.array_paper_size = maybe_list_natsort( list( self.post_card_envelope_standard ) )
		self.combobox_paper_size = wx.ComboBox( self.left_panel, wx.ID_ANY, "タイトルバーの表示", choices = self.array_paper_size, style = wx.CB_READONLY )

		#バインド
		self.combobox_paper_size.Bind( wx.EVT_COMBOBOX, self.change_paper_size )

		#枠（StaticBoxSizer）に入れる
		self.paper_size_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●用紙サイズ(はがき、各種封筒)の選択●" )
		self.paper_size_sizer = wx.StaticBoxSizer( self.paper_size_sbox, wx.VERTICAL )
		self.paper_size_sizer.Add( self.combobox_paper_size, 0, wx.ALL, 6 )


		#○○○○○参考イメージ画像に、各パーツの最大範囲を示す枠を表示するかどうか○○○○○
		self.checkbox_sample_image_areaframe = wx.CheckBox( self.left_panel, wx.ID_ANY, "参考画像に各パーツを収める範囲を示す枠を表示する" )
		self.checkbox_sample_image_areaframe.SetToolTip( "この枠は表示するだけです。Officeソフトのようにマウスで掴んで移動させたり、枠線を引っ張って大きさを変えることはできません" )

		#バインド
		self.checkbox_sample_image_areaframe.Bind( wx.EVT_CHECKBOX, self.change_sample_image_areaframe )

		#枠（StaticBoxSizer）に入れる
		self.areaframe_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●枠の表示・非表示●" )
		self.areaframe_sb_sizer = wx.StaticBoxSizer( self.areaframe_sbox, wx.VERTICAL )
		self.areaframe_sb_sizer.Add( self.checkbox_sample_image_areaframe, 0, wx.LEFT, 10 )


		#○○○○○郵便番号○○○○○

		#郵便番号の起点となる位置
		self.postalcode_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.postalcode_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.postalcode_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.postalcode_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "左上から右に" ) )
		self.postalcode_position_sizer.Add( self.postalcode_position_x )
		self.postalcode_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.postalcode_position_sizer.Add( self.postalcode_position_y )
		self.postalcode_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#郵便番号の位置の入力欄をバインド
		self.postalcode_position_x.Bind( wx.EVT_SPINCTRL, self.send_postalcode_position_x )
		self.postalcode_position_y.Bind( wx.EVT_SPINCTRL, self.send_postalcode_position_y )

		self.pc_array_horizontal = ( "左", "右", "左右は中央" )
		self.postalcode_direction_horizontal = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.pc_array_horizontal, style = wx.CB_READONLY )

		self.pc_array_vertical = ( "上", "下", "上下は中央" )
		self.postalcode_direction_vertical = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.pc_array_vertical, style = wx.CB_READONLY )

		#1行にまとめる
		self.postalcode_direction_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.postalcode_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置の基準点から" ) )
		self.postalcode_direction_sizer.Add( self.postalcode_direction_horizontal )
		self.postalcode_direction_sizer.Add( self.postalcode_direction_vertical )
		self.postalcode_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "に配置" ) )

		#郵便番号の方向をバインド
		self.postalcode_direction_horizontal.Bind( wx.EVT_COMBOBOX, self.send_postalcode_direction_horizontal )
		self.postalcode_direction_vertical.Bind( wx.EVT_COMBOBOX, self.send_postalcode_direction_vertical )

		#宛先郵便番号の一字あたりのサイズ
		self.postalcode_letterwidth = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.postalcode_letterheight = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.postalcode_lettersize_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.postalcode_lettersize_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "一文字あたり、幅" ) )
		self.postalcode_lettersize_sizer.Add( self.postalcode_letterwidth )
		self.postalcode_lettersize_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ" ) )
		self.postalcode_lettersize_sizer.Add( self.postalcode_letterheight )
		self.postalcode_lettersize_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mmに収める" ) )

		#宛先郵便番号の一字あたりのサイズをバインド
		self.postalcode_letterwidth.Bind( wx.EVT_SPINCTRL, self.send_postalcode_letterwidth )
		self.postalcode_letterheight.Bind( wx.EVT_SPINCTRL, self.send_postalcode_letterheight )

		#宛先郵便番号における、各番号の位置
		self.postalcode_placement2 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.postalcode_placement3 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.postalcode_placement4 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.postalcode_placement5 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.postalcode_placement6 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.postalcode_placement7 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )

		#1行にまとめる
		self.postalcode_placement_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.postalcode_placement_sizer.Add( self.postalcode_placement2 )
		self.postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.postalcode_placement_sizer.Add( self.postalcode_placement3 )
		self.postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.postalcode_placement_sizer.Add( self.postalcode_placement4 )
		self.postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.postalcode_placement_sizer.Add( self.postalcode_placement5 )
		self.postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.postalcode_placement_sizer.Add( self.postalcode_placement6 )
		self.postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.postalcode_placement_sizer.Add( self.postalcode_placement7 )
		self.postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#郵便番号の配置の入力欄をバインド
		self.postalcode_placement2.Bind( wx.EVT_SPINCTRL, self.send_postalcode_placement2 )
		self.postalcode_placement3.Bind( wx.EVT_SPINCTRL, self.send_postalcode_placement3 )
		self.postalcode_placement4.Bind( wx.EVT_SPINCTRL, self.send_postalcode_placement4 )
		self.postalcode_placement5.Bind( wx.EVT_SPINCTRL, self.send_postalcode_placement5 )
		self.postalcode_placement6.Bind( wx.EVT_SPINCTRL, self.send_postalcode_placement6 )
		self.postalcode_placement7.Bind( wx.EVT_SPINCTRL, self.send_postalcode_placement7 )

		#枠（StaticBoxSizer）に入れる
		self.postalcode_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●宛先の郵便番号●" )
		self.postalcode_sb_sizer = wx.StaticBoxSizer( self.postalcode_sbox, wx.VERTICAL )
		self.postalcode_sb_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "○郵便番号全体の位置" ), 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.postalcode_sb_sizer.Add( self.postalcode_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.postalcode_sb_sizer.Add( self.postalcode_direction_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.postalcode_sb_sizer.Add( self.postalcode_lettersize_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.postalcode_sb_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "○郵便番号内の配置" ), 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.postalcode_sb_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "左端の番号の中心から" ), 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.postalcode_sb_sizer.Add( self.postalcode_placement_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#○○○○○宛先氏名欄○○○○○

		#宛先氏名の起点となる位置
		self.destination_name_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_name_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_name_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_name_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置：左上から右に" ) )
		self.destination_name_position_sizer.Add( self.destination_name_position_x )
		self.destination_name_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.destination_name_position_sizer.Add( self.destination_name_position_y )
		self.destination_name_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#宛先氏名の位置をバインド
		self.destination_name_position_x.Bind( wx.EVT_SPINCTRL, self.send_destination_name_position_x )
		self.destination_name_position_y.Bind( wx.EVT_SPINCTRL, self.send_destination_name_position_y )

		self.dn_array_horizontal = ( "左", "右", "左右は中央" )
		self.destination_name_direction_horizontal = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.dn_array_horizontal, style = wx.CB_READONLY )

		self.dn_array_vertical = ( "上", "下", "上下は中央" )
		self.destination_name_direction_vertical = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.dn_array_vertical, style = wx.CB_READONLY )

		#1行にまとめる
		self.destination_name_direction_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_name_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置の基準点から" ) )
		self.destination_name_direction_sizer.Add( self.destination_name_direction_horizontal )
		self.destination_name_direction_sizer.Add( self.destination_name_direction_vertical )
		self.destination_name_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "に配置" ) )

		#宛先氏名の方向をバインド
		self.destination_name_direction_horizontal.Bind( wx.EVT_COMBOBOX, self.send_destination_name_direction_horizontal )
		self.destination_name_direction_vertical.Bind( wx.EVT_COMBOBOX, self.send_destination_name_direction_vertical )

		#宛先氏名を書く領域のサイズ上限
		self.destination_name_size_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_name_size_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_name_size_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_name_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "宛先氏名の幅：" ) )
		self.destination_name_size_sizer.Add( self.destination_name_size_x )
		self.destination_name_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ：" ) )
		self.destination_name_size_sizer.Add( self.destination_name_size_y )
		self.destination_name_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm 以内" ) )

		#宛先氏名の位置の入力欄をバインド
		self.destination_name_size_x.Bind( wx.EVT_SPINCTRL, self.send_destination_name_size_x )
		self.destination_name_size_y.Bind( wx.EVT_SPINCTRL, self.send_destination_name_size_y )
		self.default_destination_name_position = self.image_generator.get_parts_data( "name-position" )

		#宛先氏名の二列の間隔
		self.destination_name_space = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )

		#1行にまとめる
		self.destination_name_space_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_name_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "宛先氏名の二列の間隔：" ) )
		self.destination_name_space_sizer.Add( self.destination_name_space )
		self.destination_name_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#宛先氏名の二列の間隔をバインド
		self.destination_name_space.Bind( wx.EVT_SPINCTRL, self.send_destination_name_space )

		#敬称の選択
		self.dn_array_honorific = ( "敬称は中央に一つ", "二列それぞれに敬称を付ける", "左側にのみ敬称を付ける" )
		self.destination_honorific_mode = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.dn_array_honorific, style = wx.CB_READONLY )

		#1行にまとめる
		self.honorific_attention_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.honorific_attention_sizer.Add( self.destination_honorific_mode, 0 )
		self.honorific_attention_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "※枠は敬称込みなので、↑この高さより大きくなります" ), wx.LEFT, 10 )

		#宛先氏名の敬称をバインド
		self.destination_honorific_mode.Bind( wx.EVT_COMBOBOX, self.change_destination_honorific_mode )

		#宛名2人目の配置選択
		self.dn_array_alignment = [ "下寄せ", "上寄せ" ]
		self.destination_twoname_alignment_mode = wx.ComboBox( self.left_panel, wx.ID_ANY, "配置", choices = self.dn_array_alignment, style = wx.CB_READONLY )

		#1行にまとめる
		self.twoname_alignment_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.twoname_alignment_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "宛名2人目の配置：" ), 0 )
		self.twoname_alignment_sizer.Add( self.destination_twoname_alignment_mode, 0, wx.LEFT, 10 )

		#宛名2人目の配置をバインド
		self.destination_twoname_alignment_mode.Bind( wx.EVT_COMBOBOX, self.change_destination_twoname_alignment_mode )


		#枠（StaticBoxSizer）に入れる
		self.destination_name_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●宛先氏名●" )
		self.destination_name_sb_sizer = wx.StaticBoxSizer( self.destination_name_sbox, wx.VERTICAL )
		self.destination_name_sb_sizer.Add( self.destination_name_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_name_sb_sizer.Add( self.destination_name_direction_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_name_sb_sizer.Add( self.destination_name_size_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_name_sb_sizer.Add( self.destination_name_space_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_name_sb_sizer.Add( self.honorific_attention_sizer, 0, wx.LEFT | wx.RIGHT, 10 )
		self.destination_name_sb_sizer.Add( self.twoname_alignment_sizer, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10 )


		#○○○○○宛先住所○○○○○

		#宛先住所の起点となる位置
		self.destination_address_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_address_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_address_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_address_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置：左上から右に" ) )
		self.destination_address_position_sizer.Add( self.destination_address_position_x )
		self.destination_address_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.destination_address_position_sizer.Add( self.destination_address_position_y )
		self.destination_address_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#宛先住所の位置をバインド
		self.destination_address_position_x.Bind( wx.EVT_SPINCTRL, self.send_destination_address_position_x )
		self.destination_address_position_y.Bind( wx.EVT_SPINCTRL, self.send_destination_address_position_y )

		self.da_array_horizontal = ( "左", "右", "左右は中央" )
		self.destination_address_direction_horizontal = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.da_array_horizontal, style = wx.CB_READONLY )

		self.da_array_vertical = ( "上", "下", "上下は中央" )
		self.destination_address_direction_vertical = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.da_array_vertical, style = wx.CB_READONLY )

		#1行にまとめる
		self.destination_address_direction_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_address_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置の基準点から" ) )
		self.destination_address_direction_sizer.Add( self.destination_address_direction_horizontal )
		self.destination_address_direction_sizer.Add( self.destination_address_direction_vertical )
		self.destination_address_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "に配置" ) )

		#宛先住所の方向をバインド
		self.destination_address_direction_horizontal.Bind( wx.EVT_COMBOBOX, self.send_destination_address_direction_horizontal )
		self.destination_address_direction_vertical.Bind( wx.EVT_COMBOBOX, self.send_destination_address_direction_vertical )

		#宛先住所を書く領域のサイズ上限
		self.destination_address_size_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_address_size_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_address_size_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_address_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "宛先住所の幅：" ) )
		self.destination_address_size_sizer.Add( self.destination_address_size_x )
		self.destination_address_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ：" ) )
		self.destination_address_size_sizer.Add( self.destination_address_size_y )
		self.destination_address_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm 以内" ) )

		#宛先住所のサイズの入力欄をバインド
		self.destination_address_size_x.Bind( wx.EVT_SPINCTRL, self.send_destination_address_size_x )
		self.destination_address_size_y.Bind( wx.EVT_SPINCTRL, self.send_destination_address_size_y )

		#宛先住所の二列の間隔
		self.destination_address_space = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )

		#1行にまとめる
		self.destination_address_space_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_address_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "宛先住所の二列の間隔：" ) )
		self.destination_address_space_sizer.Add( self.destination_address_space )
		self.destination_address_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#宛先住所の二列の間隔をバインド
		self.destination_address_space.Bind( wx.EVT_SPINCTRL, self.send_destination_address_space )

		#枠（StaticBoxSizer）に入れる
		self.destination_address_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●宛先住所●" )
		self.destination_address_sb_sizer = wx.StaticBoxSizer( self.destination_address_sbox, wx.VERTICAL )
		self.destination_address_sb_sizer.Add( self.destination_address_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_address_sb_sizer.Add( self.destination_address_direction_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_address_sb_sizer.Add( self.destination_address_size_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_address_sb_sizer.Add( self.destination_address_space_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#○○○○○宛先会社名○○○○○

		#宛先会社名の起点となる位置
		self.destination_company_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_company_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_company_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_company_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置：左上から右に" ) )
		self.destination_company_position_sizer.Add( self.destination_company_position_x )
		self.destination_company_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.destination_company_position_sizer.Add( self.destination_company_position_y )
		self.destination_company_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#宛先会社名の位置をバインド
		self.destination_company_position_x.Bind( wx.EVT_SPINCTRL, self.send_destination_company_position_x )
		self.destination_company_position_y.Bind( wx.EVT_SPINCTRL, self.send_destination_company_position_y )

		#宛先会社名を書く領域のサイズ上限
		self.destination_company_size_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_company_size_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_company_size_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_company_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "会社名の幅：" ) )
		self.destination_company_size_sizer.Add( self.destination_company_size_x )
		self.destination_company_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ：" ) )
		self.destination_company_size_sizer.Add( self.destination_company_size_y )
		self.destination_company_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm 以内" ) )

		#宛先会社名のサイズの入力欄をバインド
		self.destination_company_size_x.Bind( wx.EVT_SPINCTRL, self.send_destination_company_size_x )
		self.destination_company_size_y.Bind( wx.EVT_SPINCTRL, self.send_destination_company_size_y )

		#枠（StaticBoxSizer）に入れる
		self.destination_company_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●宛先会社名●" )
		self.destination_company_sb_sizer = wx.StaticBoxSizer( self.destination_company_sbox, wx.VERTICAL )
		self.destination_company_sb_sizer.Add( self.destination_company_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_company_sb_sizer.Add( self.destination_company_size_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#○○○○○宛先部署名○○○○○

		#宛先部署名の起点となる位置
		self.destination_department_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_department_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_department_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_department_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置：左上から右に" ) )
		self.destination_department_position_sizer.Add( self.destination_department_position_x )
		self.destination_department_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.destination_department_position_sizer.Add( self.destination_department_position_y )
		self.destination_department_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#宛先部署名の位置をバインド
		self.destination_department_position_x.Bind( wx.EVT_SPINCTRL, self.send_destination_department_position_x )
		self.destination_department_position_y.Bind( wx.EVT_SPINCTRL, self.send_destination_department_position_y )

		#宛先部署名を書く領域のサイズ上限
		self.destination_department_size_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.destination_department_size_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.destination_department_size_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_department_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "部署名の幅：" ) )
		self.destination_department_size_sizer.Add( self.destination_department_size_x )
		self.destination_department_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ：" ) )
		self.destination_department_size_sizer.Add( self.destination_department_size_y )
		self.destination_department_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm 以内" ) )

		#宛先部署名のサイズの入力欄をバインド
		self.destination_department_size_x.Bind( wx.EVT_SPINCTRL, self.send_destination_department_size_x )
		self.destination_department_size_y.Bind( wx.EVT_SPINCTRL, self.send_destination_department_size_y )

		#枠（StaticBoxSizer）に入れる
		self.destination_department_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●宛先部署名●" )
		self.destination_department_sb_sizer = wx.StaticBoxSizer( self.destination_department_sbox, wx.VERTICAL )
		self.destination_department_sb_sizer.Add( self.destination_department_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.destination_department_sb_sizer.Add( self.destination_department_size_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#○○○○○差出人の郵便番号○○○○○

		#差出人の郵便番号の起点となる位置
		self.our_postalcode_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_postalcode_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.our_postalcode_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_postalcode_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "左上から右に" ) )
		self.our_postalcode_position_sizer.Add( self.our_postalcode_position_x )
		self.our_postalcode_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.our_postalcode_position_sizer.Add( self.our_postalcode_position_y )
		self.our_postalcode_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#差出人の郵便番号の位置の入力欄をバインド
		self.our_postalcode_position_x.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_position_x )
		self.our_postalcode_position_y.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_position_y )

		self.opc_array_horizontal = ( "左", "右", "左右は中央" )
		self.our_postalcode_direction_horizontal = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.opc_array_horizontal, style = wx.CB_READONLY )

		self.opc_array_vertical = ( "上", "下", "上下は中央" )
		self.our_postalcode_direction_vertical = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.opc_array_vertical, style = wx.CB_READONLY )

		#1行にまとめる
		self.our_postalcode_direction_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_postalcode_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置の基準点から" ) )
		self.our_postalcode_direction_sizer.Add( self.our_postalcode_direction_horizontal )
		self.our_postalcode_direction_sizer.Add( self.our_postalcode_direction_vertical )
		self.our_postalcode_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "に配置" ) )

		#差出人の郵便番号の方向をバインド
		self.our_postalcode_direction_horizontal.Bind( wx.EVT_COMBOBOX, self.send_our_postalcode_direction_horizontal )
		self.our_postalcode_direction_vertical.Bind( wx.EVT_COMBOBOX, self.send_our_postalcode_direction_vertical )

		#差出人郵便番号の一字あたりのサイズ
		self.our_postalcode_letterwidth = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_postalcode_letterheight = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		#1行にまとめる
		self.our_postalcode_lettersize_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_postalcode_lettersize_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "一文字あたり、幅" ) )
		self.our_postalcode_lettersize_sizer.Add( self.our_postalcode_letterwidth )
		self.our_postalcode_lettersize_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ" ) )
		self.our_postalcode_lettersize_sizer.Add( self.our_postalcode_letterheight )
		self.our_postalcode_lettersize_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mmに収める" ) )

		#差出人郵便番号の一字あたりのサイズをバインド
		self.our_postalcode_letterwidth.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_letterwidth )
		self.our_postalcode_letterheight.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_letterheight )


		#差出人郵便番号における、各番号の位置
		self.our_postalcode_placement2 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_postalcode_placement3 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_postalcode_placement4 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_postalcode_placement5 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_postalcode_placement6 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_postalcode_placement7 = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )

		#1行にまとめる
		self.our_postalcode_placement_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_postalcode_placement_sizer.Add( self.our_postalcode_placement2 )
		self.our_postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.our_postalcode_placement_sizer.Add( self.our_postalcode_placement3 )
		self.our_postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.our_postalcode_placement_sizer.Add( self.our_postalcode_placement4 )
		self.our_postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.our_postalcode_placement_sizer.Add( self.our_postalcode_placement5 )
		self.our_postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.our_postalcode_placement_sizer.Add( self.our_postalcode_placement6 )
		self.our_postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、" ) )
		self.our_postalcode_placement_sizer.Add( self.our_postalcode_placement7 )
		self.our_postalcode_placement_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#差出人の郵便番号の配置の入力欄をバインド
		self.our_postalcode_placement2.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_placement2 )
		self.our_postalcode_placement3.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_placement3 )
		self.our_postalcode_placement4.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_placement4 )
		self.our_postalcode_placement5.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_placement5 )
		self.our_postalcode_placement6.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_placement6 )
		self.our_postalcode_placement7.Bind( wx.EVT_SPINCTRL, self.send_our_postalcode_placement7 )

		#枠（StaticBoxSizer）に入れる
		self.our_postalcode_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●宛先の差出人の郵便番号●" )
		self.our_postalcode_sb_sizer = wx.StaticBoxSizer( self.our_postalcode_sbox, wx.VERTICAL )
		self.our_postalcode_sb_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "○差出人の郵便番号全体の位置" ), 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_postalcode_sb_sizer.Add( self.our_postalcode_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_postalcode_sb_sizer.Add( self.our_postalcode_direction_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_postalcode_sb_sizer.Add( self.our_postalcode_lettersize_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_postalcode_sb_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "○差出人の郵便番号内の配置" ), 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_postalcode_sb_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "左端の番号の中心から" ), 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_postalcode_sb_sizer.Add( self.our_postalcode_placement_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#○○○○○差出人氏名欄○○○○○

		#差出人氏名の起点となる位置
		self.our_name_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_name_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.our_name_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_name_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置：左上から右に" ) )
		self.our_name_position_sizer.Add( self.our_name_position_x )
		self.our_name_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.our_name_position_sizer.Add( self.our_name_position_y )
		self.our_name_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#差出人氏名の位置をバインド
		self.our_name_position_x.Bind( wx.EVT_SPINCTRL, self.send_our_name_position_x )
		self.our_name_position_y.Bind( wx.EVT_SPINCTRL, self.send_our_name_position_y )

		self.on_array_horizontal = ( "左", "右", "左右は中央" )
		self.our_name_direction_horizontal = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.on_array_horizontal, style = wx.CB_READONLY )

		self.on_array_vertical = ( "上", "下", "上下は中央" )
		self.our_name_direction_vertical = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.on_array_vertical, style = wx.CB_READONLY )

		#1行にまとめる
		self.our_name_direction_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_name_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置の基準点から" ) )
		self.our_name_direction_sizer.Add( self.our_name_direction_horizontal )
		self.our_name_direction_sizer.Add( self.our_name_direction_vertical )
		self.our_name_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "に配置" ) )

		#差出人氏名の方向をバインド
		self.our_name_direction_horizontal.Bind( wx.EVT_COMBOBOX, self.send_our_name_direction_horizontal )
		self.our_name_direction_vertical.Bind( wx.EVT_COMBOBOX, self.send_our_name_direction_vertical )

		#差出人氏名を書く領域のサイズ上限
		self.our_name_size_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_name_size_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.our_name_size_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_name_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "差出人氏名の幅：" ) )
		self.our_name_size_sizer.Add( self.our_name_size_x )
		self.our_name_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ：" ) )
		self.our_name_size_sizer.Add( self.our_name_size_y )
		self.our_name_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm 以内" ) )

		#差出人氏名のサイズの入力欄をバインド
		self.our_name_size_x.Bind( wx.EVT_SPINCTRL, self.send_our_name_size_x )
		self.our_name_size_y.Bind( wx.EVT_SPINCTRL, self.send_our_name_size_y )

		#差出人氏名の二列の間隔
		self.our_name_space = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )

		#1行にまとめる
		self.our_name_space_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_name_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "差出人氏名の二列の間隔：" ) )
		self.our_name_space_sizer.Add( self.our_name_space )
		self.our_name_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#差出人氏名の二列の間隔をバインド
		self.our_name_space.Bind( wx.EVT_SPINCTRL, self.send_our_name_space )

		#枠（StaticBoxSizer）に入れる
		self.our_name_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●差出人氏名●" )
		self.our_name_sb_sizer = wx.StaticBoxSizer( self.our_name_sbox, wx.VERTICAL )
		self.our_name_sb_sizer.Add( self.our_name_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_name_sb_sizer.Add( self.our_name_direction_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_name_sb_sizer.Add( self.our_name_size_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_name_sb_sizer.Add( self.our_name_space_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#○○○○○差出人住所○○○○○

		#差出人住所の起点となる位置
		self.our_address_position_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_address_position_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.our_address_position_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_address_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置：左上から右に" ) )
		self.our_address_position_sizer.Add( self.our_address_position_x )
		self.our_address_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、下に" ) )
		self.our_address_position_sizer.Add( self.our_address_position_y )
		self.our_address_position_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#差出人住所の位置をバインド
		self.our_address_position_x.Bind( wx.EVT_SPINCTRL, self.send_our_address_position_x )
		self.our_address_position_y.Bind( wx.EVT_SPINCTRL, self.send_our_address_position_y )

		self.oa_array_horizontal = ( "左", "右", "左右は中央" )
		self.our_address_direction_horizontal = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.oa_array_horizontal, style = wx.CB_READONLY )

		self.oa_array_vertical = ( "上", "下", "上下は中央" )
		self.our_address_direction_vertical = wx.ComboBox( self.left_panel, wx.ID_ANY, "方向", choices = self.oa_array_vertical, style = wx.CB_READONLY )

		#1行にまとめる
		self.our_address_direction_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_address_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "位置の基準点から" ) )
		self.our_address_direction_sizer.Add( self.our_address_direction_horizontal )
		self.our_address_direction_sizer.Add( self.our_address_direction_vertical )
		self.our_address_direction_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "に配置" ) )

		#差出人住所の方向をバインド
		self.our_address_direction_horizontal.Bind( wx.EVT_COMBOBOX, self.send_our_address_direction_horizontal )
		self.our_address_direction_vertical.Bind( wx.EVT_COMBOBOX, self.send_our_address_direction_vertical )

		#差出人住所を書く領域のサイズ上限
		self.our_address_size_x = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )
		self.our_address_size_y = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 500, size = ( 110, 30 ) )
		#1行にまとめる
		self.our_address_size_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_address_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "差出人住所の幅：" ) )
		self.our_address_size_sizer.Add( self.our_address_size_x )
		self.our_address_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm、高さ：" ) )
		self.our_address_size_sizer.Add( self.our_address_size_y )
		self.our_address_size_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm 以内" ) )

		#差出人住所のサイズの入力欄をバインド
		self.our_address_size_x.Bind( wx.EVT_SPINCTRL, self.send_our_address_size_x )
		self.our_address_size_y.Bind( wx.EVT_SPINCTRL, self.send_our_address_size_y )

		#差出人住所の二列の間隔
		self.our_address_space = wx.SpinCtrl( self.left_panel, wx.ID_ANY, min = 0, max = 450, size = ( 110, 30 ) )

		#1行にまとめる
		self.our_address_space_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_address_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "差出人住所の二列の間隔：" ) )
		self.our_address_space_sizer.Add( self.our_address_space )
		self.our_address_space_sizer.Add( wx.StaticText( self.left_panel, wx.ID_ANY, "mm" ) )

		#差出人住所の二列の間隔をバインド
		self.our_address_space.Bind( wx.EVT_SPINCTRL, self.send_our_address_space )

		#枠（StaticBoxSizer）に入れる
		self.our_address_sbox = wx.StaticBox( self.left_panel, wx.ID_ANY, "●差出人住所●" )
		self.our_address_sb_sizer = wx.StaticBoxSizer( self.our_address_sbox, wx.VERTICAL )
		self.our_address_sb_sizer.Add( self.our_address_position_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_address_sb_sizer.Add( self.our_address_direction_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_address_sb_sizer.Add( self.our_address_size_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.our_address_sb_sizer.Add( self.our_address_space_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#○○○○○余白○○○○○
		self.default_printer_space = self.column_etc_dictionary[ "printer-space-top,bottom,left,right" ]
		self.print_space_top = wx.SpinCtrl( self.right_panel, wx.ID_ANY, value = str( self.default_printer_space[0] ) , min = 0, max = 450, size = ( 110, 30 ) )

		self.print_space_bottom = wx.SpinCtrl( self.right_panel, wx.ID_ANY, value = str( self.default_printer_space[1] ) , min = 0, max = 450, size = ( 110, 30 ) )

		self.print_space_left = wx.SpinCtrl( self.right_panel, wx.ID_ANY, value = str( self.default_printer_space[2] ) , min = 0, max = 450, size = ( 110, 30 ) )

		self.print_space_right = wx.SpinCtrl( self.right_panel, wx.ID_ANY, value = str( self.default_printer_space[3] ) , min = 0, max = 450, size = ( 110, 30 ) )

		#1行にまとめる
		self.print_space_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.print_space_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "上" ) )
		self.print_space_sizer.Add( self.print_space_top )
		self.print_space_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "、下" ) )
		self.print_space_sizer.Add( self.print_space_bottom )
		self.print_space_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "、左" ) )
		self.print_space_sizer.Add( self.print_space_left )
		self.print_space_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "、右" ) )
		self.print_space_sizer.Add( self.print_space_right )
		self.print_space_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "(単位:mm)" ) )

		#余白幅の入力欄をバインド
		self.print_space_top.Bind( wx.EVT_SPINCTRL, self.send_print_space_top )
		self.print_space_bottom.Bind( wx.EVT_SPINCTRL, self.send_print_space_bottom )
		self.print_space_left.Bind( wx.EVT_SPINCTRL, self.send_print_space_left )
		self.print_space_right.Bind( wx.EVT_SPINCTRL, self.send_print_space_right )

		self.frameprint_button = wx.Button( self.right_panel, wx.ID_ANY, "余白計測用に、枠画像を印刷する" )
		self.frameprint_button.SetToolTip( "ハガキ大の紙に、印刷可能な最大サイズで赤い四角形を印刷します。参考イメージの赤枠とは独立しているので、参考イメージ中で赤枠の位置と大きさを変更してもこちらは変化しません。" )

		#余白調査用の印刷ボタンをバインド
		self.frameprint_button.Bind( wx.EVT_BUTTON, self.print_frame_image )

		#枠（StaticBoxSizer）に入れる
		self.printspace_sbox = wx.StaticBox( self.right_panel, wx.ID_ANY, "印刷時の余白" )
		self.printspace_sb_sizer = wx.StaticBoxSizer( self.printspace_sbox, wx.VERTICAL )
		self.printspace_sb_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "枠だけの印刷をして、測った値を入れる （見本画像の中で、赤い枠で表示されます）" ), 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.printspace_sb_sizer.Add( self.print_space_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.printspace_sb_sizer.Add( self.frameprint_button, 0, wx.ALIGN_CENTER | wx.FIXED_MINSIZE )

		#○○○○○フォントサイズ調整○○○○○
		self.fontsize_scale = wx.SpinCtrl( self.right_panel, wx.ID_ANY, value = str( self.image_generator.get_parts_data( "resize％" )[0] ), min = 50, max = 150, size = ( 110, 30 ) )
		self.fontsize_scale.SetToolTip( "フォントサイズを50%～150%の範囲で調整できます。100%が標準サイズです。" )
		
		#1行にまとめる
		self.fontsize_scale_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.fontsize_scale_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "フォントサイズ：" ) )
		self.fontsize_scale_sizer.Add( self.fontsize_scale )
		self.fontsize_scale_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "%" ) )
		
		#フォントサイズ調整をバインド
		self.fontsize_scale.Bind( wx.EVT_SPINCTRL, self.send_fontsize_scale )
		
		#枠（StaticBoxSizer）に入れる
		self.fontsize_sbox = wx.StaticBox( self.right_panel, wx.ID_ANY, "フォントサイズ調整" )
		self.fontsize_sb_sizer = wx.StaticBoxSizer( self.fontsize_sbox, wx.VERTICAL )
		self.fontsize_sb_sizer.Add( wx.StaticText( self.right_panel, wx.ID_ANY, "文字が大きすぎる／小さすぎる場合に調整してください" ), 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )
		self.fontsize_sb_sizer.Add( self.fontsize_scale_sizer, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10 )


		self.left_sizer = wx.BoxSizer( wx.VERTICAL )
		self.left_sizer.Add( self.twin_buttons_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.fontchoice_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.paper_size_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.areaframe_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.postalcode_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.destination_name_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.destination_address_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.destination_company_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.destination_department_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.our_postalcode_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.our_name_sb_sizer )
		self.left_sizer.Add( wx.StaticLine( self.left_panel ), 0, wx.TOP, 10 )
		self.left_sizer.Add( self.our_address_sb_sizer )
		self.left_panel.SetSizer( self.left_sizer )
		self.left_panel.SetupScrolling()

		self.right_surface_sizer = wx.BoxSizer( wx.VERTICAL )
		self.right_surface_sizer.Add( self.sampleimage_sb_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND )

		if self.column_etc_dictionary[ "upside-down-print" ] is True:
			upside_down_message = "、余白の切り抜きが180°反転中"
		else:
			upside_down_message = ""
		self.redframe_notes = wx.StaticText( self.right_panel, wx.ID_ANY, "※赤枠の内部だけが印刷されます" + upside_down_message )
		self.right_surface_sizer.Add( self.redframe_notes, 0, wx.ALIGN_CENTER_HORIZONTAL )

		self.right_surface_sizer.Add( self.printspace_sb_sizer )
		self.right_surface_sizer.Add( self.fontsize_sb_sizer, 0, wx.TOP | wx.EXPAND, 10 )

		self.right_panel.SetSizer( self.right_surface_sizer )
		self.right_panel.SetupScrolling()
		self.right_sizer = wx.BoxSizer( wx.VERTICAL )
		self.right_sizer.Add( self.right_panel, 1, wx.ALL | wx.EXPAND )

		self.layout_total_sizer = wx.GridSizer( 1, 2, 0, 0 )
		self.layout_total_sizer.Add( self.left_panel, 1, wx.ALL | wx.EXPAND )
		self.layout_total_sizer.Add( self.right_sizer, 1, wx.ALL | wx.EXPAND )

		self.layout_tab_panel.SetSizer( self.layout_total_sizer )


		#■■■■■差出人記述・項目と列の対応・印刷可否・敬称等のパネル■■■■■
		self.initial_columns = self.grid.GetNumberCols() + 1

		#宛先の各項目に何列目の内容を割り当てるか
		self.dest_postalcode_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-postalcode", "0" ) + 1 ) , min = 1, max = 15, size = ( 110, 30 )  )
		self.dest_address1_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-address1", "0" ) + 1 ) , min = 1, max = 15, size = ( 110, 30 )  )
		self.dest_address2_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-address2", "0" ) + 1 ) , min = 1, max = 15, size = ( 110, 30 )  )
		self.dest_name1_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-name1", "0" ) + 1 ) , min = 1, max = 15, size = ( 110, 30 )  )
		self.dest_name2_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-name2", "0" ) + 1 ) , min = 1, max = 15, size = ( 110, 30 )  )
		self.dest_company_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-company", "0" ) + 1 ) , min = 1, max = 15, size = ( 110, 30 )  )
		self.dest_department_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-department", "0" ) + 1 ) , min = 1, max = 15, size = ( 110, 30 )  )

		#バインド
		self.dest_postalcode_column.Bind( wx.EVT_SPINCTRL, self.send_dest_postalcode_column )
		self.dest_address1_column.Bind( wx.EVT_SPINCTRL, self.send_dest_address1_column )
		self.dest_address2_column.Bind( wx.EVT_SPINCTRL, self.send_dest_address2_column )
		self.dest_name1_column.Bind( wx.EVT_SPINCTRL, self.send_dest_name1_column )
		self.dest_name2_column.Bind( wx.EVT_SPINCTRL, self.send_dest_name2_column )
		self.dest_company_column.Bind( wx.EVT_SPINCTRL, self.send_dest_company_column )
		self.dest_department_column.Bind( wx.EVT_SPINCTRL, self.send_dest_department_column )

		#1行……だと長過ぎてウィンドウに収まらないので2行にまとめる
		self.destination_column_sizer1 = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_column_sizer1.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "宛先郵便番号：" ) )
		self.destination_column_sizer1.Add( self.dest_postalcode_column )
		self.destination_column_sizer1.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "、　宛先住所：" ) )
		self.destination_column_sizer1.Add( self.dest_address1_column )
		self.destination_column_sizer1.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "、　宛先住所（2列め）：" ) )
		self.destination_column_sizer1.Add( self.dest_address2_column )

		self.destination_column_sizer2 = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_column_sizer2.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "宛先氏名：" ) )
		self.destination_column_sizer2.Add( self.dest_name1_column )
		self.destination_column_sizer2.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "、　宛先氏名（2列め）：" ) )
		self.destination_column_sizer2.Add( self.dest_name2_column )

		self.destination_column_sizer3 = wx.BoxSizer( wx.HORIZONTAL )
		self.destination_column_sizer3.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "会社名：" ) )
		self.destination_column_sizer3.Add( self.dest_company_column )
		self.destination_column_sizer3.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "、　部署：" ) )
		self.destination_column_sizer3.Add( self.dest_department_column )

		#枠（StaticBoxSizer）に入れる
		self.ap_destination_sbox = wx.StaticBox( self.appropriate_tab_panel, wx.ID_ANY, "●表中の何列目を宛先の各項目に割り当てるか●" )
		self.ap_destination_sb_sizer = wx.StaticBoxSizer( self.ap_destination_sbox, wx.VERTICAL )
		self.ap_destination_sb_sizer.Add( self.destination_column_sizer1, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 6 )
		self.ap_destination_sb_sizer.Add( self.destination_column_sizer2, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 6 )
		self.ap_destination_sb_sizer.Add( self.destination_column_sizer3, 1, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 6 )


		#印刷の可否を特定の列の内容で判別する
		self.checkbox_printctrl = wx.CheckBox( self.appropriate_tab_panel, wx.ID_ANY, "" )
		if self.column_etc_dictionary[ "print-control" ] is True:
			self.checkbox_printctrl.SetValue( True )
		else:
			self.checkbox_printctrl.SetValue( False )

		self.printctrl_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "print-control-column", "0" ) + 1 ) , min = 1, max = 20, size = ( 110, 30 )  )

		self.print_sign = wx.TextCtrl( self.appropriate_tab_panel, wx.ID_ANY, self.column_etc_dictionary[ "print-sign" ] )

		self.print_do_or_ignore_array = ( "だけ印刷する", "印刷せず無視する" )
		self.combobox_print_do_or_ignore = wx.ComboBox( self.appropriate_tab_panel, wx.ID_ANY, "該当した場合に印刷するのか無視するのか", choices = self.print_do_or_ignore_array, style = wx.CB_READONLY )
		if self.column_etc_dictionary[ "print-or-ignore" ] == "ignore":
			self.combobox_print_do_or_ignore.SetSelection( 1 )
		else:
			self.combobox_print_do_or_ignore.SetSelection( 0 )

		#checkboxがオフなら、入力欄を無効にする
		if self.checkbox_printctrl.GetValue() is False:
			self.printctrl_column.Disable()
			self.print_sign.Disable()
			self.combobox_print_do_or_ignore.Disable()

		#印刷の可否の入力欄などをバインド
		self.checkbox_printctrl.Bind( wx.EVT_CHECKBOX, self.change_print_control_on_off )
		self.printctrl_column.Bind( wx.EVT_SPINCTRL, self.chage_printctrl_column )
		self.print_sign.Bind( wx.EVT_TEXT, self.input_print_sign )
		self.combobox_print_do_or_ignore.Bind( wx.EVT_COMBOBOX, self.change_print_execute )

		self.ap_printable_sizer1 = wx.BoxSizer( wx.HORIZONTAL )
		self.ap_printable_sizer1.Add( self.checkbox_printctrl )
		self.ap_printable_sizer1.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "印刷の可否を有効にする" ) )
		self.ap_printable_sizer2 = wx.BoxSizer( wx.HORIZONTAL )
		self.ap_printable_sizer2.Add( self.printctrl_column )
		self.ap_printable_sizer2.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "列目が" ) )
		self.ap_printable_sizer2.Add( self.print_sign )
		self.ap_printable_sizer2.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "であった場合に" ) )
		self.ap_printable_sizer2.Add( self.combobox_print_do_or_ignore )

		#枠（StaticBoxSizer）に入れる
		self.printable_sbox = wx.StaticBox( self.appropriate_tab_panel, wx.ID_ANY, "●表中の記述で印刷の可否を分ける●" )
		self.printable_sb_sizer = wx.StaticBoxSizer( self.printable_sbox, wx.VERTICAL )
		self.printable_sb_sizer.Add( self.ap_printable_sizer1, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.printable_sb_sizer.Add( self.ap_printable_sizer2, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#敬称が表中で指定されていない場合に使う、デフォルトの敬称を定める
		self.checkbox_enable_table_honorific = wx.CheckBox( self.appropriate_tab_panel, wx.ID_ANY, "" )
		self.honorific_column = wx.SpinCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.column_etc_dictionary.get( "column-honorific", "0" ) + 1 ) , min = 1, max = 20, size = ( 110, 30 )  )
		if self.column_etc_dictionary[ "enable-honorific-in-table" ] is True:
			self.checkbox_enable_table_honorific.SetValue( True )
			self.honorific_column.Enable()
		else:
			self.checkbox_enable_table_honorific.SetValue( False )
			self.honorific_column.Disable()

		#1行にまとめる
		self.enable_table_honorific_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.enable_table_honorific_sizer.Add( self.checkbox_enable_table_honorific )
		self.enable_table_honorific_sizer.Add( self.honorific_column )
		self.enable_table_honorific_sizer.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "列目の記述を敬称として使用する" ) )

		#バインド
		self.checkbox_enable_table_honorific.Bind( wx.EVT_CHECKBOX, self.change_enable_table_honorific )
		self.honorific_column.Bind( wx.EVT_SPINCTRL, self.send_honorific_column )

		self.checkbox_enable_default_honorific = wx.CheckBox( self.appropriate_tab_panel, wx.ID_ANY, "" )
		if self.column_etc_dictionary[ "enable-default-honorific" ] is True:
			self.checkbox_enable_default_honorific.SetValue( True )
		else:
			self.checkbox_enable_default_honorific.SetValue( False )

		self.default_honorific = wx.TextCtrl( self.appropriate_tab_panel, wx.ID_ANY, self.column_etc_dictionary[ "default-honorific" ] )

		#checkboxがオフなら、入力欄を無効にする
		if self.checkbox_enable_default_honorific.GetValue() is False:
			self.default_honorific.Disable()

		#デフォルトの敬称を定めるチェックボックスと入力欄をBind
		self.checkbox_enable_default_honorific.Bind( wx.EVT_CHECKBOX, self.change_enable_default_honorific )
		self.default_honorific.Bind( wx.EVT_KILL_FOCUS, self.input_default_honorific )

		#1行にまとめる
		self.default_honorific_sizer1 = wx.BoxSizer( wx.HORIZONTAL )
		self.default_honorific_sizer1.Add( self.checkbox_enable_default_honorific )
		self.default_honorific_sizer1.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "敬称が指定されていない場合、以下の敬称を使用する （列指定が無効な場合と、表中で敬称の欄が空の場合、両方に適用されます）" ) )
		self.default_honorific_sizer2 = wx.BoxSizer( wx.HORIZONTAL )
		self.default_honorific_sizer2.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "デフォルトの敬称 ： " ) )
		self.default_honorific_sizer2.Add( self.default_honorific )

		#枠（StaticBoxSizer）に入れる
		self.default_honorific_sbox = wx.StaticBox( self.appropriate_tab_panel, wx.ID_ANY, "●敬称●" )
		self.default_honorific_sb_sizer = wx.StaticBoxSizer( self.default_honorific_sbox, wx.VERTICAL )
		self.default_honorific_sb_sizer.Add( self.enable_table_honorific_sizer, 1, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )
		self.default_honorific_sb_sizer.Add( self.default_honorific_sizer1, 1, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )
		self.default_honorific_sb_sizer.Add( self.default_honorific_sizer2, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )


		#差出人側の記入内容
		self.our_postalcode_inputbox = TextCtrlPlus( self.appropriate_tab_panel, wx.ID_ANY, str( self.our_data.get( "our-postalcode-data", "" ) ) )
		self.our_postalcode_error_text = wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "" )
		self.our_address1 = wx.TextCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.our_data.get( "our-address1-data", "" ) )  )
		self.our_address2 = wx.TextCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.our_data.get( "our-address2-data", "" ) )  )
		self.our_name1 = wx.TextCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.our_data.get( "our-name1-data", "" ) )  )
		self.our_name2 = wx.TextCtrl( self.appropriate_tab_panel, wx.ID_ANY, value = str( self.our_data.get( "our-name2-data", "" ) )  )

		#バインド
		self.our_postalcode_inputbox.Bind( wx.EVT_KILL_FOCUS, self.check_our_postalcode )
		self.our_address1.Bind( wx.EVT_KILL_FOCUS, self.change_our_address1 )
		self.our_address2.Bind( wx.EVT_KILL_FOCUS, self.change_our_address2 )
		self.our_name1.Bind( wx.EVT_KILL_FOCUS, self.change_our_name1 )
		self.our_name2.Bind( wx.EVT_KILL_FOCUS, self.change_our_name2 )

		#項目の説明と共に1行にまとめる
		self.our_postalcode_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_postalcode_sizer.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "差出人郵便番号 （半角英数字7つ）：" ) )
		self.our_postalcode_sizer.Add( self.our_postalcode_inputbox )
		self.our_postalcode_sizer.Add( self.our_postalcode_error_text )

		self.our_address1_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_address1_sizer.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "差出人住所：" ) )
		self.our_address1_sizer.Add( self.our_address1, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		self.our_address2_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_address2_sizer.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "差出人住所（2列目）：" ) )
		self.our_address2_sizer.Add( self.our_address2, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		self.our_name1_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_name1_sizer.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "差出人氏名：" ) )
		self.our_name1_sizer.Add( self.our_name1, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		self.our_name2_sizer = wx.BoxSizer( wx.HORIZONTAL )
		self.our_name2_sizer.Add( wx.StaticText( self.appropriate_tab_panel, wx.ID_ANY, "差出人氏名（2列目）：" ) )
		self.our_name2_sizer.Add( self.our_name2, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		#枠（StaticBoxSizer）に入れる
		self.ap_our_sbox = wx.StaticBox( self.appropriate_tab_panel, wx.ID_ANY, "●差出人側の記入内容● （「印刷レイアウト」タブのサンプル画像に反映されるので、そちらで確認できます）" )
		self.ap_our_sb_sizer = wx.StaticBoxSizer( self.ap_our_sbox, wx.VERTICAL )
		self.ap_our_sb_sizer.Add( self.our_postalcode_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( wx.StaticLine( self.appropriate_tab_panel ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( self.our_address1_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( wx.StaticLine( self.appropriate_tab_panel ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( self.our_address2_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( wx.StaticLine( self.appropriate_tab_panel ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( self.our_name1_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( wx.StaticLine( self.appropriate_tab_panel ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		self.ap_our_sb_sizer.Add( self.our_name2_sizer, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		#並べてパネルに貼る
		self.appropriate_tab_sizer = wx.BoxSizer( wx.VERTICAL )
		self.appropriate_tab_sizer.Add( self.ap_destination_sb_sizer, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.appropriate_tab_sizer.Add( self.printable_sb_sizer, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 10 )
		self.appropriate_tab_sizer.Add( self.default_honorific_sb_sizer, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 10 )
		self.appropriate_tab_sizer.Add( self.ap_our_sb_sizer, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		self.appropriate_tab_panel.SetSizer( self.appropriate_tab_sizer )
		self.appropriate_tab_panel.SetupScrolling()


		#■■■■■ソフトウェア設定のパネル■■■■■
		self.checkbox_window_maximize = wx.CheckBox( self.setting_tab_panel, wx.ID_ANY, "起動時に最大化する" )
		if self.software_setting[ "window_maximize" ] is True:
			self.checkbox_window_maximize.SetValue( True )
		else:
			self.checkbox_window_maximize.SetValue( False )

		self.size_enter_x = wx.SpinCtrl( self.setting_tab_panel, wx.ID_ANY, value = str( self.software_setting[ "window_size" ][0] ), min = 0, max = wx.DisplaySize()[0]  )
		self.size_enter_y = wx.SpinCtrl( self.setting_tab_panel, wx.ID_ANY, value = str( self.software_setting[ "window_size" ][1] ), min = 0, max = wx.DisplaySize()[1]  )

		#1行にまとめる
		textline_winsize = wx.BoxSizer( wx.HORIZONTAL )
		textline_winsize.Add( wx.StaticText( self.setting_tab_panel, wx.ID_ANY, "ウィンドウサイズ　幅：" ) )
		textline_winsize.Add( self.size_enter_x )
		textline_winsize.Add( wx.StaticText( self.setting_tab_panel, wx.ID_ANY, "　、　高さ：" ) )
		textline_winsize.Add( self.size_enter_y )

		#バインド
		self.checkbox_window_maximize.Bind( wx.EVT_CHECKBOX, self.send_window_maximize )
		self.size_enter_x.Bind( wx.EVT_KILL_FOCUS, self.send_window_size_x )
		self.size_enter_y.Bind( wx.EVT_KILL_FOCUS, self.send_window_size_y )

		#枠（StaticBoxSizer）に入れる
		self.window_mode_size_sbox = wx.StaticBox( self.setting_tab_panel, wx.ID_ANY, "●起動時のウィンドウの最大化、通常ウィンドウ時のサイズ●" )
		self.window_mode_size_sizer = wx.StaticBoxSizer( self.window_mode_size_sbox, wx.VERTICAL )
		self.window_mode_size_sizer.Add( self.checkbox_window_maximize, 1, wx.ALL | wx.EXPAND, 10 )
		self.window_mode_size_sizer.Add( textline_winsize, 1, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 10 )
		self.window_mode_size_sizer.Add( wx.StaticText( self.setting_tab_panel, wx.ID_ANY, "※特にこの最大化とウィンドウサイズは、設定後に保存ボタンを押しておかないと意味がありません）" ), 1, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 10 )

		#表のフォントとフォントサイズ

		#表のフォント
		self.combobox_table_font = wx.ComboBox( self.setting_tab_panel, wx.ID_ANY, "", style = wx.CB_READONLY )
		for fontdata in self.table_fontlist:
			self.combobox_table_font.Append( fontdata[0] )

		#初期表示設定
		initial_tablefont = self.software_setting[ "table-font" ]
		self.tablefont_name_list = [ x[0] for x in self.table_fontlist ]

		if initial_tablefont in self.tablefont_name_list:
			self.combobox_table_font.SetSelection( self.tablefont_name_list.index( initial_tablefont ) )

		#説明とつなげて1行にまとめる
		textline_table_font = wx.BoxSizer( wx.HORIZONTAL )
		textline_table_font.Add( wx.StaticText( self.setting_tab_panel, wx.ID_ANY, "表のフォント：" ) )
		textline_table_font.Add( self.combobox_table_font )

		#バインド
		self.combobox_table_font.Bind( wx.EVT_COMBOBOX, self.change_table_font )

		#表のフォントサイズ
		self.table_fontsize_input = wx.SpinCtrl( self.setting_tab_panel, wx.ID_ANY, value = str( self.software_setting[ "table-fontsize" ] ) , min = 0, max = 50 )
		self.table_fontsize_input.SetValue( self.grid.GetDefaultCellFont().GetPointSize() )

		#説明とつなげて1行にまとめる
		textline_table_fontsize = wx.BoxSizer( wx.HORIZONTAL )
		textline_table_fontsize.Add( wx.StaticText( self.setting_tab_panel, wx.ID_ANY, "表のフォントのサイズ：" ) )
		textline_table_fontsize.Add( self.table_fontsize_input )

		#枠（StaticBoxSizer）に入れる
		self.csv_table_font_sbox = wx.StaticBox( self.setting_tab_panel, wx.ID_ANY, "●表のフォント●" )
		self.csv_table_font_sizer = wx.StaticBoxSizer( self.csv_table_font_sbox, wx.VERTICAL )
		self.csv_table_font_sizer.Add( textline_table_font, 1, wx.ALL | wx.EXPAND, 10 )
		self.csv_table_font_sizer.Add( textline_table_fontsize, 1, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 10 )

		#バインド
		self.table_fontsize_input.Bind( wx.EVT_SPINCTRL, self.change_table_fontsize )

		#タイトルバーに、読み込んだCSVファイルの情報（ファイル名かパス）を表示するかどうか
		self.array_titlebar_mode = ( "ファイル名を表示する", "パスを表示する", "ファイルの情報を表示しない" )
		self.combobox_titlebar_mode = wx.ComboBox( self.setting_tab_panel, wx.ID_ANY, "タイトルバーの表示", choices = self.array_titlebar_mode, style = wx.CB_READONLY )
		if self.software_setting[ "write-fileinfo-on-titlebar" ] == "filename":
			self.combobox_titlebar_mode.SetSelection(0)
		elif self.software_setting[ "write-fileinfo-on-titlebar" ] == "filepath":
			self.combobox_titlebar_mode.SetSelection(1)
		else:
			self.combobox_titlebar_mode.SetSelection(2)

		#1行にまとめる
		textline_titlebar = wx.BoxSizer( wx.HORIZONTAL )
		textline_titlebar.Add( wx.StaticText( self.setting_tab_panel, wx.ID_ANY, "タイトルバーに" ) )
		textline_titlebar.Add( self.combobox_titlebar_mode )

		#バインド
		self.combobox_titlebar_mode.Bind( wx.EVT_COMBOBOX, self.change_titlebar_mode )

		#枠（StaticBoxSizer）に入れる
		self.titlebar_mode_sbox = wx.StaticBox( self.setting_tab_panel, wx.ID_ANY, "●タイトルバーにCSVファイルの情報（ファイル名、パス）を表示するかどうか●" )
		self.titlebar_mode_sizer = wx.StaticBoxSizer( self.titlebar_mode_sbox, wx.VERTICAL )
		self.titlebar_mode_sizer.Add( textline_titlebar, 1, wx.ALL | wx.EXPAND, 10 )

		#設定を保存するボタン
		self.save_settings_button = wx.Button( self.setting_tab_panel, wx.ID_ANY, "レイアウト、その他の設定を設定ファイル(iniファイル)に保存する", size=( 500,60 ) )

		#バインド
		self.save_settings_button.Bind( wx.EVT_BUTTON, self.save_settings )

		self.setting_tab_sizer = wx.BoxSizer( wx.VERTICAL )
		self.setting_tab_sizer.Add( self.window_mode_size_sizer, 0, wx.ALL | wx.EXPAND, 10 )
		self.setting_tab_sizer.Add( self.csv_table_font_sizer, 0, wx.ALL | wx.EXPAND, 10 )
		self.setting_tab_sizer.Add( self.titlebar_mode_sizer, 0, wx.ALL | wx.EXPAND, 10 )
		self.setting_tab_sizer.Add( self.save_settings_button, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM | wx.FIXED_MINSIZE )

		self.setting_tab_panel.SetSizer( self.setting_tab_sizer )
		self.setting_tab_panel.SetupScrolling()


		#ウィンドウを配置し終わったところで、レイアウトタブの左半分の初期値設定
		self.layout_widgets_initialize()

		#最後にパネルにサンプルイメージを表示しておく
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

		#レイアウトを確定させる（起動時のボタン重なり防止）
		self.atena_tab_panel.Layout()
		self.Layout()
		self.Refresh()


	#レイアウトタブの左半分のフォントや配置の入力欄の初期値を決める関数
	#最初は各ウィジェットを作成するのと同時にしていたが、レイアウトファイルの読み込みで
	#再初期化する必要が生じたので、ここにまとめて関数化する
	#以前作った似たような関数「set_initial_values_for_layout_input_widgets」もこちらに統合する
	def layout_widgets_initialize( self ):
		#○○○○○フォントの選択部分を初期化する○○○○○
		initial_font = self.image_generator.get_parts_data( "fontfile" )
		font_name_list = [ x[0] for x in self.fonts_data ]

		if os.path.isfile( initial_font ):
			initial_fontname = os.path.splitext( os.path.basename( initial_font ) )[0]
		else:
			initial_fontname = os.path.splitext( initial_font )[0]
			#もし設定辞書中のフォントの記載がファイルのパスでなく、取得したフォント名一覧にあったら
			#宛名画像作成オブジェクト内の辞書の記述を一覧から検索したパスに書き換えておく
			if initial_fontname in font_name_list:
				self.image_generator.set_parts_data( "fontfile", self.fonts_data[ font_name_list.index( initial_fontname ) ][1] )

		if initial_fontname in font_name_list:
			self.combobox_font.SetSelection( font_name_list.index( initial_fontname ) )

		elif os.path.isfile( initial_font ):
			self.fonts_data.insert( 0, [ initial_fontname, initial_font ] )
			#元のリストの内容が変わったので。コンボボックスをまっさらにして追加しなおし
			self.combobox_font.Clear()
			for fontdata in self.fonts_data:
				self.combobox_font.Append( fontdata[0], fontdata[1] )
			self.combobox_font.SetSelection( 0 )

		#○○○○○用紙サイズ（はがき、各種封筒）の選択を初期化する○○○○○
		papersize_mode = self.paper_size_data[ "category" ] + "," + str( self.paper_size_data[ "width" ] ) + "x" + str( self.paper_size_data[ "height" ] )
		if papersize_mode in self.array_paper_size:
			self.combobox_paper_size.SetStringSelection( papersize_mode )

		#○○○○○参考イメージ画像に、各パーツの最大範囲を示す枠を表示するかどうかを初期化○○○○○
		if self.column_etc_dictionary[ "sampleimage-areaframe" ] is True:
			self.checkbox_sample_image_areaframe.SetValue( True )
		else:
			self.checkbox_sample_image_areaframe.SetValue( False )

		#○○○○○郵便番号の位置などの入力を初期化する○○○○○
		self.default_postalcode_position = self.image_generator.get_parts_data( "postalcode-position" )
		self.postalcode_position_x.SetValue( str( self.default_postalcode_position[0] ) )
		self.postalcode_position_y.SetValue( str( self.default_postalcode_position[1] ) )

		self.default_postalcode_lettersize = self.image_generator.get_parts_data( "postalcode-letter-areasize" )
		self.postalcode_letterwidth.SetValue( str( self.default_postalcode_lettersize[0] ) )
		self.postalcode_letterheight.SetValue( str( self.default_postalcode_lettersize[1] ) )

		self.default_postalcode_placement = self.image_generator.get_parts_data( "postalcode-placement" )
		self.postalcode_placement2.SetValue( str( self.default_postalcode_placement[0] ) )
		self.postalcode_placement3.SetValue( str( self.default_postalcode_placement[1] ) )
		self.postalcode_placement4.SetValue( str( self.default_postalcode_placement[2] ) )
		self.postalcode_placement5.SetValue( str( self.default_postalcode_placement[3] ) )
		self.postalcode_placement6.SetValue( str( self.default_postalcode_placement[4] ) )
		self.postalcode_placement7.SetValue( str( self.default_postalcode_placement[5] ) )

		self.default_postalcode_direction = self.image_generator.get_parts_data( "postalcode-direction" )

		if self.default_postalcode_direction[0] == "left":
			self.postalcode_direction_horizontal.SetSelection( 0 )
		elif self.default_postalcode_direction[0] == "right":
			self.postalcode_direction_horizontal.SetSelection( 1 )
		else:
			self.postalcode_direction_horizontal.SetSelection( 2 )

		if self.default_postalcode_direction[1] == "up":
			self.postalcode_direction_vertical.SetSelection( 0 )
		elif self.default_postalcode_direction[1] == "down":
			self.postalcode_direction_vertical.SetSelection( 1 )
		else:
			self.postalcode_direction_vertical.SetSelection( 2 )

		#○○○○○宛先氏名欄の位置などの入力を初期化する○○○○○

		self.default_destination_name_position = self.image_generator.get_parts_data( "name-position" )
		self.destination_name_position_x.SetValue( str( self.default_destination_name_position[0] ) )
		self.destination_name_position_y.SetValue( str( self.default_destination_name_position[1] ) )

		self.default_destination_name_size = self.image_generator.get_parts_data( "name-areasize" )
		self.destination_name_size_x.SetValue( str( self.default_destination_name_size[0] ) )
		self.destination_name_size_y.SetValue( str( self.default_destination_name_size[1] ) )

		self.default_destination_name_space = self.image_generator.get_parts_data( "name-bind-space" )
		self.destination_name_space.SetValue( str( self.default_destination_name_space ) )

		self.default_destination_name_direction = self.image_generator.get_parts_data( "name-direction" )

		if self.default_destination_name_direction[0] == "left":
			self.destination_name_direction_horizontal.SetSelection( 0 )
		elif self.default_destination_name_direction[0] == "right":
			self.destination_name_direction_horizontal.SetSelection( 1 )
		else:
			self.destination_name_direction_horizontal.SetSelection( 2 )

		if self.default_destination_name_direction[1] == "up":
			self.destination_name_direction_vertical.SetSelection( 0 )
		elif self.default_destination_name_direction[1] == "down":
			self.destination_name_direction_vertical.SetSelection( 1 )
		else:
			self.destination_name_direction_vertical.SetSelection( 2 )

		#敬称の選択を初期化する
		self.default_honorific_mode = self.image_generator.get_parts_data( "twoname-honorific-mode" )

		if self.default_honorific_mode == 1:
			self.destination_honorific_mode.SetSelection(0)
		elif self.default_honorific_mode == 2:
			self.destination_honorific_mode.SetSelection(1)
		else:
			self.destination_honorific_mode.SetSelection(2)

		#宛名2人目の配置選択を初期化する
		self.default_twoname_alignment_mode = self.image_generator.get_parts_data( "twoname-alignment-mode" )

		if self.default_twoname_alignment_mode == "top":
			self.destination_twoname_alignment_mode.SetSelection(1)
		else:
			self.destination_twoname_alignment_mode.SetSelection(0)

		#○○○○○宛先住所の位置などの入力を初期化する○○○○○
		self.default_destination_address_position = self.image_generator.get_parts_data( "address-position" )
		self.destination_address_position_x.SetValue( str( self.default_destination_address_position[0] ) )
		self.destination_address_position_y.SetValue( str( self.default_destination_address_position[1] ) )

		self.default_destination_address_size = self.image_generator.get_parts_data( "address-areasize" )
		self.destination_address_size_x.SetValue( str( self.default_destination_address_size[0] ) )
		self.destination_address_size_y.SetValue( str( self.default_destination_address_size[1] ) )

		self.default_destination_address_space = self.image_generator.get_parts_data( "address-bind-space" )
		self.destination_address_space.SetValue( str( self.default_destination_address_space ) )

		self.default_destination_address_direction = self.image_generator.get_parts_data( "address-direction" )

		if self.default_destination_address_direction[0] == "left":
			self.destination_address_direction_horizontal.SetSelection( 0 )
		elif self.default_destination_address_direction[0] == "right":
			self.destination_address_direction_horizontal.SetSelection( 1 )
		else:
			self.destination_address_direction_horizontal.SetSelection( 2 )

		if self.default_destination_address_direction[1] == "up":
			self.destination_address_direction_vertical.SetSelection( 0 )
		elif self.default_destination_address_direction[1] == "down":
			self.destination_address_direction_vertical.SetSelection( 1 )
		else:
			self.destination_address_direction_vertical.SetSelection( 2 )

		#○○○○○会社名の位置などの入力を初期化する○○○○○
		self.default_destination_company_position = self.image_generator.get_parts_data( "company-position" )
		self.destination_company_position_x.SetValue( str( self.default_destination_company_position[0] ) )
		self.destination_company_position_y.SetValue( str( self.default_destination_company_position[1] ) )

		self.default_destination_company_size = self.image_generator.get_parts_data( "company-areasize" )
		self.destination_company_size_x.SetValue( str( self.default_destination_company_size[0] ) )
		self.destination_company_size_y.SetValue( str( self.default_destination_company_size[1] ) )

		#○○○○○部署名の位置などの入力を初期化する○○○○○
		self.default_destination_department_position = self.image_generator.get_parts_data( "department-position" )
		self.destination_department_position_x.SetValue( str( self.default_destination_department_position[0] ) )
		self.destination_department_position_y.SetValue( str( self.default_destination_department_position[1] ) )

		self.default_destination_department_size = self.image_generator.get_parts_data( "department-areasize" )
		self.destination_department_size_x.SetValue( str( self.default_destination_department_size[0] ) )
		self.destination_department_size_y.SetValue( str( self.default_destination_department_size[1] ) )

		#○○○○○差出人の郵便番号の位置などの入力を初期化する○○○○○
		self.default_our_postalcode_position = self.image_generator.get_parts_data( "our-postalcode-position" )
		self.our_postalcode_position_x.SetValue( str( self.default_our_postalcode_position[0] ) )
		self.our_postalcode_position_y.SetValue( str( self.default_our_postalcode_position[1] ) )

		self.default_our_postalcode_lettersize = self.image_generator.get_parts_data( "our-postalcode-letter-areasize" )
		self.our_postalcode_letterwidth.SetValue( str( self.default_our_postalcode_lettersize[0] ) )
		self.our_postalcode_letterheight.SetValue( str( self.default_our_postalcode_lettersize[1] ) )


		self.default_our_postalcode_placement = self.image_generator.get_parts_data( "our-postalcode-placement" )
		self.our_postalcode_placement2.SetValue( str( self.default_our_postalcode_placement[0] ) )
		self.our_postalcode_placement3.SetValue( str( self.default_our_postalcode_placement[1] ) )
		self.our_postalcode_placement4.SetValue( str( self.default_our_postalcode_placement[2] ) )
		self.our_postalcode_placement5.SetValue( str( self.default_our_postalcode_placement[3] ) )
		self.our_postalcode_placement6.SetValue( str( self.default_our_postalcode_placement[4] ) )
		self.our_postalcode_placement7.SetValue( str( self.default_our_postalcode_placement[5] ) )

		self.default_our_postalcode_direction = self.image_generator.get_parts_data( "our-postalcode-direction" )

		if self.default_our_postalcode_direction[0] == "left":
			self.our_postalcode_direction_horizontal.SetSelection( 0 )
		elif self.default_our_postalcode_direction[0] == "right":
			self.our_postalcode_direction_horizontal.SetSelection( 1 )
		else:
			self.our_postalcode_direction_horizontal.SetSelection( 2 )

		if self.default_our_postalcode_direction[1] == "up":
			self.our_postalcode_direction_vertical.SetSelection( 0 )
		elif self.default_our_postalcode_direction[1] == "down":
			self.our_postalcode_direction_vertical.SetSelection( 1 )
		else:
			self.our_postalcode_direction_vertical.SetSelection( 2 )

		#○○○○○差出人氏名欄の位置などの入力を初期化する○○○○○
		self.default_our_name_position = self.image_generator.get_parts_data( "our-name-position" )
		self.our_name_position_x.SetValue( str( self.default_our_name_position[0] ) )
		self.our_name_position_y.SetValue( str( self.default_our_name_position[1] ) )

		self.default_our_name_size = self.image_generator.get_parts_data( "our-name-areasize" )
		self.our_name_size_x.SetValue( str( self.default_our_name_size[0] ) )
		self.our_name_size_y.SetValue( str( self.default_our_name_size[1] ) )

		self.default_our_name_space = self.image_generator.get_parts_data( "our-name-bind-space" )
		self.our_name_space.SetValue( str( self.default_our_name_space ) )

		self.default_our_name_direction = self.image_generator.get_parts_data( "our-name-direction" )

		if self.default_our_name_direction[0] == "left":
			self.our_name_direction_horizontal.SetSelection( 0 )
		elif self.default_our_name_direction[0] == "right":
			self.our_name_direction_horizontal.SetSelection( 1 )
		else:
			self.our_name_direction_horizontal.SetSelection( 2 )

		if self.default_our_name_direction[1] == "up":
			self.our_name_direction_vertical.SetSelection( 0 )
		elif self.default_our_name_direction[1] == "down":
			self.our_name_direction_vertical.SetSelection( 1 )
		else:
			self.our_name_direction_vertical.SetSelection( 2 )

		#○○○○○差出人住所の位置などの入力を初期化する○○○○○
		self.default_our_address_position = self.image_generator.get_parts_data( "our-address-position" )
		self.our_address_position_x.SetValue( str( self.default_our_address_position[0] ) )
		self.our_address_position_y.SetValue( str( self.default_our_address_position[1] ) )

		self.default_our_address_size = self.image_generator.get_parts_data( "our-address-areasize" )
		self.our_address_size_x.SetValue( str( self.default_our_address_size[0] ) )
		self.our_address_size_y.SetValue( str( self.default_our_address_size[1] ) )

		self.default_our_address_space = self.image_generator.get_parts_data( "our-address-bind-space" )
		self.our_address_space.SetValue( str( self.default_our_address_space ) )

		self.default_our_address_direction = self.image_generator.get_parts_data( "our-address-direction" )

		if self.default_our_address_direction[0] == "left":
			self.our_address_direction_horizontal.SetSelection( 0 )
		elif self.default_our_address_direction[0] == "right":
			self.our_address_direction_horizontal.SetSelection( 1 )
		else:
			self.our_address_direction_horizontal.SetSelection( 2 )

		if self.default_our_address_direction[1] == "up":
			self.our_address_direction_vertical.SetSelection( 0 )
		elif self.default_our_address_direction[1] == "down":
			self.our_address_direction_vertical.SetSelection( 1 )
		else:
			self.our_address_direction_vertical.SetSelection( 2 )


	#デバイスコンテキストを取得してパネルに画像を表示する
	def OnPaint( self, event = None ):
		deviceContext = wx.PaintDC( self.sample_image_panel )
		deviceContext.Clear()
		deviceContext.SetPen( wx.Pen( wx.BLACK, 4 ) )
		deviceContext.DrawBitmap( self.wx_bitmap_image, self.display_position[0], self.display_position[1] )


	#以下、wx.SpinCtrlによる入力欄の転送関数をまとめて書く
	def send_postalcode_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-position", value = self.postalcode_position_x.GetValue(), list_position = 0 )

	def send_postalcode_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-position", value = self.postalcode_position_y.GetValue(), list_position = 1 )

	def send_postalcode_letterwidth( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-letter-areasize", value = self.postalcode_letterwidth.GetValue(), list_position = 0 )

	def send_postalcode_letterheight( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-letter-areasize", value = self.postalcode_letterheight.GetValue(), list_position = 1 )

	def send_postalcode_placement2( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-placement", value = self.postalcode_placement2.GetValue(), list_position = 0 )

	def send_postalcode_placement3( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-placement", value = self.postalcode_placement3.GetValue(), list_position = 1 )

	def send_postalcode_placement4( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-placement", value = self.postalcode_placement4.GetValue(), list_position = 2 )

	def send_postalcode_placement5( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-placement", value = self.postalcode_placement5.GetValue(), list_position = 3 )

	def send_postalcode_placement6( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-placement", value = self.postalcode_placement6.GetValue(), list_position = 4 )

	def send_postalcode_placement7( self, event ):
		self.changedict_image_restructure_int( dict_key = "postalcode-placement", value = self.postalcode_placement7.GetValue(), list_position = 5 )

	def send_destination_name_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "name-position", value = self.destination_name_position_x.GetValue(), list_position = 0 )

	def send_destination_name_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "name-position", value = self.destination_name_position_y.GetValue(), list_position = 1 )

	def send_destination_name_size_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "name-areasize", value = self.destination_name_size_x.GetValue(), list_position = 0 )

	def send_destination_name_size_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "name-areasize", value = self.destination_name_size_y.GetValue(), list_position = 1 )

	def send_destination_name_space( self, event ):
		self.changedict_image_restructure_int( dict_key = "name-bind-space", value = self.destination_name_space.GetValue() )

	def send_destination_address_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "address-position", value = self.destination_address_position_x.GetValue(), list_position = 0 )

	def send_destination_address_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "address-position", value = self.destination_address_position_y.GetValue(), list_position = 1 )

	def send_destination_address_size_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "address-areasize", value = self.destination_address_size_x.GetValue(), list_position = 0 )

	def send_destination_address_size_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "address-areasize", value = self.destination_address_size_y.GetValue(), list_position = 1 )

	def send_destination_address_space( self, event ):
		self.changedict_image_restructure_int( dict_key = "address-bind-space", value = self.destination_address_space.GetValue() )

	def send_our_postalcode_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-position", value = self.our_postalcode_position_x.GetValue(), list_position = 0 )

	def send_our_postalcode_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-position", value = self.our_postalcode_position_y.GetValue(), list_position = 1 )

	def send_our_postalcode_letterwidth( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-letter-areasize", value = self.our_postalcode_letterwidth.GetValue(), list_position = 0 )

	def send_our_postalcode_letterheight( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-letter-areasize", value = self.our_postalcode_letterheight.GetValue(), list_position = 1 )

	def send_our_postalcode_placement2( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-placement", value = self.our_postalcode_placement2.GetValue(), list_position = 0 )

	def send_our_postalcode_placement3( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-placement", value = self.our_postalcode_placement3.GetValue(), list_position = 1 )

	def send_our_postalcode_placement4( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-placement", value = self.our_postalcode_placement4.GetValue(), list_position = 2 )

	def send_our_postalcode_placement5( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-placement", value = self.our_postalcode_placement5.GetValue(), list_position = 3 )

	def send_our_postalcode_placement6( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-placement", value = self.our_postalcode_placement6.GetValue(), list_position = 4 )

	def send_our_postalcode_placement7( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-postalcode-placement", value = self.our_postalcode_placement7.GetValue(), list_position = 5 )

	def send_our_name_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-name-position", value = self.our_name_position_x.GetValue(), list_position = 0 )

	def send_our_name_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-name-position", value = self.our_name_position_y.GetValue(), list_position = 1 )

	def send_our_name_size_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-name-areasize", value = self.our_name_size_x.GetValue(), list_position = 0 )

	def send_our_name_size_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-name-areasize", value = self.our_name_size_y.GetValue(), list_position = 1 )

	def send_our_name_space( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-name-bind-space", value = self.our_name_space.GetValue() )

	def send_our_address_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-address-position", value = self.our_address_position_x.GetValue(), list_position = 0 )

	def send_our_address_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-address-position", value = self.our_address_position_y.GetValue(), list_position = 1 )

	def send_our_address_size_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-address-areasize", value = self.our_address_size_x.GetValue(), list_position = 0 )

	def send_our_address_size_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-address-areasize", value = self.our_address_size_y.GetValue(), list_position = 1 )

	def send_our_address_space( self, event ):
		self.changedict_image_restructure_int( dict_key = "our-address-bind-space", value = self.our_address_space.GetValue() )

	def send_destination_company_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "company-position", value = self.destination_company_position_x.GetValue(), list_position = 0 )

	def send_destination_company_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "company-position", value = self.destination_company_position_y.GetValue(), list_position = 1 )

	def send_destination_company_size_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "company-areasize", value = self.destination_company_size_x.GetValue(), list_position = 0 )

	def send_destination_company_size_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "company-areasize", value = self.destination_company_size_y.GetValue(), list_position = 1 )

	def send_destination_department_position_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "department-position", value = self.destination_department_position_x.GetValue(), list_position = 0 )

	def send_destination_department_position_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "department-position", value = self.destination_department_position_y.GetValue(), list_position = 1 )

	def send_destination_department_size_x( self, event ):
		self.changedict_image_restructure_int( dict_key = "department-areasize", value = self.destination_department_size_x.GetValue(), list_position = 0 )

	def send_destination_department_size_y( self, event ):
		self.changedict_image_restructure_int( dict_key = "department-areasize", value = self.destination_department_size_y.GetValue(), list_position = 1 )

	#以下の余白指定と列の割り当ては、レイアウト辞書内の値ではないので、送る関数が違う
	def send_print_space_top( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "printer-space-top,bottom,left,right", value = self.print_space_top.GetValue(), list_position = 0, make_atena_image = False )

	def send_print_space_bottom( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "printer-space-top,bottom,left,right", value = self.print_space_bottom.GetValue(), list_position = 1, make_atena_image = False )

	def send_print_space_left( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "printer-space-top,bottom,left,right", value = self.print_space_left.GetValue(), list_position = 2, make_atena_image = False )

	def send_print_space_right( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "printer-space-top,bottom,left,right", value = self.print_space_right.GetValue(), list_position = 3, make_atena_image = False )

	def send_fontsize_scale( self, event ):
		scale_value = self.fontsize_scale.GetValue()
		self.image_generator.set_parts_data( "resize％", [ scale_value, scale_value ] )
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#上記の「make_atena_image = False」は、サンプルイメージの更新はするが
	#ハガキ画像の再取得はしない（赤枠合成以降の処理のみやりなおす）

	#以下の「image_reflesh = False」はサンプルイメージの更新そのものをしない
	#（表の何列目を割り当てるかだけでサンプルイメージに関係ないので）

	def send_dest_postalcode_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-postalcode", value = self.dest_postalcode_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	def send_dest_address1_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-address1", value = self.dest_address1_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	def send_dest_address2_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-address2", value = self.dest_address2_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	def send_dest_name1_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-name1", value = self.dest_name1_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	def send_dest_name2_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-name2", value = self.dest_name2_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	def send_dest_company_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-company", value = self.dest_company_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	def send_dest_department_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-department", value = self.dest_department_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	def send_honorific_column( self, event ):
		self.change_columndict_image_restructure_int( dict_key = "column-honorific", value = self.honorific_column.GetValue() - 1, list_position = None, image_reflesh = False )
		self.set_grid_labels()

	#以下のソフトウェア設定の変更もまた、送る関数が違う

	def send_window_maximize( self, event ):
		if self.checkbox_window_maximize.GetValue() is True:
			self.change_setting_dict( dict_key = "window_maximize", value = True )
		else:
			self.change_setting_dict( dict_key = "window_maximize", value = False )

	def send_window_size_x( self, event ):
		self.change_setting_dict( dict_key = "window_size", value = self.size_enter_x.GetValue(), list_position = 0 )

	def send_window_size_y( self, event ):
		self.change_setting_dict( dict_key = "window_size", value = self.size_enter_y.GetValue(), list_position = 1 )


	#以下、comboboxによる方向指定の転送関数をまとめて書く
	def send_postalcode_direction_horizontal( self, event ):
		self.changedict_image_restructure_horizontal( dict_key = "postalcode-direction", value = self.postalcode_direction_horizontal.GetStringSelection() )

	def send_postalcode_direction_vertical( self, event ):
		self.changedict_image_restructure_vertical( dict_key = "postalcode-direction", value = self.postalcode_direction_vertical.GetStringSelection() )

	def send_destination_name_direction_horizontal( self, event ):
		self.changedict_image_restructure_horizontal( dict_key = "name-direction", value = self.destination_name_direction_horizontal.GetStringSelection() )

	def send_destination_name_direction_vertical( self, event ):
		self.changedict_image_restructure_vertical( dict_key = "name-direction", value = self.destination_name_direction_vertical.GetStringSelection() )

	def send_destination_address_direction_horizontal( self, event ):
		self.changedict_image_restructure_horizontal( dict_key = "address-direction", value = self.destination_address_direction_horizontal.GetStringSelection() )

	def send_destination_address_direction_vertical( self, event ):
		self.changedict_image_restructure_vertical( dict_key = "address-direction", value = self.destination_address_direction_vertical.GetStringSelection() )

	def send_our_postalcode_direction_horizontal( self, event ):
		self.changedict_image_restructure_horizontal( dict_key = "our-postalcode-direction", value = self.our_postalcode_direction_horizontal.GetStringSelection() )

	def send_our_postalcode_direction_vertical( self, event ):
		self.changedict_image_restructure_vertical( dict_key = "our-postalcode-direction", value = self.our_postalcode_direction_vertical.GetStringSelection() )

	def send_our_name_direction_horizontal( self, event ):
		self.changedict_image_restructure_horizontal( dict_key = "our-name-direction", value = self.our_name_direction_horizontal.GetStringSelection() )

	def send_our_name_direction_vertical( self, event ):
		self.changedict_image_restructure_vertical( dict_key = "our-name-direction", value = self.our_name_direction_vertical.GetStringSelection() )

	def send_our_address_direction_horizontal( self, event ):
		self.changedict_image_restructure_horizontal( dict_key = "our-address-direction", value = self.our_address_direction_horizontal.GetStringSelection() )

	def send_our_address_direction_vertical( self, event ):
		self.changedict_image_restructure_vertical( dict_key = "our-address-direction", value = self.our_address_direction_vertical.GetStringSelection() )


	#レイアウト辞書[set_parts_data]を変更して、サンプル画像を再構成する
	#（wx.SpinCtrlの整数値をリストに入れる）
	def changedict_image_restructure_int( self, dict_key, value, list_position = None ):
		self.image_generator.set_parts_data( dict_key, value, list_position )
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#設定値辞書[column_etc_dictionary]を変更して、サンプル画像を再構成する
	#（wx.SpinCtrlの整数値をリストに入れる）
	def change_columndict_image_restructure_int( self, dict_key, value, list_position = None, image_reflesh = True, make_atena_image = True ):

		#list_positionは、リストの中の何番目の値を書き換えるかの位置指定のこと。
		#list_positionが渡されていない（None）なら、書き換え対象がリストではないか
		#リストであってもリスト全体をまるごと置き換える、という扱いでvalueを辞書に入れる。
		if list_position is None:
			self.column_etc_dictionary[ dict_key ] = value

		#書き換え対象がListで、list_positionが指定されているなら、リストの一部だけを変更する
		elif isinstance( self.column_etc_dictionary[ dict_key ], list ) and isinstance( list_position, int ):
			self.column_etc_dictionary[ dict_key ][ list_position ] = value

		else:
			return False

		if image_reflesh is True:
			self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )


	#レイアウト辞書を変更して、サンプル画像を再構成する（comboboxの水平方向指定をリストに入れる）
	def changedict_image_restructure_horizontal( self, dict_key, value ):
		if value == "左":
			direction = "left"
		elif value == "右":
			direction = "right"
		else:
			direction = "center"
		self.image_generator.set_parts_data( dict_key, direction, 0 )

		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#レイアウト辞書を変更して、サンプル画像を再構成する（comboboxの垂直方向指定をリストに入れる）
	def changedict_image_restructure_vertical( self, dict_key, value ):
		if value == "上":
			direction = "up"
		elif value == "下":
			direction = "down"
		else:
			direction = "center"
		self.image_generator.set_parts_data( dict_key, direction, 1 )

		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#ソフトウェア設定の辞書[self.software_setting]を変更する
	#（wx.SpinCtrlの整数値をリストに入れる）
	def change_setting_dict( self, dict_key, value, list_position = None ):
		target_value = self.software_setting[ dict_key ]

		if isinstance( target_value, list ) and isinstance( list_position, int ) and list_position >= 0 and list_position < len( target_value ):
			target_value[ list_position ] = value
			self.software_setting[ dict_key ] = target_value

		elif list_position is None:
			self.software_setting[ dict_key ] = value


	#敬称を中央に1つか、二列それぞれに付けるか切り替える操作
	def change_destination_honorific_mode( self, event ):
		if self.destination_honorific_mode.GetSelection() == 0:
			self.image_generator.set_parts_data( "twoname-honorific-mode", 1 )
		elif self.destination_honorific_mode.GetSelection() == 1:
			self.image_generator.set_parts_data( "twoname-honorific-mode", 2 )
		else:
			self.image_generator.set_parts_data( "twoname-honorific-mode", 3 )

		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#宛名2人目の配置を上寄せか下寄せか切り替える操作
	def change_destination_twoname_alignment_mode( self, event ):
		if self.destination_twoname_alignment_mode.GetSelection() == 1:
			self.image_generator.set_parts_data( "twoname-alignment-mode", "top" )
		else:
			self.image_generator.set_parts_data( "twoname-alignment-mode", "bottom" )

		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#デフォルトの敬称を使うかどうかのチェックの処理
	def change_enable_default_honorific( self, event ):
		self.column_etc_dictionary[ "enable-default-honorific" ] =  self.checkbox_enable_default_honorific.GetValue() #設置値辞書に書き込む

		#checkboxのオンオフによって、入力欄の有効無効を切り替える
		if self.checkbox_enable_default_honorific.GetValue() is True:
			self.default_honorific.Enable()
		else:
			self.default_honorific.Disable()

		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#指定がない場合のデフォルトの敬称を設置値辞書に書き込む
	def input_default_honorific( self, event ):
		default_honorific = self.default_honorific.GetValue()
		self.column_etc_dictionary[ "default-honorific" ] = self.default_honorific.GetValue()

		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#敬称に列を割り当てるか無効にするか
	def change_enable_table_honorific( self, event ):
		if self.checkbox_enable_table_honorific.GetValue() is True:
			self.column_etc_dictionary[ "enable-honorific-in-table" ] = True
			self.honorific_column.Enable()
		else:
			self.column_etc_dictionary[ "enable-honorific-in-table" ] = False
			self.honorific_column.Disable()

		self.set_grid_labels()

	#タイトルバーにファイル情報を表示するかの設定変更
	def change_titlebar_mode( self, event ):
		titlebar_mode = self.combobox_titlebar_mode.GetValue()

		if "ファイル名" in titlebar_mode:
			self.software_setting[ "write-fileinfo-on-titlebar" ] = "filename"
		elif "パス" in titlebar_mode:
			self.software_setting[ "write-fileinfo-on-titlebar" ] = "filepath"
		else:
			self.software_setting[ "write-fileinfo-on-titlebar" ] = "no-fileinfo"

		#現在のタイトルバーを現在の設定で表示し直す
		self.write_titlebar( self.software_setting[ "write-fileinfo-on-titlebar" ] )

	#印刷を上下180°回転して印刷するかの設定変更
	def change_upsidedown_print( self, event ):
		self.column_etc_dictionary[ "upside-down-print" ] = self.upsidedown_print_checkbox.GetValue()

		#レイアウト参考画像部分の注意書きを更新
		if self.column_etc_dictionary[ "upside-down-print" ] is True:
			upside_down_message = "、余白の切り抜きが180°反転中"
		else:
			upside_down_message = ""
		self.redframe_notes.SetLabel( "※赤枠の内部だけが印刷されます" + upside_down_message )

		#サンプル画像の赤枠を更新する
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	#現在の郵便番号や住所氏名の各パーツの配置と用紙種別＆サイズをレイアウトファイルとして保存する
	def layoutfile_save( self, event ):
		#現在時刻の取得とデフォルトファイル名の作成
		now_time = datetime.datetime.now()
		nowtime_text = now_time.strftime( "%Y年%m月%d日%H時%M分" )
		default_layoutfile_name = "Riosanatea用はがき封筒レイアウト設定（用紙：" + self.paper_size_data[ "category" ] + "）-" + nowtime_text + ".hhl"

		fdialog = wx.FileDialog( self, "保存するレイアウトファイル名を入力してください", defaultDir = os.path.split( sys.argv[0] )[0], defaultFile = default_layoutfile_name, style = wx.FD_SAVE, wildcard = "Layout files (*.hhl)|*.hhl" )

		if fdialog.ShowModal() == wx.ID_OK:
			filename = fdialog.GetFilename()
			dirpath = fdialog.GetDirectory()

			#保存するデータをまとめる
			paper_and_layout_data = "paper-size = " + json.dumps( self.paper_size_data )
			paper_and_layout_data += "\n\n"
			paper_and_layout_data += "atena-layout = " + json.dumps( self.image_generator.get_parts_dictionary() )

			with open( os.path.join( dirpath, filename ), "w" ) as f:
				f.write( paper_and_layout_data )

		fdialog.Destroy()

	#レイアウトファイルを読み込んで現在の設定に反映させる
	def layoutfile_open( self, event ):
		paper_size_savedata_json = ""
		atena_layout_savedata_json = ""
		saved_paper_size_dict = {}
		saved_atena_layout_dict = {}

		fdialog = wx.FileDialog( self, "ファイルを選択してください", defaultDir = os.path.split( sys.argv[0] )[0], wildcard = "Layout files (*.hhl)|*.hhl" )

		if fdialog.ShowModal() == wx.ID_OK:
			filename = fdialog.GetFilename()
			dirpath = fdialog.GetDirectory()
			filepath = os.path.join ( dirpath, filename )

			with open( filepath ) as f:

				for line in f:
					line_text = line.rstrip('\r').rstrip('\n')

					if re.match( "paper-size = ", line_text ):
						paper_size_savedata_json = re.sub( "^paper-size = ", "", line_text )

						#JSON化していた辞書を復元して元の辞書を更新する
						restored_dict = self.restore_dictionary_from_json( paper_size_savedata_json )

						if restored_dict is False:
							error_dialog = wx.MessageDialog( self,  message = "レイアウトファイルから項目「paper-size」の設定を復元できませんでした。\nレイアウトファイルが壊れている可能性があります。\nこのレイアウトファイルの内容は適用しません。", caption = "レイアウトファイル読み込みエラー", style = wx.OK | wx.ICON_ERROR )
							error_dialog.ShowModal()
							error_dialog.Destroy()
							return "レイアウトファイルの「paper-size」設定を復元できませんでした"
						else:
							saved_paper_size_dict = restored_dict

						self.paper_size_data.update( saved_paper_size_dict )

					elif re.match( "atena-layout = ", line_text ):
						atena_layout_savedata_json = re.sub( "^atena-layout = ", "", line_text )

						#JSON化していた辞書を復元して、はがきイメージ作成インスタンスの設定値上書き用辞書を更新する
						restored_dict = self.restore_dictionary_from_json( atena_layout_savedata_json )

						if restored_dict is False:
							error_dialog =  wx.MessageDialog( self,  message = "レイアウトファイルから項目「atena-layout」の設定を復元できませんでした。\nレイアウトファイルが壊れている可能性があります。\nこのレイアウトファイルの内容は適用しません。", caption = "レイアウトファイル読み込みエラー", style = wx.OK | wx.ICON_ERROR )
							error_dialog.ShowModal()
							error_dialog.Destroy()
							return "レイアウトファイルの「atena-layout」設定を復元できませんでした"
						else:
							#違う環境のレイアウトファイルを読んだらフォントファイルがなくて異常終了することがあった。
							#その対策として、それまでのレイアウト辞書をコピーしておいて
							#読み込んだ辞書のフォントファイルが存在しないなら以前のフォントに上書きするようにした。
							parts_dict_copy = self.image_generator.get_parts_dictionary()

							if not "fontfile" in restored_dict:
								error_dialog = wx.MessageDialog( self,  message = "レイアウトファイル中にフォントファイルの情報が存在しないので、これまでのフォント「" + parts_dict_copy[ "fontfile" ] + "」を使用します。", caption = "Error", style = wx.OK | wx.ICON_ERROR )
								error_dialog.ShowModal()
								error_dialog.Destroy()

								restored_dict[ "fontfile" ] = parts_dict_copy[ "fontfile" ]

							elif not os.path.isfile( restored_dict[ "fontfile" ] ):
								error_dialog = wx.MessageDialog( self,  message = "レイアウトファイルに記載されているフォントファイル「" + restored_dict[ "fontfile" ] + "」が存在しないので、これまでのフォント「" + parts_dict_copy[ "fontfile" ] + "」を使用します。", caption = "Error", style = wx.OK | wx.ICON_ERROR )
								error_dialog.ShowModal()
								error_dialog.Destroy()

								restored_dict[ "fontfile" ] = parts_dict_copy[ "fontfile" ]

							saved_atena_layout_dict = restored_dict

			if len( saved_paper_size_dict ) == 0:
				error_dialog = wx.MessageDialog( self,  message = "レイアウトファイルからpaper-sizeの項目が復元できていないので、このレイアウトファイルの内容は適用しません。", caption = "Error", style = wx.OK | wx.ICON_ERROR )
				error_dialog.ShowModal()
				error_dialog.Destroy()
				return "レイアウトファイルから「paper-size」設定の復元に失敗しました"

			elif len( saved_atena_layout_dict ) == 0:
				error_dialog = wx.MessageDialog( self,  message = "レイアウトファイルからatena-layoutの項目が復元できていないので、このレイアウトファイルの内容は適用しません。", caption = "Error", style = wx.OK | wx.ICON_ERROR )
				error_dialog.ShowModal()
				error_dialog.Destroy()
				return "レイアウトファイルから「atena-layout」設定の復元に失敗しました"

			self.overwrite_dict_for_image_generator = saved_atena_layout_dict

			#更新された用紙サイズと設定値の辞書をもとに宛名画像作成インスタンスを再作成する
			self.image_generator = atena_image_maker( papersize_widthheight_millimetre =  ( self.paper_size_data[ "width" ], self.paper_size_data[ "height" ] ), overwrite_settings = self.overwrite_dict_for_image_generator )

			#各入力欄などの数値や選択の表示を再設定する
			#一部、更新された宛名画像作成オブジェクトの値を使用するので、再作成の後に置かないといけない
			self.layout_widgets_initialize()

			#サンプルイメージを更新する
			self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

		fdialog.Destroy()

	#印刷用紙のサイズが変更された場合に設定辞書や表示を連動して変更させる
	def change_paper_size( self, event ):
		paper_size_cluster = self.combobox_paper_size.GetValue()
		paper_size_category = paper_size_cluster.split( "," )[0]
		paper_size_width = int( paper_size_cluster.split( "," )[1].split( "x" )[0] )
		paper_size_height = int( paper_size_cluster.split( "," )[1].split( "x" )[1] )
		glue_area_height = 0

		#大本の用紙サイズの設定を更新する
		self.paper_size_data.update( { "category" : paper_size_category, "width" : paper_size_width, "height" : paper_size_height } )

		#用紙変更前のパーツ配置をバックアップする
		old_parts_dict = self.image_generator.get_parts_dictionary()
		new_parts_dict = copy.deepcopy( old_parts_dict )

		#現在の配置状態のバックアップから、引き継がれては困る要素（フォントサイズ）を削除
		for i in [ font_key for font_key in new_parts_dict.keys() if "fontsize" in font_key ]:
			new_parts_dict.pop(i)

		standard_parts_dict = self.image_generator.get_standard_parts_dictionary()

		#はがき用のデフォルト配置のコピーから、引き継がれては困る要素（フォントのパスとフォントサイズ）を削除
		for i in [ font_key for font_key in standard_parts_dict.keys() if "font" in font_key ]:
			standard_parts_dict.pop(i)

		#パーツ配置の一部をある程度自動調整するか、ダイアログで尋ねる
		if "はがき" in paper_size_category:
			dialog_additional_message = ""
		else:
			dialog_additional_message = "\n\n※ 郵便番号の枠が印刷されている封筒の場合、自動調整しても郵便番号は枠へ正確には収まらないと思います。お手数ですがテスト印刷と手動調整をしてください。"

		auto_relocation_question_dialog =  wx.MessageDialog( None,  message = "用紙サイズを変更しますが、郵便番号や住所氏名の配置と収納範囲をおおまかに自動調整しておきますか？\n\nNoを選ぶと前の用紙の配置のまま用紙の大きさだけが変わります。\nそうすると、宛先氏名が中央にないなど手動調整の手間が増えると予想されます。" + dialog_additional_message, caption = "自動調整の可否", style = wx.YES_NO | wx.ICON_QUESTION )

		if auto_relocation_question_dialog.ShowModal() == wx.ID_YES:

			if paper_size_category == "はがき" or paper_size_category == "往復はがき":
				new_parts_dict.update( standard_parts_dict )

			else:
				list_composite = lambda hagaki_standard_size, paper_size_rate, control_rate: [ int( hagaki_standard_size[0] + hagaki_standard_size[0] * ( paper_size_rate[0] - 1 ) * control_rate ), int( hagaki_standard_size[1] + hagaki_standard_size[1] * ( paper_size_rate[1] - 1 ) * control_rate ) ]
				parts_relocate_rate = ( paper_size_width / 100, paper_size_height / 148 )

				#はがきを大幅に超えるような大きさの用紙で、郵便番号や宛名をそのままの比率で
				#拡大すると大きすぎるので、増減分をやや抑え気味にするための追加の比率を設定する
				parts_size_adjust_rate = 0.7

				#各パーツの位置の自動調整
				new_parts_dict[ "postalcode-position" ] = list_composite( standard_parts_dict[ "postalcode-position" ], parts_relocate_rate, 1 )
				new_parts_dict[ "name-position" ] = list_composite( standard_parts_dict[ "name-position" ], parts_relocate_rate, 1 )
				new_parts_dict[ "address-position" ] = list_composite( standard_parts_dict[ "address-position" ], parts_relocate_rate, 1 )
				new_parts_dict[ "our-postalcode-position" ] = list_composite( standard_parts_dict[ "our-postalcode-position" ], parts_relocate_rate, 1 )
				new_parts_dict[ "our-name-position" ] = list_composite( standard_parts_dict[ "our-name-position" ], parts_relocate_rate, 1 )
				new_parts_dict[ "our-address-position" ] = list_composite( standard_parts_dict[ "our-address-position" ], parts_relocate_rate, 1 )

				#各パーツを収める範囲の自動調整
				new_parts_dict[ "postalcode-letter-areasize" ] = list_composite( standard_parts_dict[ "postalcode-letter-areasize" ], parts_relocate_rate, parts_size_adjust_rate )
				new_parts_dict[ "name-areasize" ] = list_composite( standard_parts_dict[ "name-areasize" ], parts_relocate_rate, parts_size_adjust_rate )
				new_parts_dict[ "address-areasize" ] = list_composite( standard_parts_dict[ "address-areasize" ], parts_relocate_rate, parts_size_adjust_rate )
				new_parts_dict[ "our-postalcode-letter-areasize" ] = list_composite( standard_parts_dict[ "our-postalcode-letter-areasize" ], parts_relocate_rate, parts_size_adjust_rate )
				new_parts_dict[ "our-name-areasize" ] = list_composite( standard_parts_dict[ "our-name-areasize" ], parts_relocate_rate, parts_size_adjust_rate )
				new_parts_dict[ "our-address-areasize" ] = list_composite( standard_parts_dict[ "our-address-areasize" ], parts_relocate_rate, parts_size_adjust_rate )

				#郵便番号の間隔の自動調整
				new_parts_dict[ "postalcode-placement" ] = [ int( i * parts_relocate_rate[0] * parts_size_adjust_rate ) for i in standard_parts_dict[ "postalcode-placement" ] ]
				new_parts_dict[ "our-postalcode-placement" ] = [ int( i * parts_relocate_rate[0] * parts_size_adjust_rate ) for i in standard_parts_dict[ "our-postalcode-placement" ] ]

				#住所氏名における二列の間隔の自動調整
				new_parts_dict[ "name-bind-space" ] = int( standard_parts_dict[ "name-bind-space" ] * parts_relocate_rate[0] )
				new_parts_dict[ "address-bind-space" ] = int( standard_parts_dict[ "address-bind-space" ] * parts_relocate_rate[0] )
				new_parts_dict[ "our-name-bind-space" ] = int( standard_parts_dict[ "our-name-bind-space" ] * parts_relocate_rate[0] )
				new_parts_dict[ "our-address-bind-space" ] = int( standard_parts_dict[ "our-address-bind-space" ] * parts_relocate_rate[0] )

			auto_relocation_question_dialog.Destroy()

		#宛名画像生成インスタンスの作り直し
		self.image_generator = atena_image_maker( papersize_widthheight_millimetre =  ( self.paper_size_data[ "width" ], self.paper_size_data[ "height" ] ), overwrite_settings = new_parts_dict )

		#レイアウト参考画像部分の用紙サイズ表示を更新
		self.sampleimage_sbox.SetLabel( "印刷の参考イメージ ( " + self.paper_size_data[ "category" ] + "、" + str( self.paper_size_data[ "width" ] ) + "mm x " + str( self.paper_size_data[ "height" ] ) + "mm )" )

		#レイアウト参考画像を更新
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

		#各入力欄の値を再設定する
		self.layout_widgets_initialize()

	#タイトルバーの表示を更新する
	def write_titlebar( self, titlebar_mode ):

		if titlebar_mode == "filename":
			self.SetTitle( self.base_window_title + " （" + os.path.basename( self.opened_file_path ) + "）" )
		elif titlebar_mode == "filepath":
			self.SetTitle( self.base_window_title + " （" + self.opened_file_path + "）" )
		else:
			self.SetTitle( self.base_window_title )


	#差出人郵便番号が半角数字7個（または空白）かどうか判定する
	def check_our_postalcode( self, event ):
		input_value = self.our_postalcode_inputbox.GetValue()
		regular_expression_check = False

		#Python2以下では、全角文字1つが一文字とみなされないのか
		#[0-9０１２３４５６７８９]{3}としても正常に判定できない
		if re.match( "^([0-9]|０|１|２|３|４|５|６|７|８|９){3}( |　|-|−|ー|―|‐)?([0-9]|０|１|２|３|４|５|６|７|８|９){4}$", input_value ):
			self.our_data[ "our-postalcode-data" ] = str( input_value )
			self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )
			regular_expression_check = True

		else:
			if input_value == "":
				self.our_data[ "our-postalcode-data" ] = ""
				self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )
				self.our_postalcode_error_text.SetLabel( "" )
				self.our_postalcode_inputbox.set_check( False )
			else:
				self.our_postalcode_inputbox.SetValue( str( self.our_data.get( "our-postalcode-data", "" ) ) )
				self.our_postalcode_inputbox.set_check( True )
				self.our_postalcode_error_text.SetLabel( "　入力値が数字7個ではなかったので、以前の値に戻しました（場合によっては空白化）" )

		#ここらへんの処理は、自分でもなにがなんだかわからない。
		#非整数値の手動入力や貼り付けで不可解な流れになるため、場当たり的に対応した結果。
		if regular_expression_check is True:
			if self.our_postalcode_inputbox.get_check() is False:
				self.our_postalcode_error_text.SetLabel( "" )
			self.our_postalcode_inputbox.set_check( False )

	def change_our_address1( self, event ):
		address1_value = self.our_address1.GetValue()

		self.our_data[ "our-address1-data" ] = address1_value
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	def change_our_address2( self, event ):
		address2_value = self.our_address2.GetValue()

		self.our_data[ "our-address2-data" ] = address2_value
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	def change_our_name1( self, event ):
		name1_value = self.our_name1.GetValue()

		self.our_data[ "our-name1-data" ] = name1_value
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

	def change_our_name2( self, event ):
		name2_value = self.our_name2.GetValue()

		self.our_data[ "our-name2-data" ] = name2_value
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )


	#宛名のサンプルイメージを取得して（赤枠をつけてリサイズしてから）パネルに表示する
	def show_sample_image( self, make_atena_image = True, cutted_atena_image_upside_down = False ):

		#余白幅を変更する場合（make_atena_image is False）は、宛名画像の構成を
		#変えるわけではないので画像の再取得は飛ばす
		if make_atena_image is True:
			if self.checkbox_enable_default_honorific.GetValue() is True and self.default_honorific.GetValue() != "":
				sample_honorific = self.default_honorific.GetValue()
			else:
				sample_honorific = "様"

			data_example = { "postal-code" : self.destination_postalcode_example, "name1" : self.destination_name1_example, "name2" : self.destination_name2_example, "address1" : self.destination_address1_example, "address2" : self.destination_address2_example, "company" : "株式会社サンプル", "department" : "営業部", "honorific" : sample_honorific, "our-postal-code" : self.our_data[ "our-postalcode-data" ], "our-name1" : self.our_data[ "our-name1-data" ], "our-name2" : self.our_data[ "our-name2-data" ], "our-address1" : self.our_data[ "our-address1-data" ], "our-address2" : self.our_data[ "our-address2-data" ] }

			#宛名画像の取得
			self.sample_grayscale_image = self.image_generator.get_atena_image( data_example, area_frame = self.column_etc_dictionary[ "sampleimage-areaframe" ] )

		self.sample_color_image = ImageOps.colorize( self.sample_grayscale_image, ( 0, 0, 0 ), ( 255, 255, 255 ) )

		#赤枠をつける処理
		redline_width = self.image_generator.get_parts_data( "redline-width" )
		space_tblr_list = self.image_generator.convert_mm_to_pixel( self.column_etc_dictionary[ "printer-space-top,bottom,left,right" ] )


		if cutted_atena_image_upside_down is False:
			linepoint_lu = [ space_tblr_list[2], space_tblr_list[0] ] #赤線をひく左上の座標（実際には線幅分の補正も入る場合あり）
			linepoint_ru = [ self.sample_color_image.size[0] - space_tblr_list[3], space_tblr_list[0] ] #右上の座標
			linepoint_rb = [ self.sample_color_image.size[0] - space_tblr_list[3], self.sample_color_image.size[1] - space_tblr_list[1] ] #右下の座標
			linepoint_lb = [ space_tblr_list[2], self.sample_color_image.size[1] - space_tblr_list[1] ] #左下の座標

		else:
			linepoint_lu = [ space_tblr_list[3], space_tblr_list[1] ]
			linepoint_ru = [ self.sample_color_image.size[0] - space_tblr_list[2], space_tblr_list[1] ]
			linepoint_rb = [ self.sample_color_image.size[0] - space_tblr_list[2], self.sample_color_image.size[1] - space_tblr_list[0] ]
			linepoint_lb = [ space_tblr_list[3], self.sample_color_image.size[1] - space_tblr_list[0] ]

		#左上から右上に赤線をひく
		draw = ImageDraw.Draw( self.sample_color_image )
		draw.rectangle( ( ( linepoint_lu[0], linepoint_lu[1] ), ( linepoint_ru[0], linepoint_ru[1] + redline_width ) ), fill = "red" )

		#右上から右下に赤線をひく
		draw.rectangle( ( ( linepoint_ru[0] - redline_width, linepoint_ru[1] ), ( linepoint_rb[0], linepoint_rb[1] ) ), fill = "red" )

		#左下から右下に赤線をひく
		draw.rectangle( ( ( linepoint_lb[0], linepoint_lb[1] - redline_width ), ( linepoint_rb[0], linepoint_rb[1] ) ), fill = "red" )

		#左上から左下に赤線をひく
		draw.rectangle( ( ( linepoint_lu[0], linepoint_lu[1] ), ( linepoint_lb[0] + redline_width, linepoint_lb[1] ) ), fill = "red" )

		#PILイメージをwx.Imageに変換する
		self.wx_sample_image = wx.Image( self.sample_color_image.size[0], self.sample_color_image.size[1] )
		self.wx_sample_image.SetData( self.sample_color_image.convert('RGB').tobytes() )

		#Rescale関数は元の画像を上書きするようなので、拡大縮小を繰り返すうちに粗くなっていく
		#それを防ぐためにself.wx_sample_imageはそのままとっておいて
		#self.wx_resized_sampleにコピーしてからリサイズする（copy.deepcopyは使えない）
		self.wx_resized_sample = self.wx_sample_image.Copy()

		#パネルよりも横長なら、パネルの幅に合わせてリサイズする
		if float( self.wx_resized_sample.GetSize()[0] ) / float( self.wx_resized_sample.GetSize()[1] ) > float( self.sample_image_panel.GetSize()[0] ) / float( self.sample_image_panel.GetSize()[1] ):

			self.wx_resized_sample.Rescale( self.sample_image_panel.GetSize()[0], int( self.wx_resized_sample.GetSize()[1] * self.sample_image_panel.GetSize()[0] / self.wx_resized_sample.GetSize()[0] ), quality=wx.IMAGE_QUALITY_HIGH )
			self.display_position[0] = 0
			self.display_position[1] = int( ( self.sample_image_panel.GetSize()[1] - self.wx_resized_sample.GetSize()[1] ) / 2 )

		else:
			#パネルよりも縦長なら、パネルの高さに合わせてリサイズする
			self.wx_resized_sample.Rescale( int( self.wx_resized_sample.GetSize()[0] * self.sample_image_panel.GetSize()[1] / self.wx_resized_sample.GetSize()[1] ), self.sample_image_panel.GetSize()[1], quality=wx.IMAGE_QUALITY_HIGH )
			self.display_position[0] = int( ( self.sample_image_panel.GetSize()[0] - self.wx_resized_sample.GetSize()[0] ) / 2 )
			self.display_position[1] = 0

		self.wx_bitmap_image =  self.wx_resized_sample.ConvertToBitmap()
		self.Refresh()


	#パネルの大きさが変わったら、パネルに合わせてサンプル画像をリサイズする
	def adjust_sample_image_with_panel( self, event ):
		#パネルサイズの変更に合わせてリサイズするたびに、元のself.wx_sample_imageから再度コピーする（copy.deepcopyは使えない）
		self.wx_resized_sample = self.wx_sample_image.Copy()

		#パネルよりも横長なら、パネルの幅に合わせてリサイズする
		if float( self.wx_resized_sample.GetSize()[0] ) / float( self.wx_resized_sample.GetSize()[1] ) > float( self.sample_image_panel.GetSize()[0] ) / float( self.sample_image_panel.GetSize()[1] ):

			self.wx_resized_sample.Rescale( self.sample_image_panel.GetSize()[0], int( self.wx_resized_sample.GetSize()[1] * self.sample_image_panel.GetSize()[0] / self.wx_resized_sample.GetSize()[0] ), quality=wx.IMAGE_QUALITY_HIGH )
			self.display_position[0] = 0
			self.display_position[1] = int( ( self.sample_image_panel.GetSize()[1] - self.wx_resized_sample.GetSize()[1] ) / 2 )

		else:
			#パネルよりも縦長なら、パネルの高さに合わせてリサイズする
			self.wx_resized_sample.Rescale( int( self.wx_resized_sample.GetSize()[0] * self.sample_image_panel.GetSize()[1] / self.wx_resized_sample.GetSize()[1] ), self.sample_image_panel.GetSize()[1], quality=wx.IMAGE_QUALITY_HIGH )
			self.display_position[0] = int( ( self.sample_image_panel.GetSize()[0] - self.wx_resized_sample.GetSize()[0] ) / 2 )
			self.display_position[1] = 0

		self.wx_bitmap_image =  self.wx_resized_sample.ConvertToBitmap()
		self.Refresh()


	def get_fontlist( self ):
		font_rawdata = []
		font_list = []
		fc_command_existence = True

		#シェルコマンドの「fc-list」でフォントの一覧が取得できるそうなので
		#それをsubprocessで動かして返される値を取得する
		try:
			p = subprocess.Popen( [ "fc-list" ], stdout=subprocess.PIPE )
		except:
			fc_command_existence = False

		#「fc-list」のシェルコマンドがずっと使える保証もないと思うので
		#subprocess.Popenが順調に進んだ場合のみ一覧の取得処理を続行する
		if fc_command_existence == True:
			for line in p.stdout.readlines():
				font_rawdata.append( line.decode( "utf-8" ).rstrip( '\n' ) )

			for j in font_rawdata:
				fontpath = j.split( ":" )[0]

				if fontpath != "":
					fontname = os.path.splitext( os.path.basename( fontpath ) )[0]

					#フォントによっては画像化のときにエラーになるので、問題があるフォントはスルーする
					#具体的には「NotoColorEmoji.ttf」でエラーになる
					#このテェックのせいで、スクリプトの起動が少しもたつくようになった
					try:
						fnt = ImageFont.truetype( font = fontpath, size = 20, encoding = "unic" )
					except OSError:
						continue

					#重複が多いので、すでにリスト中にある場合は入れないようにして重複を解消する
					if not [ fontname, fontpath ] in font_list:
						font_list.append( [ fontname, fontpath ] )

		font_list.sort()
		return font_list


	#コンボボックスからフォントの変更があったら、宛名オブジェクト内の
	#フォント設定を変更して、台紙画像のサイズも再決定して、サンプル画像を再取得・再表示する
	def change_font( self, event ):
		obj = event.GetEventObject()
		choiced_fontpath = obj.GetClientData(obj.GetSelection())
		self.image_generator.set_parts_data( "fontfile", choiced_fontpath )
		self.image_generator.determine_fontmat_size( choiced_fontpath )
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )


	#参考画像のパーツ範囲枠の表示切り替え
	def change_sample_image_areaframe( self, event ):
		self.column_etc_dictionary[ "sampleimage-areaframe" ] = self.checkbox_sample_image_areaframe.GetValue()
		self.show_sample_image( cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )


	#余白計測用の枠画像を印刷
	def print_frame_image( self, event ):
		red_frame_image = self.image_generator.get_red_frame_image()

		pil_printing( red_frame_image, paper_size = "Custom." + str( self.paper_size_data[ "width" ] ) + "x" + str( self.paper_size_data[ "height" ] ) + "mm" )

	#印刷の可否判別を有効にするか無効にするか
	def change_print_control_on_off( self, event ):

		if self.checkbox_printctrl.GetValue() is True:
			self.column_etc_dictionary[ "print-control" ] = True
			self.printctrl_column.Enable()
			self.print_sign.Enable()
			self.combobox_print_do_or_ignore.Enable()

		else:
			self.column_etc_dictionary[ "print-control" ] = False
			self.printctrl_column.Disable()
			self.print_sign.Disable()
			self.combobox_print_do_or_ignore.Disable()

		self.set_grid_labels()

	#印刷の可否を判別する記述の列を指定する
	def chage_printctrl_column( self, event ):
		self.column_etc_dictionary[ "print-control-column" ] = self.printctrl_column.GetValue() - 1
		self.set_grid_labels()

	#印刷の可否を判別する基準の文字列を指定する
	def input_print_sign( self, event ):
		self.column_etc_dictionary[ "print-sign" ] = self.print_sign.GetValue()

	#該当した場合に印刷するのか、該当したら無視するのかの方向性
	def change_print_execute( self, event ):
		if self.combobox_print_do_or_ignore.GetSelection() == 0:
			self.column_etc_dictionary[ "print-or-ignore" ] = "print"
		else:
			self.column_etc_dictionary[ "print-or-ignore" ] = "ignore"


	#設定保存ボタンから、宛名インスタンスやGUIの設定値辞書をINIファイルに保存する
	def save_settings( self, event ):
		inifile_path = os.path.splitext( sys.argv[0] )[0] + ".ini"
		ini_data = ""

		ini_data += "[settings]\n\n"
		#↓以前は辞書項目中の「%」がINIの読み込みを妨害していたのでBase64化して保存していたが
		#半角の%が原因だとわかったので全角の％として解決し、json化したそのままの保存に変更した
		ini_data += "paper-size = " + json.dumps( self.paper_size_data ) + "\n\n"
		ini_data += "atena-layout = " + json.dumps( self.image_generator.get_parts_dictionary() ) + "\n\n"
		ini_data += "column-etc = " + json.dumps( self.column_etc_dictionary ) + "\n\n"
		ini_data += "our-data = " + json.dumps( self.our_data ) + "\n\n"
		ini_data += "software-setting = " + json.dumps( self.software_setting ) + "\n\n"

		if self.outer_font != "":
			ini_data += "outer-font = " + self.outer_font + "\n\n"

		with open( inifile_path, "w" ) as f:
			f.write( ini_data )

		#保存した後で、各設定値辞書の設定保存した時点での内容をコピーして記録しておく
		#ソフトの終了時に、その時点での設定と比較して設定変更があったかチェックするためのもの
		self.savepoint_paper_size_data = copy.deepcopy( self.paper_size_data )
		self.savepoint_parts_dictionary = copy.deepcopy( self.image_generator.get_parts_dictionary() )
		self.savepoint_software_setting = copy.deepcopy( self.software_setting )
		self.savepoint_column_etc_dictionary = copy.deepcopy( self.column_etc_dictionary )
		self.savepoint_our_data = copy.deepcopy( self.our_data )


	#INIファイルを読み込んでJSON化していた辞書を復元する
	def load_settings( self ):
		inifile_path = os.path.splitext( sys.argv[0] )[0] + ".ini"

		if os.path.isfile( inifile_path ):
			inisetting = configparser.ConfigParser()
			inisetting.read( inifile_path )

			if inisetting.has_option( "settings", "paper-size" ):
				saved_paper_size_json = inisetting.get( "settings", "paper-size" )
				restored_dict = self.restore_dictionary_from_json( saved_paper_size_json )

				if restored_dict is False:
					self.setting_json_error_dialog( "column-etc" )
					saved_paper_size_dict = {}
				else:
					saved_paper_size_dict = restored_dict

				for paper_size_value in saved_paper_size_dict.items():
					self.paper_size_data[ paper_size_value[0] ] = paper_size_value[1]

			if inisetting.has_option( "settings", "atena-layout" ):
				saved_atena_layout_json = inisetting.get( "settings", "atena-layout" )
				restored_dict = self.restore_dictionary_from_json( saved_atena_layout_json )

				if restored_dict is False:
					self.setting_json_error_dialog( "atena-layout" )
				else:
					self.overwrite_dict_for_image_generator = restored_dict

			if inisetting.has_option( "settings", "hagaki-layout" ):
				settingfile_error_dialog = wx.MessageDialog( None,  message = "設定ファイルが旧式のままのようなので、今回は本体ウィンドウを起動せずに終了します。\n\n設定ファイルの仕様が変更されています。\nお手数ですが設定ファイルをエディタで開いて「hagaki-layout = 」を「atena-layout = 」に書き換えてから起動しなおしてください。", caption = "Error", style = wx.OK | wx.ICON_EXCLAMATION )
				settingfile_error_dialog.ShowModal()
				settingfile_error_dialog.Destroy()
				self.Close()

			if inisetting.has_option( "settings", "column-etc" ):
				saved_column_etc_json = inisetting.get( "settings", "column-etc" )
				restored_dict = self.restore_dictionary_from_json( saved_column_etc_json )

				if restored_dict is False:
					self.setting_json_error_dialog( "column-etc" )
					saved_column_etc_dict = {}
				else:
					saved_column_etc_dict = restored_dict

				for column_key_value in saved_column_etc_dict.items():
					self.column_etc_dictionary[ column_key_value[0] ] = column_key_value[1]

			if inisetting.has_option( "settings", "our-data" ):
				saved_our_data_json = inisetting.get( "settings", "our-data" )
				restored_dict = self.restore_dictionary_from_json( saved_our_data_json )

				if restored_dict is False:
					self.setting_json_error_dialog( "our-data" )
					saved_our_data_dict = {}
				else:
					saved_our_data_dict = restored_dict

				for our_key_value in saved_our_data_dict.items():
					self.our_data[ our_key_value[0] ] = our_key_value[1]

			if inisetting.has_option( "settings", "software-setting" ):
				saved_software_setting_json = inisetting.get( "settings", "software-setting" )
				restored_dict = self.restore_dictionary_from_json( saved_software_setting_json )

				if restored_dict is False:
					self.setting_json_error_dialog( "software-setting" )
					saved_software_setting_dict = {}
				else:
					saved_software_setting_dict = restored_dict

				for setting_key_value in saved_software_setting_dict.items():
					self.software_setting[ setting_key_value[0] ] = setting_key_value[1]

			if inisetting.has_option( "settings", "outer-font" ):
				self.outer_font = inisetting.get( "settings", "outer-font" )

				#外部フォントが指定されていても、実際に使うフォントに入れ替えるのは
				#従来のフォントが存在しない場合のみ
				if not os.path.isfile( self.image_generator.get_parts_data( "fontfile" ) ):
					self.image_generator.set_parts_data( "fontfile", self.outer_font )


	#JSON形式に変換されていた辞書を復元する関数
	def restore_dictionary_from_json( self, json_dictionary ):

		try:
			new_dictionary = json.loads( json_dictionary )
		except:
			return False

		return new_dictionary


	#INIファイルの設定値をjsonで復元する過程でエラーが生じたときに、エラーダイアログを表示する
	def setting_json_error_dialog( self, setting_index ):
		error_dialog =  wx.MessageDialog( self,  message = "INIファイルから項目「" + setting_index + "」の設定を復元できませんでした。\nINIファイルが壊れているか、仕様が変更される以前のファイルだった可能性があります。お手数ですが、設定をやり直して上書き保存してください。\n\n設定保存は「ソフトウェア設定」タブの「レイアウト、その他の設定を設定ファイル(iniファイル)に保存する」ボタンを押せばできます。", caption = "設定読み込みエラー", style = wx.OK | wx.ICON_ERROR )
		error_dialog.ShowModal()
		error_dialog.Destroy()

	#-----以下、表に関する関数-----

	#読み込んだCSVファイルの内容、または履歴中の一つを表に反映させる
	def set_table( self, string_2dimension_list ):
		#Rows(縦の行数)をlistの行数に合わせる
		current_rows = len( string_2dimension_list )

		if current_rows - self.grid.GetNumberRows() > 0:
			self.grid.AppendRows( current_rows - self.grid.GetNumberRows() )
		elif current_rows - self.grid.GetNumberRows() < 0:
			self.grid.DeleteRows( current_rows, self.grid.GetNumberRows() - current_rows )

		#Cols（横の列数）をlistの最大列数に合わせる
		max_len = max( [ len( x ) for x in string_2dimension_list ] ) #CSVがすべての行で同じ列数とは限らないので、最大の長さを求める

		if max_len - self.grid.GetNumberCols() > 0:
			self.grid.AppendCols( max_len - self.grid.GetNumberCols() )
		elif max_len - self.grid.GetNumberCols() < 0:
			self.grid.DeleteCols( max_len, self.grid.GetNumberCols() - max_len )

		#それぞれのセルにCSVの内容を入れていく
		for j in range( self.grid.GetNumberRows() ):
			for i in range( self.grid.GetNumberCols() ):
				try:
					self.grid.SetCellValue( j, i, string_2dimension_list[j][i] ) #セルに値を設定する（値は文字列のみ、( 縦の位置, 横の位置, 値 )になる）
				except:
					#CSVが行によって項目数が不揃いだったりすると、項目が少ない行では
					#欠落が生じるので、そこには空白を充てる
					self.grid.SetCellValue( j, i, "" )

		#行・列のサイズを自動調整
		self.grid.AutoSizeColumns()
		self.grid.AutoSizeRows()

	#表に変更があった場合に、表の内容をリスト化して履歴に加える
	def add_table_to_history( self, event ):

		#リドゥ中で、履歴中の現在位置が最初でなかった場合は、最新〜現在までの履歴を削除する
		if self.current_history_position != 0:
			for delete in range( self.current_history_position ):
				self.table_history.pop( 0 )
			self.current_history_position = 0

		#履歴の数が一定数に達していれば、それ以上の履歴を削除する
		if len( self.table_history ) > self.history_stock_max:
			for delete in range( len( self.table_history ) - self.history_stock_max ):
				self.table_history.pop()

		#本番の作業（表の内容をリスト化して履歴に加える）
		current_table = []
		temporary_list = []

		for j in range( self.grid.GetNumberRows() ):
			temporary_list = []

			for i in range( self.grid.GetNumberCols() ):
				temporary_list.append( self.grid.GetCellValue( j, i ) )
			current_table.append( temporary_list )

		self.table_history.insert( 0, current_table )

		#（表を操作した）イベントで使用されたなら、ステータスバーをクリア
		if event is not None:
			self.statusbar.SetStatusText( "" )


	#列のラベルを設定する
	def set_grid_labels( self ):
		#役割のある列（選択によらない必須のもの）のkeyと貼り付けたいラベルの対応
		special_labels = [ ["column-postalcode", "郵便番号" ], [ "column-address1", "宛先住所" ], [ "column-address2", "宛先住所（二列目）" ], [ "column-name1", "宛名" ], [ "column-name2", "宛名（二人目）" ], [ "column-company", "会社名" ], [ "column-department", "部署" ] ]

		#役割をもたせた列の最大値を求める
		special_column = [ self.column_etc_dictionary[ x[0] ] for x in special_labels ]
		if self.column_etc_dictionary[ "print-control" ] is True:
			special_column.append( self.column_etc_dictionary[ "print-control-column" ] )
		if self.column_etc_dictionary[ "enable-honorific-in-table" ] is True:
			special_column.append( self.column_etc_dictionary[ "column-honorific" ] )
		max_special_column = max( special_column )

		#役割のある最大列より表の列数が足りなければ、表を横に広げる
		#（逆に、余った場合に列を削除する、ということはしない）
		if max_special_column >= self.grid.GetNumberCols():
			self.grid.AppendCols( max_special_column - self.grid.GetNumberCols() + 1 )

		#指定のない列でもABC...というのはわかりにくいので、番号をつける
		nomal_labels = ["列" + str( x ) for x in range( 1, self.grid.GetNumberCols() + 1 ) ]

		#用意した一般番号ラベルから、役割の指定がある特殊列を役割ラベルに入れ替える
		for atena_label in ( special_labels ):
			target_column = self.column_etc_dictionary[ atena_label[0] ]
			if target_column < self.grid.GetNumberCols():
				nomal_labels[ target_column ] = atena_label[1]

		if self.column_etc_dictionary[ "print-control" ] is True:
			target_column = self.column_etc_dictionary[ "print-control-column" ]
			if target_column < self.grid.GetNumberCols():
				nomal_labels[ target_column ] = "印刷の可否"

		if self.column_etc_dictionary[ "enable-honorific-in-table" ] is True:
			target_column = self.column_etc_dictionary[ "column-honorific" ]
			if target_column < self.grid.GetNumberCols():
				nomal_labels[ target_column ] = "敬称"

		#実際の表のラベルを入れ替えていく
		for column in range( self.grid.GetNumberCols() ):
			self.grid.SetColLabelValue( column, nomal_labels[ column ] )

		#列幅を自動調整
		self.grid.AutoSizeColumns()


	#ファイル選択ダイアログからCSVファイルを選び、CSVファイルを開く関数に渡す
	def fileselect_and_opencsv( self, event ):
		#まず、現在の表の内容が（変更がない時点の）控えと同一かチェックする
		current_table = self.get_current_table_list()

		if current_table != self.table_checkpoint:
			question_dialog = wx.MessageDialog( parent = self, message = "現在の表の内容が変更されていますが保存されていません。\nこのまま開くと現在の内容は失われます。\n\n開く前に保存しますか？", caption = "表内容の変更に関する確認", style = wx.YES_NO | wx.ICON_QUESTION )

			if question_dialog.ShowModal() == wx.ID_YES:
				self.save_csv_file( None )
			question_dialog.Destroy()

			message_dialog = wx.MessageDialog( parent = self, message = "保存プロセスが終了したので、ファイルを開くプロセスに移ります。", caption = "現在の内容を保存しました", style = wx.OK | wx.ICON_INFORMATION )
			message_dialog.ShowModal()
			message_dialog.Destroy()

		csv_path = ""

		#ダイアログで最初に開くディレクトリを決める
		if os.path.isfile( self.opened_file_path ):
			default_dir = os.path.dirname( self.opened_file_path )
		else:
			default_dir = os.path.split( sys.argv[0] )[0]

		fdialog = wx.FileDialog( self, "CSVファイルを選択してください", defaultDir = default_dir, wildcard = "CSV files (*.csv;*.CSV)|*.csv;*.CSV" )

		#どのCSVファイルを読むかを選択
		if fdialog.ShowModal() == wx.ID_OK:
			filename = fdialog.GetFilename()
			dirpath = fdialog.GetDirectory()

			csv_path = os.path.join ( dirpath, filename )

		fdialog.Destroy()


		if csv_path != "":
			self.open_csv_file( csv_path )


	#CSVファイルを開いて履歴に登録し、表に表示する
	def open_csv_file( self, csv_path ):

		if csv_path != "" and os.path.isfile( csv_path ):

			#CSVファイルの読み込み
			#関数の中で読み込みテストをすることで、いくつかの文字コードに対応したCSV読み込みをする
			csv_data = csv_to_list( csv_path )

			#ファイル読み込み時点を履歴の起点とするので、履歴をリセットする
			self.table_history = []
			self.current_history_position = 0

			#読み込んだCSVのデータを表に書き込む
			self.set_table( csv_data )

			#印刷行範囲のデフォルト値を表の行数に合わせる
			self.print_start_line.SetValue( 1 )
			self.print_end_line.SetValue( self.grid.GetNumberRows() )

			#ファイル読み込み直後の表の内容を履歴に登録する
			#（add_table_to_history( None )を起動してもいいが、ここでは簡単にCSVのデータを代わりに履歴に追加しておく）
			self.table_history.insert( 0, csv_data )
			self.statusbar.SetStatusText( "CSVファイル「" + os.path.basename( csv_path ) + "」を開きました" )

			#表の内容の控えをリセット・更新する（変更が保存されているかのチェック用）
			self.table_checkpoint = self.get_current_table_list()

			#開いたファイルのパスを控えておく
			self.opened_file_path = csv_path

			#CSVの列数のままだとラベルを貼るために必要な列数に足りないかもしれないので
			#その調整も兼ねてラベルを貼り直す
			self.set_grid_labels()

			#タイトルバーの表示を更新する
			self.write_titlebar( self.software_setting[ "write-fileinfo-on-titlebar" ] )


	#表の内容をCSVファイルとして保存する
	def save_csv_file( self, event ):
		#表の内容をリスト化
		current_table = self.get_current_table_list()

		csv_path = ""

		#ダイアログで最初に開くディレクトリと、デフォルトの保存候補名を決める
		if os.path.isfile( self.opened_file_path ):
			save_dir = os.path.dirname( self.opened_file_path )
			default_save_name = os.path.basename( self.opened_file_path )
		else:
			save_dir = os.path.split( sys.argv[0] )[0]
			current_time = datetime.datetime.now()
			default_save_name = current_time.strftime("%Y年%m月%d日%H時%M分%S秒") + ".csv"

		fdialog = wx.FileDialog( self, "保存するファイル名を決めてください", defaultDir = save_dir, defaultFile = default_save_name, style = wx.FD_SAVE, wildcard = "CSV files (*.csv;*.CSV)|*.csv;*.CSV" )

		if fdialog.ShowModal() == wx.ID_OK:
			filename = fdialog.GetFilename()
			dirpath = fdialog.GetDirectory()

			#csvの拡張子がなければ付け足す
			if not re.search( "[.](csv|CSV)$", filename ):
				filename += ".csv"

			csv_path = os.path.join ( dirpath, filename )

			with open( csv_path, "w" ) as f:
				writer = csv.writer( f, quoting = csv.QUOTE_ALL )
				writer.writerows( current_table )

		fdialog.Destroy()

		#表の内容の控えをリセット・更新する（変更が保存されているかのチェック用）
		self.table_checkpoint = current_table


	#現在の表の内容をリスト化して取得する
	def get_current_table_list( self ):
		current_table_list = []
		temporary_list = []

		for j in range( self.grid.GetNumberRows() ):
			temporary_list = []

			for i in range( self.grid.GetNumberCols() ):
				cell_value = self.grid.GetCellValue( j, i )
				temporary_list.append( cell_value )
			current_table_list.append( temporary_list )

		return current_table_list


	#タブの切り替えに応じて、住所表のタブでのみステータスバーを出して、それ以外では隠す
	def showhide_statusbar_with_tab( self, event ):
		if self.notebook.GetSelection() == 0:
			self.statusbar.Show()
		else:
			self.statusbar.Hide()

	#印刷イメージ確認用のダイアログを表示する
	def show_preview_dialog( self, event ):
		#現在選択中の行を得る
		selected_row_position_list = self.grid.GetSelectedRows()
		selected_row_position = 0

		if len( selected_row_position_list ) > 0:
			#行が取得できていれば、それらのリストの最初の行番号だけにする
			selected_row_position = selected_row_position_list[0]

		current_table = self.get_current_table_list()

		image_preview_dialog = AtenaPreviewDialog( paper_data_dict = self.paper_size_data, destination_list = current_table, column_data = self.column_etc_dictionary, our_data = self.our_data, min_line_int = self.print_start_line.GetValue(), max_line_int = self.print_end_line.GetValue(), image_generator_instance = self.image_generator, space_tblr_mm_list = self.column_etc_dictionary[ "printer-space-top,bottom,left,right" ], cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ], current_row = selected_row_position )

		image_preview_dialog.ShowModal()
		image_preview_dialog.Destroy()


	#行か列を追加ないし削減する
	def row_col_add_del( self, event ):
		#現在選択中の行を得る
		selected_row_position_list = self.grid.GetSelectedRows()
		selected_row_position = 0

		if len( selected_row_position_list ) == 0:
			#行を選択していないなら、選択中のセルのある行を代わりに使用する
			selected_row_position = self.grid.GetGridCursorRow()
		else:
			#行が取得できていれば、それらのリストの最初の行番号だけにする
			selected_row_position = selected_row_position_list[0]

		row_col_plus_minus_dialog = RowColPlusMinusDialog()
		row_col_plus_minus_dialog.set_current_row( selected_row_position )

		if row_col_plus_minus_dialog.ShowModal() == wx.ID_OK:
			xy_plusminus = row_col_plus_minus_dialog.get_row_col_add_del()

			#役割をもたせた列（住所、氏名、敬称、可否）の最大値を求めておく
			special_labels = [ ["column-postalcode", "郵便番号" ], [ "column-address1", "宛先住所" ], [ "column-address2", "宛先住所（二列目）" ], [ "column-name1", "宛名" ], [ "column-name2", "宛名（二人目）" ], [ "column-company", "会社名" ], [ "column-department", "部署" ] ]
			special_column = [ self.column_etc_dictionary[ x[0] ] for x in special_labels ]
			if self.column_etc_dictionary[ "print-control" ] is True:
				special_column.append( self.column_etc_dictionary[ "print-control-column" ] )
			if self.column_etc_dictionary[ "enable-honorific-in-table" ] is True:
				special_column.append( self.column_etc_dictionary[ "column-honorific" ] )
			max_special_column = max( special_column )

			#行の加減
			if xy_plusminus[0] == "row":
				#行の追加
				if xy_plusminus[1] > 0:
					self.grid.AppendRows( xy_plusminus[1] )
					self.statusbar.SetStatusText( str( xy_plusminus[1] ) + "行追加しました" )

				#行の消去
				else:
					if self.grid.GetNumberRows() - abs( xy_plusminus[1] ) < 0:
						self.grid.DeleteRows( 0, self.grid.GetNumberRows() )
						self.statusbar.SetStatusText( "すべての行を消しました" )
					else:
						self.grid.DeleteRows( self.grid.GetNumberRows() - abs( xy_plusminus[1] ), abs( xy_plusminus[1] ) )
						self.statusbar.SetStatusText( str( abs( xy_plusminus[1] ) ) + "行削除しました" )

			#列の加減
			elif xy_plusminus[0] == "col":
				#列の追加
				if xy_plusminus[1] > 0:
					self.grid.AppendCols( xy_plusminus[1] )
					self.set_grid_labels() #追加した列のラベルがアルファベットになってしまうので、ラベルを貼り直す
					self.statusbar.SetStatusText( str( xy_plusminus[1] ) + "列追加しました" )

				#列の消去
				else:
					if max_special_column >= self.grid.GetNumberCols() - 1:
						self.statusbar.SetStatusText( "住所、氏名、敬称、可否といった役割列が、端の列まで占めているので削減しません" )

					#列を消すと役割列まで消してしまう場合、役割列の一つ外までのみ消す
					elif self.grid.GetNumberCols() - 1 - abs( xy_plusminus[1] ) < max_special_column:
						self.grid.DeleteCols( max_special_column + 1, self.grid.GetNumberCols() - max_special_column )
						self.statusbar.SetStatusText( "住所、氏名、敬称、可否といった役割列まで消去対象でしたが、その外までしか消しません" )
					elif self.grid.GetNumberCols() - abs( xy_plusminus[1] ) < 0:
						self.grid.DeleteCols( 0, self.grid.GetNumberCols() )
						self.statusbar.SetStatusText( "すべての列を消しました" )
					else:
						self.grid.DeleteCols( self.grid.GetNumberCols() - abs( xy_plusminus[1] ), abs( xy_plusminus[1] ) )
						self.statusbar.SetStatusText( str( abs( xy_plusminus[1] ) ) + "列削除しました" )

			#途中の指定された行に1行加える
			elif xy_plusminus[0] == "row-insert":
				self.grid.InsertRows( pos = selected_row_position, numRows = xy_plusminus[1] )
				self.statusbar.SetStatusText( str( selected_row_position + 1 ) + "行目に" + str( xy_plusminus[1] ) + "行追加しました" )

			#途中の指定された行のみ消す
			elif xy_plusminus[0] == "row-cut":
				if xy_plusminus[1] < self.grid.GetNumberRows():
					self.grid.DeleteRows( pos = selected_row_position, numRows = xy_plusminus[1] )
					self.statusbar.SetStatusText( str( selected_row_position + 1 ) + "行目から" + str( xy_plusminus[1] ) + "行削除しました" )

			#表（grid）がアクティブにならないのでSetFocusしておく
			self.grid.SetFocus()
			#履歴を更新する
			self.add_table_to_history( None )

		row_col_plus_minus_dialog.Destroy()


	#表の検索・置換
	def table_search( self, event ):

		search_dialog = SearchReplaceDialog()

		if search_dialog.ShowModal() == wx.ID_OK:

			#検索・置換の共通処理（検索語に該当するセルを検索する）
			search_word = search_dialog.get_search_word()
			if search_word == "":
				self.statusbar.SetStatusText( "検索語が空なので、検索しません" )
				search_dialog.Destroy()
				return False
			temporary_list = []

			#検索の本体といえる処理
			for j in range( self.grid.GetNumberRows() ):
				for i in range( self.grid.GetNumberCols() ):
					cell_value = self.grid.GetCellValue( j, i )
					if search_word in cell_value:
						temporary_list.append( [ i, j ] )

			#検索の結果、該当セルがなかった場合
			if temporary_list == []:
				self.statusbar.SetStatusText( "「" + search_word + "」が見つかりませんでした" )
			#置換
			elif search_dialog.get_replace_mode() is True:

				for find_cell in temporary_list:
					#該当するセルの内容を取得する
					cell_value = self.grid.GetCellValue( find_cell[1], find_cell[0] )

					#ダイアログから置き換える単語を取得する
					replace_word = search_dialog.get_replace_word()

					#取得した内容を置換し、セルを書き換える
					replaced_value = cell_value.replace( search_word, replace_word )
					self.grid.SetCellValue( find_cell[1], find_cell[0], replaced_value )

				#置換が一通り終わったら、履歴を更新する
				self.add_table_to_history( None )

				self.statusbar.SetStatusText( str( len( temporary_list ) ) + "件を置換しました" )

			#検索（というより、前段階で洗いだした結果の格納と表示）
			else:

				self.find_list = temporary_list

				self.statusbar.SetStatusText( "「" + search_word + "」が" + str( len( self.find_list ) ) + "件該当しました （ F3 で次の該当セル、 Shift + F3 で前の該当セルに移動 ）" )
				#ヒットした最初のセルに移動する（Createなどと同様に縦位置、横位置の順）
				cell_point = ( self.find_list[0][1], self.find_list[0][0] )
				self.grid.SetGridCursor( cell_point[0], cell_point[1] ) #該当セルを選択状態にする
				self.grid.MakeCellVisible( cell_point[0], cell_point[1] ) #該当セルまでスクロールさせる
				#（上記の2行以外にSetFocusの必要もあるが、ダイアログ終了後に置くことにする）

				self.current_find_number = 0 #検索結果内の表示位置をリセット

		search_dialog.Destroy()

		#検索する・しないに関わらず、表（grid）をアクティブにしておく
		self.grid.SetFocus()


	#表のフォント変更
	def change_table_font( self, event ):
		selected_font = self.combobox_table_font.GetValue()
		current_fontsize = self.grid.GetDefaultCellFont().GetPointSize()

		#先に設定値の辞書を書き換えておく
		self.software_setting[ "table-font" ] = selected_font

		if selected_font == "":
			#コンボボックスの空欄を選んだら、起動時に取得したデフォルトフォントを適用する
			changed_fontdata = wx.Font( pointSize = current_fontsize, family = wx.DEFAULT, style = wx.NORMAL, weight = wx.NORMAL, underline=False, faceName = self.table_systemdefault_font )
			self.grid.SetDefaultCellFont( changed_fontdata )

		else:
			changed_fontdata = wx.Font( pointSize = current_fontsize, family = wx.DEFAULT, style = wx.NORMAL, weight = wx.NORMAL, underline=False, faceName = selected_font )
			self.grid.SetDefaultCellFont( changed_fontdata )

			#表の行・列のサイズを自動調整する
			self.grid.AutoSizeColumns()
			self.grid.AutoSizeRows()


	#表のフォントサイズ変更
	def change_table_fontsize( self, event ):
		input_fontsize = self.table_fontsize_input.GetValue()
		current_font = self.software_setting[ "table-font" ]

		#先に設定値の辞書を書き換えておく
		self.software_setting[ "table-fontsize" ] = input_fontsize

		if input_fontsize > 0:

			if current_font == "":
				changed_fontdata = wx.Font( pointSize = input_fontsize, family = wx.DEFAULT, style = wx.NORMAL, weight = wx.NORMAL, underline=False )
			else:
				changed_fontdata = wx.Font( pointSize = input_fontsize, family = wx.DEFAULT, style = wx.NORMAL, weight = wx.NORMAL, underline=False, faceName = current_font )

			self.grid.SetDefaultCellFont( changed_fontdata )

			#フォントあるいはフォントサイズを変更したので、表の行・列のサイズを調整する
			self.grid.AutoSizeColumns()
			self.grid.AutoSizeRows()


	#-----表関係の関数はここまで-----


	#宛名印刷の関数をスレッドとして開始あるいは中断する
	def atenaprint_thread_control( self, event ):
		current_label = self.print_button.GetLabel()

		if current_label == "宛名印刷する":

			#A4を超える用紙サイズであれば、印刷を始める前に警告ダイアログを出す
			if self.paper_size_data[ "width" ] > 210:
				a4over_warning_dialog = wx.MessageDialog( self,  message = "用紙の幅がA4を超える大きさになっていますが、プリンターはそのサイズの印刷に対応していますか？\n\n対応していなければいつまで待っても印刷されないかもしれないので注意してください。", caption = "A4超の用紙の印刷についての注意", style = wx.OK | wx.ICON_EXCLAMATION )
				a4over_warning_dialog.ShowModal()
				a4over_warning_dialog.Destroy()

			elif self.paper_size_data[ "height" ] > 297:
				a4over_warning_dialog = wx.MessageDialog( self,  message = "封筒フタ高さを含めた用紙の高さがA4を超えていますが、プリンターはそのサイズの印刷に対応していますか？\n\n対応していなければいつまで待っても印刷されないかもしれないので注意してください。\n\n（ただし縦に長い分には、かなり長くても印刷を受け付けてくれる可能性はあります）", caption = "A4超の用紙の印刷についての注意", style = wx.OK | wx.ICON_EXCLAMATION )
				a4over_warning_dialog.ShowModal()
				a4over_warning_dialog.Destroy()

			if "はがき" in self.paper_size_data[ "category" ]:

				if self.column_etc_dictionary[ "upside-down-print" ] is True:
					upsidedown_alert_dialog = wx.MessageDialog( self,  message = "はがき印刷なのに（封筒印刷用の）上下反転印刷モードになっています。\n反転を解除して印刷しますか？\n\n（どちらにせよ、当ソフトの印刷方向とプリンターへ用紙を差し込む方向が一致しているかも再確認しておいてください。）", caption = "反転印刷の警告と確認", style = wx.YES_NO | wx.ICON_QUESTION )

					if upsidedown_alert_dialog.ShowModal() == wx.ID_YES:
						self.upsidedown_print_checkbox.SetValue( False )
						self.column_etc_dictionary[ "upside-down-print" ] = False

					upsidedown_alert_dialog.Destroy()

			else:

				if self.column_etc_dictionary[ "upside-down-print" ] is False:
					upsidedown_alert_dialog = wx.MessageDialog( self,  message = "（用紙がはがきではないので）おそらく封筒印刷だと思われます。\nしかし通常方向の印刷モードになっています。\n\n封筒のフタ部分から印刷しようとすると、フタの高さだけ印刷位置がずれたり、プリンターの用紙送りが正常に行えないかもしれません。\n上下反転印刷モードにして印刷しますか？\n\n（どちらにせよ、当ソフトの印刷方向とプリンターへ用紙を差し込む方向が一致しているかも再確認しておいてください。）", caption = "反転印刷の警告と確認", style = wx.YES_NO | wx.ICON_QUESTION )

					if upsidedown_alert_dialog.ShowModal() == wx.ID_YES:
						self.upsidedown_print_checkbox.SetValue( True )
						self.column_etc_dictionary[ "upside-down-print" ] = True

					upsidedown_alert_dialog.Destroy()

			current_label = self.print_button.GetLabel()

			self.print_button.SetLabel( "印刷を中止" )
			self.thread = threading.Thread( target = self.atena_print, args = ( None, ), daemon=True )
			self.thread.setDaemon( True )
			self.thread.start()

		#ボタンのラベルが「宛名印刷する」以外、つまり「印刷を中止」のときに押したら、中止する
		else:
			self.print_stop_flag = True


	#宛名の印刷
	def atena_print( self, event ):
		current_table = self.get_current_table_list()

		print_size = "Custom." + str( self.paper_size_data[ "width" ] ) + "x" + str( self.paper_size_data[ "height" ] ) + "mm"

		min_line = self.print_start_line.GetValue() - 1 #GUI上の行番号は1,2,3...だが処理上の行は0,1,2...なので-1しておく
		max_line = self.print_end_line.GetValue() - 1

		#開始行（最小値）が終端行（最大値）より大きければ、開始と終端を入れ替える
		if min_line > max_line:
			temporary_values = [ min_line, max_line ]
			max_line, min_line = temporary_values

		#開始行、終端行が全行数を超えないようにする
		if min_line >= len( current_table ):
			min_line = len( current_table ) - 1
		if max_line >= len( current_table ):
			max_line = len( current_table ) - 1

		for line_number in range( min_line, max_line + 1 ):
			#スレッド化した関数の中で直接GUI操作するとウィンドウが異常終了する場合があるので
			#関数呼び出しをwx.CallAfterで包む
			wx.CallAfter( self.statusbar.SetStatusText, str( line_number + 1 ) + "行目を処理中です" )
			current_line = current_table[ line_number ]
			print_check = ""

			#特定の列の内容で印刷の可否を判別する場合
			if self.column_etc_dictionary[ "print-control" ] is True:
				print_flag = current_line[ self.column_etc_dictionary[ "print-control-column" ] ]
				printsign_value = self.column_etc_dictionary[ "print-sign" ]

				if self.column_etc_dictionary[ "print-or-ignore" ] == "ignore":
					#列の内容が一致したら無視する設定において、一致した場合
					if print_flag == printsign_value:
						wx.CallAfter( self.statusbar.SetStatusText, str( line_number + 1 ) + "行目は印刷しません（" + str( self.column_etc_dictionary[ "print-control-column" ] + 1 ) + "列目が「" + print_flag + "」）" )
						print_check = "no"

				else:
					#列の内容が一致したら印刷する（一致「しなかったら無視する」）設定において、一致しなかった場合
					if print_flag != printsign_value:
						wx.CallAfter( self.statusbar.SetStatusText, str( line_number + 1 ) + "行目は印刷しません（" + str( self.column_etc_dictionary[ "print-control-column" ] + 1 ) + "列目が「" + print_flag + "」）" )
						print_check = "no"

			if print_check != "no":
				print_data = ""

				dest_postal_code = current_line[ self.column_etc_dictionary[ "column-postalcode" ] ]
				dest_address1 = current_line[ self.column_etc_dictionary[ "column-address1" ] ]
				dest_address2 = current_line[ self.column_etc_dictionary[ "column-address2" ]  ]
				dest_name1 = current_line[ self.column_etc_dictionary[ "column-name1" ] ]
				dest_name2 = current_line[ self.column_etc_dictionary[ "column-name2" ] ]

				dest_company = current_line[ self.column_etc_dictionary[ "column-company" ] ]
				dest_department = current_line[ self.column_etc_dictionary[ "column-department" ] ]

				#敬称
				#表中の?列目の敬称を使用する設定で、指定された列が実在して（列数の範囲内）、その列に記述があった場合
				if self.column_etc_dictionary[ "enable-honorific-in-table" ] is True and self.column_etc_dictionary[ "column-honorific" ] < len( current_line ) and current_line[ self.column_etc_dictionary[ "column-honorific" ] ] != "":
					honorific = current_line[ self.column_etc_dictionary[ "column-honorific" ] ]

				#（前段のifで）表の敬称を使わない設定か敬称の列が空白で、なおかつ設定で指定されている敬称を適用する場合
				elif self.column_etc_dictionary[ "enable-default-honorific" ] is True:
					honorific = self.column_etc_dictionary[ "default-honorific" ]

				else:
					honorific = ""

				print_data = { "postal-code" : dest_postal_code, "name1" : dest_name1, "name2" : dest_name2, "address1" : dest_address1, "address2" : dest_address2, "company" : dest_company, "department" : dest_department, "honorific" : honorific, "our-postal-code" : self.our_data[ "our-postalcode-data" ], "our-name1" : self.our_data[ "our-name1-data" ], "our-name2" : self.our_data[ "our-name2-data" ], "our-address1" : self.our_data[ "our-address1-data" ], "our-address2" : self.our_data[ "our-address2-data" ] }

				print_grayscale_image = self.image_generator.get_cutted_atena_image( data_dict = print_data, space_tblr_mm_list = self.column_etc_dictionary[ "printer-space-top,bottom,left,right" ], cutted_atena_image_upside_down = self.column_etc_dictionary[ "upside-down-print" ] )

				try:
					pil_printing( pil_image = print_grayscale_image, paper_size = print_size, upside_down = self.column_etc_dictionary[ "upside-down-print" ] )
				except Exception as e:
					# 印刷エラーが発生した場合、ダイアログで表示して処理を中止
					error_message = str(e)
					wx.CallAfter( self.print_button.SetLabel, "宛名印刷する" )
					wx.CallAfter( self.SetStatusText, "印刷エラーが発生しました" )
					wx.CallAfter( self.stop_message_dialog, error_message )
					self.print_stop_flag = False
					return False

			#印刷中止用の変数がTrueなら、ステータスバーやダイアログで中止を表明して関数を終了する
			if self.print_stop_flag is True:
				wx.CallAfter( self.SetStatusText, str( line_number + 1 ) + "行目までで印刷を中止しました" )
				#「印刷中止」にしていたボタンのラベルを元に戻しておく
				wx.CallAfter( self.print_button.SetLabel, "宛名印刷する" )
				wx.CallAfter( self.stop_message_dialog, str( line_number + 1 ) + "行目までで印刷を中止しました" )

				self.print_stop_flag = False
				return False

		#印刷終了後に、「中止ボタン」に変えていた印刷ボタンのラベルを元に戻しておく
		wx.CallAfter( self.print_button.SetLabel, "宛名印刷する" )
		#ステータスバーを空欄に戻す
		wx.CallAfter( self.statusbar.SetStatusText, "" )


	#印刷を中止した場合にダイアログを表示する
	def stop_message_dialog( self, message = "" ):
		error_dialog =  wx.MessageDialog( self,  message = message, caption = "Error", style = wx.OK | wx.ICON_ERROR )
		error_dialog.ShowModal()
		error_dialog.Destroy()


	#郵便番号検索
	def postalcode_search( self, event ):
		official_postalcode_file_name = "KEN_ALL.CSV"
		official_postalcode_file_path = os.path.join ( os.path.split( sys.argv[0] )[0], official_postalcode_file_name )

		if os.path.isfile( official_postalcode_file_path ) is True:
			postalcode_dialog = PostalcodeSearchDialog( official_postalcode_file_path )
			postalcode_dialog.ShowModal()
			postalcode_dialog.Destroy()

		else:
			self.statusbar.SetStatusText( "郵政公社から配布されている「" + official_postalcode_file_name + "」ファイルが同じディレクトリにないので、郵便番号検索はできません" )


	#履歴の移動
	def goto_history_point( self, event ):
		self.history_dialog = HistoryDialog( self.current_history_position, len( self.table_history ) )

		if self.history_dialog.ShowModal() == wx.ID_OK:
			self.current_history_position = self.history_dialog.get_list_value()
			self.set_table( self.table_history[ self.current_history_position ] )
			self.set_grid_labels() #行数・列数が変化したかもしれないのでラベルを貼り直す
			self.statusbar.SetStatusText( "履歴番号" + str( self.current_history_position ) + "（最新が0）にやり直しました" )

		self.history_dialog.Destroy()

	#メインウィンドウを閉じて終了する（実際には終了前に設定変更チェックに入る）
	def window_close( self, event ):
		self.Close()

	#終了時に設定変更があったかチェックして、変更があれば保存するか尋ねる
	#表の変更が保存されているかのチェックも追加した。
	def check_at_close( self, event ):

		#まず、設定が（保存された状態の）控えから変更されているかのチェック
		if self.savepoint_paper_size_data != self.paper_size_data or self.savepoint_parts_dictionary != self.image_generator.get_parts_dictionary() or self.savepoint_software_setting != self.software_setting or self.savepoint_column_etc_dictionary != self.column_etc_dictionary or self.savepoint_our_data != self.our_data:
			question_dialog = wx.MessageDialog( self,  message = "設定が変更されています。設定を保存しますか？\nNoで保存せずに終了します。", caption = "終了時の設定変更チェック", style = wx.YES_NO | wx.ICON_QUESTION )

			if question_dialog.ShowModal() == wx.ID_YES:
				self.save_settings( None )

			question_dialog.Destroy()


		#次に、表の内容が控えと同一かチェックする
		current_table = self.get_current_table_list()

		if current_table != self.table_checkpoint:
			question_dialog = wx.MessageDialog( parent = self, message = "現在の表の内容が変更されていますが保存されていません。\n\n表の内容を保存しますか？\nNoで保存せずに終了します。", caption = "表内容の変更に関する確認", style = wx.YES_NO | wx.ICON_QUESTION )

			if question_dialog.ShowModal() == wx.ID_YES:
				self.save_csv_file( None )
			question_dialog.Destroy()

		#設定と表内容の変更確認が終わったので、本当にウィンドウを閉じる
		self.Destroy()


	#以下、キー操作の処理

	#Shift + F3で、検索結果を逆にたどる
	def key_shift_f3( self, event ):

		if self.find_list == []:
			self.statusbar.SetStatusText( "検索結果が空なので、移動できません" )

		elif self.current_find_number == 0:
			self.current_find_number = len( self.find_list ) - 1
			cell_point = ( self.find_list[ len( self.find_list ) - 1 ][1], self.find_list[ len( self.find_list ) - 1 ][0] )
			self.grid.SetGridCursor( cell_point[0], cell_point[1] )
			self.grid.MakeCellVisible( cell_point[0], cell_point[1] )
			self.statusbar.SetStatusText( "最初の検索結果まで来ていたので、最後に飛びました［" + str( self.find_list[ self.current_find_number ][1] + 1 ) + "行、" + str( self.find_list[ self.current_find_number ][0] + 1 ) + "列 ］，（ " + str( self.current_find_number + 1 ) + " / " + str( len( self.find_list ) ) + " 番目）" )

		else:
			self.current_find_number -= 1
			cell_point = ( self.find_list[ self.current_find_number ][1], self.find_list[ self.current_find_number ][0] )
			self.grid.SetGridCursor( cell_point[0], cell_point[1] )
			self.grid.MakeCellVisible( cell_point[0], cell_point[1] )
			self.statusbar.SetStatusText( "検索結果［ " + str( self.find_list[ self.current_find_number ][1] + 1 ) + "行、" + str( self.find_list[ self.current_find_number ][0] + 1 ) + "列 ］，（ " + str( self.current_find_number + 1 ) + " / " + str( len( self.find_list ) ) + " 番目）" )

	#shiftのないF3で、検索結果を順送り
	def key_nomal_f3( self, event ):

		if self.find_list == []:
			self.statusbar.SetStatusText( "検索結果が空なので、移動できません" )

		elif self.current_find_number == len( self.find_list ) - 1:
			self.current_find_number = 0
			cell_point = ( self.find_list[0][1], self.find_list[0][0] )
			self.grid.SetGridCursor( cell_point[0], cell_point[1] )
			self.grid.MakeCellVisible( cell_point[0], cell_point[1] )
			self.statusbar.SetStatusText( "最後の検索結果まで来ていたので、最初に戻りました［" + str( self.find_list[ self.current_find_number ][1] + 1 ) + "行、" + str( self.find_list[ self.current_find_number ][0] + 1 ) + "列 ］，（ " + str( self.current_find_number + 1 ) + " / " + str( len( self.find_list ) ) + " 番目）" )

		else:
			self.current_find_number += 1
			cell_point = ( self.find_list[ self.current_find_number ][1], self.find_list[ self.current_find_number ][0] )
			self.grid.SetGridCursor( cell_point[0], cell_point[1] )
			self.grid.MakeCellVisible( cell_point[0], cell_point[1] )
			self.statusbar.SetStatusText( "検索結果［ " + str( self.find_list[ self.current_find_number ][1] + 1 ) + "行、" + str( self.find_list[ self.current_find_number ][0] + 1 ) + "列 ］，（ " + str( self.current_find_number + 1 ) + " / " + str( len( self.find_list ) ) + " 番目）" )

	#Ctrl + Fで、検索ダイアログを起動する
	def key_ctrl_f( self, event ):
		self.table_search( None )

	#Shift + Ctrl + Zで、履歴を進める
	def key_ctrl_shift_z( self, event ):

		if self.current_history_position == 0:
			self.statusbar.SetStatusText( "現在、履歴の最新位置（履歴番号0）にあるので、やり直しできません" )
		else:
			self.current_history_position -= 1
			self.set_table( self.table_history[ self.current_history_position ] )
			self.set_grid_labels() #行数・列数が変化したかもしれないのでラベルを貼り直す
			self.statusbar.SetStatusText( "履歴番号" + str( self.current_history_position ) + "（最新が0）にやり直しました" )

	#ShiftのないCtrl + Zで、履歴を戻る
	def key_ctrl_z( self, event ):

		if self.current_history_position >= len( self.table_history ) - 1:
			self.statusbar.SetStatusText( "現在、ストックの中で最古の履歴（履歴番号" + str( self.current_history_position ) + "）なので、これ以上戻れません" )
		else:
			self.current_history_position += 1
			self.set_table( self.table_history[ self.current_history_position ] )
			self.set_grid_labels() #行数・列数が変化したかもしれないのでラベルを貼り直す
			self.statusbar.SetStatusText( "履歴番号" + str( self.current_history_position ) + "（最新が0）に戻しました" )

	#キー操作の処理はここまで



#-----ここまで、メインウィンドウ-----


#印刷イメージ確認用のダイアログ
class AtenaPreviewDialog( wx.Dialog ):

	def __init__( self, paper_data_dict, destination_list, column_data, our_data, min_line_int, max_line_int, image_generator_instance, space_tblr_mm_list, cutted_atena_image_upside_down,  current_row = 0 ):

		wx.Dialog.__init__( self, None, -1, "宛名印刷イメージの確認", size = ( 500, 720 ) )

		self.dest_list = copy.copy( destination_list ) #現在の表の内容のすべて
		self.column_dict = copy.deepcopy( column_data ) #column_etc_dictionaryの内容、つまり各列と宛名、住所などの対応関係やその他の情報
		self.our_dict = copy.deepcopy( our_data ) #our_data辞書、つまり差出人の情報
		self.image_generator = copy.deepcopy( image_generator_instance )
		self.space_list = copy.copy( space_tblr_mm_list )
		self.cutted_atena_image_upside_down = cutted_atena_image_upside_down

		self.min_line = min_line_int
		self.max_line = max_line_int

		#開始行 > 終端行なら、値を入れ替える
		if self.min_line > self.max_line:
			temporary_values = [ self.min_line, self.max_line ]
			self.max_line, self.min_line = temporary_values

		#最大値が全行数を超えないようにする
		if self.min_line > len( destination_list ):
			self.min_line = len( destination_list )
		if self.max_line > len( destination_list ):
			self.max_line = len( destination_list )

		#調整した行範囲をもとにしてダイアログのタイトルを再設定する
		self.SetTitle( "宛名印刷イメージの確認（" + str( self.min_line ) + "行目〜" + str( self.max_line ) + "行目）" )

		#宛名イメージの範囲がわかりやすいように、パネルを貼ってダイアログ全体を灰色にする
		self.color_panel = wx.Panel( self, wx.ID_ANY )
		self.color_panel.SetBackgroundColour( "#CCCCCC" )

		#宛名イメージを表示するパネル
		self.image_panel = wx.Panel( self.color_panel, wx.ID_ANY )
		self.image_panel.SetBackgroundColour( "#CCCCCC" )
		self.image_panel.Bind( wx.EVT_PAINT, self.OnPaint )
		#パネルの大きさが変わったら画像がリサイズされるようにバインド
		self.image_panel.Bind( wx.EVT_SIZE, self.send_adjust_image_with_panel )

		#貼り付ける画像イメージ（これに、それぞれの宛名画像を代入してRefresh()する）
		self.preview_image = wx.Image( 100, 100 )
		self.resized_preview_image = self.preview_image.Copy()
		self.preview_bitmap_image = self.resized_preview_image.ConvertToBitmap()
		self.image_position = [ 0, 0 ]

		#印刷しない行において代わりに表示する画像を用意する
		no_print_text = "この行は印刷しません"

		no_print_image_pil = Image.new( "RGB", ( 300, 30 ), ( 255, 255, 255 ) )
		fnt = ImageFont.truetype( self.image_generator.get_parts_data( "fontfile" ), 20, encoding="unic" )
		textimage_draw = ImageDraw.Draw( no_print_image_pil )
		textimage_draw.text( ( 0, 0 ), no_print_text, font=fnt, fill="black" )

		self.no_print_image = wx.Image( no_print_image_pil.size[0], no_print_image_pil.size[1] ) #wx.Imageに変換
		self.no_print_image.SetData( no_print_image_pil.convert('RGB').tobytes() )

		#何行目の宛名イメージを表示するか、切り替えるためのスライダー
		self.slider = wx.Slider( self.color_panel, style = wx.SL_HORIZONTAL | wx.SL_LABELS )
		#スライダーの最小最大値の設定は、まず最大値の拡大から行う
		#最小値を先に設定すると、最小値が変更前のMax値(デフォルトの100)を超えた場合にエラーで0にされてしまう
		self.slider.SetMax( self.max_line )
		self.slider.SetMin( self.min_line )

		#現在選択中の行にデフォルト位置を合わせておく（current_rowは0行から数えた値なので、+1で1行目からの行数に修正する）
		if current_row + 1 >= self.min_line and current_row + 1 <= self.max_line:
			self.slider.SetValue( current_row + 1 )

		#最小値と最大値が同じ（対象画像がひとつだけ）でもスライダーが動くので無効化
		if self.min_line >= self.max_line:
			self.slider.Disable()

		#スライダーを枠（StaticBoxSizer）に入れる
		if self.min_line >= self.max_line:
			self.line_slider_sbox = wx.StaticBox( self.color_panel, wx.ID_ANY, "スライダーは無効にしています" )
		else:
			self.line_slider_sbox = wx.StaticBox( self.color_panel, wx.ID_ANY, "↓のSliderを動かすと、" + str( self.min_line ) + "行〜" + str( self.max_line ) + "行まで表示を切り替えます" )
		self.line_slider_sizer = wx.StaticBoxSizer( self.line_slider_sbox, wx.VERTICAL )
		self.line_slider_sizer.Add( self.slider, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		self.print_disable_message = wx.StaticText( self.color_panel, wx.ID_ANY, "" )

		self.slider.Bind( wx.EVT_SLIDER, self.replace_image )

		button = wx.Button( self.color_panel, wx.ID_OK, "OK" )
		button.SetDefault()

		sizer = wx.BoxSizer( wx.VERTICAL )
		sizer.Add( self.print_disable_message, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		sizer.Add( self.image_panel, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		sizer.Add( wx.StaticText( self.color_panel, wx.ID_ANY, "用紙 ： " + paper_data_dict[ "category" ] + " ( " + str( paper_data_dict[ "width" ] ) + "mm x " + str( paper_data_dict[ "height" ] ) + "mm )" ), 0, wx.ALIGN_CENTER_HORIZONTAL | wx.TOP, 4 )
		sizer.Add( wx.StaticText( self.color_panel, wx.ID_ANY, "※ 上下左右の余白が取り除かれた形で表示されています" ), 0, wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM, 10 )
		sizer.Add( self.line_slider_sizer, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		sizer.Add( button, 0, wx.ALIGN_CENTER )
		self.color_panel.SetSizer(sizer)

		base_sizer = wx.BoxSizer( wx.VERTICAL )
		base_sizer.Add( self.color_panel, 1, wx.ALL | wx.EXPAND, 10 )
		self.SetSizer( base_sizer )

		self.show_preview_image()

	def OnPaint( self, event=None ):
		deviceContext = wx.PaintDC( self.image_panel )
		deviceContext.Clear()
		deviceContext.SetPen( wx.Pen( wx.BLACK, 4 ) )
		deviceContext.DrawBitmap( self.preview_bitmap_image, self.image_position[0], self.image_position[1] )


	#スライダーを動かしたら、その行のデータの画像を取得して切り替える
	def replace_image( self, event ):
		self.show_preview_image()

	#余白カット済みの宛名イメージを取得してパネルに表示する
	def show_preview_image( self ):
		line_number = self.slider.GetValue()

		#表の中で、スライダーから指定された行のデータ。
		current_destination = self.dest_list[ line_number - 1 ] #GUI上の行番号は1,2,3...だが処理上の行は0,1,2...なので-1しておく

		#特定の列の内容で印刷の可否を判別する場合
		if self.column_dict[ "print-control" ] is True:
			print_flag = current_destination[ self.column_dict[ "print-control-column" ] ]
			printsign_value = self.column_dict[ "print-sign" ]

			if self.column_dict[ "print-or-ignore" ] == "ignore":
				#列の内容が一致したら無視する設定において、一致した場合
				if print_flag == printsign_value:
					self.print_disable_message.SetLabel( "この" + str( line_number ) + "行目は印刷しません（" + str( self.column_dict[ "print-control-column" ] + 1 ) + "列目が「" + print_flag + "」）" )
					self.preview_image = self.no_print_image
					self.adjust_image_with_panel()
					return False

			else:
				#列の内容が一致したら印刷する（一致「しなかったら無視する」）設定において、一致しなかった場合
				if print_flag != printsign_value:
					self.print_disable_message.SetLabel( "この" + str( line_number ) + "行目は印刷しません（" + str( self.column_dict[ "print-control-column" ] + 1 ) + "列目が「" + print_flag + "」）" )
					self.preview_image = self.no_print_image
					self.adjust_image_with_panel()
					return False

		self.print_disable_message.SetLabel( str( line_number ) + "行目" )
		print_data = self.make_current_data( current_destination )

		preview_grayscale_image = self.image_generator.get_cutted_atena_image( print_data, self.space_list, return_pasted_image = True, cutted_atena_image_upside_down = self.cutted_atena_image_upside_down )
		preview_color_image = ImageOps.colorize( preview_grayscale_image, ( 0, 0, 0 ), ( 255, 255, 255 ) )

		self.preview_image = wx.Image( preview_color_image.size[0], preview_color_image.size[1] )
		self.preview_image.SetData( preview_color_image.convert('RGB').tobytes() )
		self.resized_preview_image =self.preview_image.Copy()

		#このままパネルのサイズに合わせて画像をリサイズしてから表示しても
		#ダイアログ起動直後に限っては、画像が極小になりまともに表示されない。
		#初回はFrameに応じてパネルが拡大する前なのでパネルサイズが20X20と小さいため。
		#パネルサイズの変動に応じて画像をリサイズする関数も作ってバインドしないといけない。

		#パネルよりも横長なら、パネルの幅に合わせてリサイズする（初回以外でのみ有効）
		if float( self.resized_preview_image.GetSize()[0] ) / float( self.resized_preview_image.GetSize()[1] ) > float( self.image_panel.GetSize()[0] ) / float( self.image_panel.GetSize()[1] ):

			self.resized_preview_image.Rescale( self.image_panel.GetSize()[0], int( self.resized_preview_image.GetSize()[1] * self.image_panel.GetSize()[0] / self.resized_preview_image.GetSize()[0] ), quality=wx.IMAGE_QUALITY_HIGH )
			self.image_position[0] = 0
			self.image_position[1] = int( ( self.image_panel.GetSize()[1] - self.resized_preview_image.GetSize()[1] ) / 2 )

		else:
			#パネルよりも縦長なら、パネルの高さに合わせてリサイズする
			self.resized_preview_image.Rescale( int( self.resized_preview_image.GetSize()[0] * self.image_panel.GetSize()[1] / self.resized_preview_image.GetSize()[1] ), self.image_panel.GetSize()[1], quality=wx.IMAGE_QUALITY_HIGH )
			self.image_position[0] = int( ( self.image_panel.GetSize()[0] - self.resized_preview_image.GetSize()[0] ) / 2 )
			self.image_position[1] = 0

		self.preview_bitmap_image =  self.resized_preview_image.ConvertToBitmap()
		self.Refresh()

	#パネルの大きさが変わったら、パネルに合わせてサンプル画像をリサイズする
	def send_adjust_image_with_panel( self, event ):
		self.adjust_image_with_panel()

	#パネルに合わせてサンプル画像をリサイズする
	def adjust_image_with_panel( self ):
		#パネルサイズの変更に合わせてリサイズするたびに、元のself.preview_imageから再度コピーする（copy.deepcopyは使えない）
		self.resized_preview_image = self.preview_image.Copy()

		#パネルよりも横長なら、パネルの幅に合わせてリサイズする
		if float( self.resized_preview_image.GetSize()[0] ) / float( self.resized_preview_image.GetSize()[1] ) > float( self.image_panel.GetSize()[0] ) / float( self.image_panel.GetSize()[1] ):

			self.resized_preview_image.Rescale( self.image_panel.GetSize()[0], int( self.resized_preview_image.GetSize()[1] * self.image_panel.GetSize()[0] / self.resized_preview_image.GetSize()[0] ), quality=wx.IMAGE_QUALITY_HIGH )
			self.image_position[0] = 0
			self.image_position[1] = int( ( self.image_panel.GetSize()[1] - self.resized_preview_image.GetSize()[1] ) / 2 )

		else:
			#パネルよりも縦長なら、パネルの高さに合わせてリサイズする
			self.resized_preview_image.Rescale( int( self.resized_preview_image.GetSize()[0] * self.image_panel.GetSize()[1] / self.resized_preview_image.GetSize()[1] ), self.image_panel.GetSize()[1], quality=wx.IMAGE_QUALITY_HIGH )
			self.image_position[0] = int( ( self.image_panel.GetSize()[0] - self.resized_preview_image.GetSize()[0] ) / 2 )
			self.image_position[1] = 0

		self.preview_bitmap_image =  self.resized_preview_image.ConvertToBitmap()
		self.Refresh()

	#スライダーが示す位置の行から住所氏名などのデータ辞書を構築する
	def make_current_data( self, current_destination ):

		d_postal_code = current_destination[ self.column_dict[ "column-postalcode" ] ]
		d_address1 = current_destination[ self.column_dict[ "column-address1" ] ]
		d_address2 = current_destination[ self.column_dict[ "column-address2" ]  ]
		d_name1 = current_destination[ self.column_dict[ "column-name1" ] ]
		d_name2 = current_destination[ self.column_dict[ "column-name2" ] ]

		d_company = current_destination[ self.column_dict[ "column-company" ] ]
		d_department = current_destination[ self.column_dict[ "column-department" ] ]

		#敬称
		#表中の?列目の敬称を使用する設定で、指定された列が実在して（列数の範囲内）、その列に記述があった場合
		if self.column_dict[ "enable-honorific-in-table" ] is True and self.column_dict[ "column-honorific" ] < len( current_destination ) and current_destination[ self.column_dict[ "column-honorific" ] ] != "":
			honorific = current_destination[ self.column_dict[ "column-honorific" ] ]

		#（前段のifで）表の敬称を使わない設定か敬称の列が空白で、なおかつ設定で指定されている敬称を適用する場合
		elif self.column_dict[ "enable-default-honorific" ] is True:
			honorific = self.column_dict[ "default-honorific" ]

		else:
			honorific = ""

		print_data = { "postal-code" : d_postal_code, "name1" : d_name1, "name2" : d_name2, "address1" : d_address1, "address2" : d_address2, "company" : d_company, "department" : d_department, "honorific" : honorific, "our-postal-code" : self.our_dict[ "our-postalcode-data" ], "our-name1" : self.our_dict[ "our-name1-data" ], "our-name2" : self.our_dict[ "our-name2-data" ], "our-address1" : self.our_dict[ "our-address1-data" ], "our-address2" : self.our_dict[ "our-address2-data" ] }

		return print_data


#郵便番号検索のダイアログ
class PostalcodeSearchDialog( wx.Dialog ):

	def __init__( self, csv_path ):
		wx.Dialog.__init__( self, None, -1, "郵便番号検索", size = ( 900, 600 ) )

		self.code_address_list = []
		self.code_address_linelist = []

		#公式の郵便番号CSVファイルがあるなら、郵便番号検索用に読み込む
		if os.path.isfile( csv_path ):
			self.code_address_list = csv_to_list( csv_path )

			for i in range( len( self.code_address_list ) ):
				hit_text =  "郵便番号：" + self.code_address_list[ i ][2] + "、地域名：" + self.code_address_list[ i ][6] + self.code_address_list[ i ][7]
				if self.code_address_list[ i ][8] == "以下に掲載がない場合":
					hit_text += "　（以下に掲載がない場合）"
				else:
					hit_text += self.code_address_list[ i ][8]

				self.code_address_linelist.append( hit_text )

		#ここから、このダイアログのGUI部分
		self.input_search_word = wx.TextCtrl( self, wx.ID_ANY, style=wx.TE_PROCESS_ENTER )
		self.search_button = wx.Button( self, wx.ID_ANY, "検索" )
		self.hitting_message = wx.StaticText( self, wx.ID_ANY, "" )

		#検索語の入力欄でEnterを押すと、検索を始めるようにバインド
		self.input_search_word.Bind( wx.EVT_TEXT_ENTER, self.search_and_display )

		#1行にまとめる
		self.input_and_go = wx.BoxSizer( wx.HORIZONTAL )
		self.input_and_go.Add( self.input_search_word, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		self.input_and_go.Add( self.search_button, 0, wx.FIXED_MINSIZE )
		self.input_and_go.Add( self.hitting_message, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )

		#検索結果の表示領域
		self.find_result = wx.TextCtrl( self, style = wx.TE_MULTILINE | wx.TE_READONLY )

		#バインド
		self.search_button.Bind( wx.EVT_BUTTON, self.search_and_display )

		self.ok_button = wx.Button( self, wx.ID_OK, "OK" )

		sizer = wx.BoxSizer( wx.VERTICAL )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, "入力欄に郵便番号（半角数字7桁）あるいは地域名の一部分を打ち込んで、Enterを押すか検索ボタンを押してください" ), 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, "　　なお、検索結果は必要な部分を手動でコピーしてお使いください" ), 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )
		sizer.Add( self.input_and_go, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )
		sizer.Add( self.find_result, 1, wx.ALL | wx.EXPAND, 10 )
		sizer.Add( self.ok_button, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.FIXED_MINSIZE )
		self.SetSizer( sizer )


	def search_and_display( self, event ):
		hit_lines = ""
		hit = 0
		search_word = self.input_search_word.GetValue()

		for line_num in range( len( self.code_address_linelist ) ):
			#検索語が含まれている行に当たったら、内容を格納する
			if search_word in self.code_address_linelist[ line_num ]:
				hit_lines += self.code_address_linelist[ line_num ] + "\n\n"
				hit += 1

		if hit_lines == "":
			#1件も該当しなかったら、結果表示を空白化
			self.find_result.SetValue( "" )
			self.hitting_message.SetLabel( "" )

		else:
			#一度「〜.SetValue( "" )」で空白化してから書き込まないと、表記が追加されていく
			self.find_result.SetValue( "" )
			self.find_result.write( hit_lines )
			self.hitting_message.SetLabel( str( hit ) + "件が該当しました" )


#行か列を追加・削減する選択ダイアログ
class RowColPlusMinusDialog( wx.Dialog ):

	def __init__( self ):
		wx.Dialog.__init__( self, None, -1, "行ないし列を追加・削減", size = ( 700, 440 ) )

		current_row_position = 0

		row_col_array = ( "最後尾に行を追加", "最後から行を削除", "最後尾に列を追加", "最後から列を削除", "途中の行に追加", "途中の行から削除" )
		self.radio_box = wx.RadioBox( self, wx.ID_ANY, "行か列をどうするか", choices = row_col_array, style = wx.RA_VERTICAL )

		self.value_comment = wx.StaticText( self, wx.ID_ANY, "追加・削除する数を入力してOKボタンを押してください" )
		self.plusminus_value = wx.SpinCtrl( self, wx.ID_ANY, min = 1 )

		self.ok_button = wx.Button( self, wx.ID_OK, "OK" )
		self.cancel_button = wx.Button( self, wx.ID_CANCEL, "Cancel" )
		button_sizer = wx.BoxSizer( wx.HORIZONTAL )
		button_sizer.Add( self.ok_button, 0, wx.FIXED_MINSIZE )
		button_sizer.Add( self.cancel_button, 0, wx.FIXED_MINSIZE )

		sizer = wx.BoxSizer( wx.VERTICAL )
		sizer.Add( self.radio_box, 1, wx.ALL | wx.EXPAND, 10 )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, "※途中の列に追加・削除は、住所や名前の順番に混乱を招くので実装していません" ), 0, wx.ALIGN_CENTER_HORIZONTAL )
		sizer.Add( wx.StaticLine( self ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		sizer.Add( self.value_comment, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )
		sizer.Add( self.plusminus_value )
		sizer.Add( button_sizer, 0, wx.BOTTOM | wx.ALIGN_CENTER_HORIZONTAL )
		self.SetSizer( sizer )

	def set_current_row( self, row_position = 0 ):
		current_row_position = row_position
		self.radio_box.SetItemLabel( 4, "選択された行 or 選択中セルのある行（" + str( current_row_position + 1 ) + "行目）に追加" )
		self.radio_box.SetItemLabel( 5, "選択された行 or 選択中セルのある行（" + str( current_row_position + 1 ) + "行目）から削除" )

	def get_row_col_add_del( self ):
		choice = self.radio_box.GetSelection()

		if choice == 0:
			return ( "row", self.plusminus_value.GetValue() )
		elif choice == 1:
			return ( "row", self.plusminus_value.GetValue() * -1 )
		elif choice == 2:
			return ( "col", self.plusminus_value.GetValue() )
		elif choice == 3:
			return ( "col", self.plusminus_value.GetValue() * -1 )
		elif choice == 4:
			return ( "row-insert", self.plusminus_value.GetValue() )
		else:
			return ( "row-cut", self.plusminus_value.GetValue() )


#検索・置換ダイアログ
class SearchReplaceDialog( wx.Dialog ):

	def __init__( self ):
		wx.Dialog.__init__( self, None, -1, "検索または置換", size = ( 600, 420 ) )

		self.input_box_message = wx.StaticText( self, wx.ID_ANY, "検索する単語を入れてください" )
		self.input_search_word = wx.TextCtrl( self, wx.ID_ANY, style=wx.TE_PROCESS_ENTER )

		self.checkbox_replace = wx.CheckBox( self, wx.ID_ANY, "検索ではなく置換をする" )
		self.checkbox_replace.SetValue( False )

		self.replace_box_message = wx.StaticText( self, wx.ID_ANY, "（検索モードなので無効）どんな言葉に置き換えるか" )
		self.input_replace_word = wx.TextCtrl( self, wx.ID_ANY, style=wx.TE_PROCESS_ENTER )
		if self.checkbox_replace.GetValue() is False:
			self.input_replace_word.Disable()

		#チェックボックスのバインド
		self.checkbox_replace.Bind( wx.EVT_CHECKBOX, self.change_search_or_replace )

		#検索と置換の入力欄でEnterを押すと、wx.ID_OKを返しダイアログを閉じるようにバインド
		self.input_search_word.Bind( wx.EVT_TEXT_ENTER, self.ok_and_dialog_close )
		self.input_replace_word.Bind( wx.EVT_TEXT_ENTER, self.ok_and_dialog_close )

		self.ok_button = wx.Button( self, wx.ID_OK, "OK" )
		self.cancel_button = wx.Button( self, wx.ID_CANCEL, "Cancel" )
		button_sizer = wx.BoxSizer( wx.HORIZONTAL )
		button_sizer.Add( self.ok_button, 0, wx.LEFT | wx.RIGHT | wx.FIXED_MINSIZE, 50 )
		button_sizer.Add( self.cancel_button, 0, wx.LEFT | wx.RIGHT | wx.FIXED_MINSIZE, 50 )

		sizer = wx.BoxSizer( wx.VERTICAL )
		sizer.Add( self.input_box_message, 0, wx.ALL | wx.EXPAND, 10 )
		sizer.Add( self.input_search_word, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		sizer.Add( wx.StaticLine( self ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		sizer.Add( self.checkbox_replace, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		sizer.Add( wx.StaticLine( self ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		sizer.Add( self.replace_box_message, 0, wx.ALL | wx.EXPAND, 10 )
		sizer.Add( self.input_replace_word, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 10 )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, "※ 置換では、すべての該当箇所を一度に置換します" ), 0, wx.ALIGN_CENTER_HORIZONTAL )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, "検索とは違い、該当箇所にはジャンプしません" ), 0, wx.ALIGN_CENTER_HORIZONTAL )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, "※ 変換確定済み入力欄でEnterを押すと、OKボタンと同様に閉じます" ), 0, wx.ALIGN_CENTER_HORIZONTAL )
		sizer.Add( wx.StaticLine( self ), 0, wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )
		sizer.Add( button_sizer, 0, wx.BOTTOM | wx.ALIGN_CENTER_HORIZONTAL )
		self.SetSizer( sizer )

	def change_search_or_replace( self, event ):

		if self.checkbox_replace.GetValue() is True:
			self.input_box_message.SetLabel( "置換する単語を入れてください" )
			self.replace_box_message.SetLabel( "どんな言葉に置き換えるか入れてください" )
			self.input_replace_word.Enable()

		else:
			self.input_box_message.SetLabel( "検索する単語を入れてください" )
			self.replace_box_message.SetLabel( "（検索モードなので無効）置換したい言葉" )
			self.input_replace_word.Disable()

	def get_replace_mode( self ):
		return self.checkbox_replace.GetValue()

	def get_search_word( self ):
		return self.input_search_word.GetValue()

	def get_replace_word( self ):
		return self.input_replace_word.GetValue()

	#OKボタンを押した時のようにwx.ID_OKを返してダイアログを閉じる
	def ok_and_dialog_close( self, event ):
		self.EndModal( wx.ID_OK )


#履歴ダイアログ
class HistoryDialog( wx.Dialog ):

	def __init__( self, current_history_position, history_length ):
		wx.Dialog.__init__( self, None, -1, "表・セルを変更した履歴", size = ( 500, 600 ) )

		self.redo_undo_listbox = wx.ListBox( self, wx.ID_ANY, choices = [], style = wx.LB_SINGLE | wx.LB_NEEDED_SB )

		#ダブルクリックしたら、wx.ID_OKを返してダイアログを閉じるようにバインドする
		self.redo_undo_listbox.Bind( wx.EVT_LISTBOX_DCLICK, self.ok_and_dialog_close )

		if history_length > 1:

			#ダイアログ上部へ表示する説明の構築
			if current_history_position == 0:
				message1 = "現在の表は履歴 0 （最新の更新状態）です。"
			elif current_history_position == history_length - 1:
				message1 = "現在の表は履歴 " + str( current_history_position ) + " です。（辿れる履歴の中で最古）"
			else:
				message1 = "現在の表は履歴 " + str( current_history_position ) + " です。 （ 0 が最新、 " + str( history_length - 1 ) + " が最古の履歴）"

			message2 = "どの履歴に移動するか選んでください"

			#リストボックスへ履歴の選択肢を組み立てながら追加する
			for i in range( history_length ):
				if i < current_history_position:
					if i == 0:
						self.redo_undo_listbox.Append( "履歴 0 （最新の状態）に進む（リドゥ）", 0 )
					else:
						self.redo_undo_listbox.Append( "履歴 " + str( i ) + " に進む（リドゥ）", i )
				elif i > current_history_position:
					self.redo_undo_listbox.Append( "履歴 " + str( i ) + " まで戻る（アンドゥ）", i )

		else:
			message1 = "戻ったり進んだりするだけの履歴の蓄積がありません。"
			message2 = "ですので現在、このダイアログは無効です。"
			self.redo_undo_listbox.Disable()

		#OKボタンとCancelボタンの行の構築
		self.ok_button = wx.Button( self, wx.ID_OK, "OK" )
		self.cancel_button = wx.Button( self, wx.ID_CANCEL, "Cancel" )
		button_sizer = wx.BoxSizer( wx.HORIZONTAL )
		button_sizer.Add( self.ok_button, 0, wx.LEFT | wx.RIGHT | wx.FIXED_MINSIZE, 50 )
		button_sizer.Add( self.cancel_button, 0, wx.LEFT | wx.RIGHT | wx.FIXED_MINSIZE, 50 )

		#ダイアログに配置していく
		sizer = wx.BoxSizer( wx.VERTICAL )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, message1 ), 0, wx.ALIGN_CENTER_HORIZONTAL )
		sizer.Add( wx.StaticText( self, wx.ID_ANY, message2 ), 0, wx.ALIGN_CENTER_HORIZONTAL )
		sizer.Add( self.redo_undo_listbox, 1, wx.ALL | wx.EXPAND, 10 )
		sizer.Add( button_sizer, 0, wx.BOTTOM | wx.ALIGN_CENTER_HORIZONTAL )
		self.SetSizer( sizer )

	#OKでダイアログを閉じた後で、リストの選択位置を取得するための関数
	def get_list_value( self ):
		return self.redo_undo_listbox.GetClientData( self.redo_undo_listbox.GetSelection() )

	#OKボタンを押した時のようにwx.ID_OKを返してダイアログを閉じる
	def ok_and_dialog_close( self, event ):
		self.EndModal( wx.ID_OK )


#用紙変更時に住所氏名などの位置を自動調整するか尋ね、糊付け部分に相当する高さを指定するダイアログ
class AutoRelocationAndPartsDownDialog( wx.Dialog ):

	def __init__( self, parent, paper_category = "はがき", paper_height = 148 ):
		wx.Dialog.__init__( self, parent, -1, "自動調整の可否", size = ( 800, 300 ) )

		glue_height = int( paper_height * 0.108 )

		auto_relocation_message = "用紙サイズを変更しますが、郵便番号や住所氏名の配置と収納範囲をおおまかに自動調整しておきますか？\n\nNoを選ぶと前の用紙の配置のまま用紙の大きさだけが変わります。\nそうすると、宛先氏名が中央にないなど手動調整の手間が増えると予想されます。"
		self.message = wx.StaticText( self, wx.ID_ANY, auto_relocation_message )

		self.parts_down_mm = wx.SpinCtrl( self, wx.ID_ANY, value = "0" , min = 0, max = 100 )

		#下端にOK,Cancelボタンを移動させるために挟み込むスペース
		self.spacer_panel = wx.Panel( self, wx.ID_ANY )

		self.cancel_button = wx.Button( self, wx.ID_CANCEL, "No" )
		self.ok_button = wx.Button( self, wx.ID_OK, "Yes" )
		self.ok_button.SetDefault()

		twin_buttons_sizer = wx.BoxSizer( wx.HORIZONTAL )
		twin_buttons_sizer.Add( self.cancel_button, 0, wx.LEFT | wx.RIGHT | wx.FIXED_MINSIZE, 0 )
		twin_buttons_sizer.Add( self.ok_button, 0, wx.LEFT | wx.RIGHT | wx.FIXED_MINSIZE, 0 )

		sizer = wx.BoxSizer( wx.VERTICAL )
		sizer.Add( self.message, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 10 )

		if "はがき" in paper_category:
			self.parts_down_mm.Hide()

		else:
			#一行にまとめる
			self.parts_down_sizer = wx.BoxSizer( wx.HORIZONTAL )
			self.parts_down_sizer.Add( wx.StaticText( self, wx.ID_ANY, "※ 封筒のフタ部分に相当する" ) )
			self.parts_down_sizer.Add( self.parts_down_mm )
			self.parts_down_sizer.Add( wx.StaticText( self, wx.ID_ANY, "mm の長さ、用紙を縦に拡大して全パーツを下げる" ) )
			self.default_honorific_sizer2 = wx.BoxSizer( wx.HORIZONTAL )

			self.parts_down_mm.SetValue( glue_height )
			sizer.Add( wx.StaticLine( self, style = wx.LI_HORIZONTAL ), 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 10 )
			sizer.Add( self.parts_down_sizer, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM | wx.EXPAND, 10 )

		sizer.Add( self.spacer_panel, 1 )
		sizer.Add( twin_buttons_sizer, 0, wx.BOTTOM | wx.ALL | wx.EXPAND, 4 )
		self.SetSizer(sizer)

		self.ok_button.SetFocus()


	def get_parts_down_mm( self ):
		return self.parts_down_mm.GetValue()


# csv_to_list() は csv_utils.py に移動しました
# エントリーポイント（if __name__ == "__main__":）は main.py に移動しました
