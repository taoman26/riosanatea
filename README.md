# Riosanatea - はがき宛名印刷ソフト

Ubuntu 22.04 / 24.04 対応の宛名印刷ソフトウェアです。

天杉 善哉氏が配布しているスクリプトを改変したもので、オリジナルスクリプトは以下から入手可能です。

[宛名印刷ソフト「Riosanatea」](https://sound.jp/zenzai/python-script/riosanatea.html)


## 特徴

- はがきや封筒への宛名印刷に対応
- CSV形式の住所録から一括印刷可能
- レイアウトの自由なカスタマイズ
- Ubuntu 22.04 と 24.04 の両方で動作

## インストール

### 必要なライブラリのインストール

```bash
pip install -r requirements.txt
```

または個別にインストール:

```bash
pip install wxPython>=4.1.0 Pillow>=9.0.0
```

### システム要件

- Python 3.6 以上
- Ubuntu 22.04 または 24.04
- CUPS印刷システム

## 使い方

### 基本的な起動

```bash
python3 main.py
```

### CSVファイルを指定して起動

```bash
python3 main.py /path/to/addressbook.csv
```

### 実行権限を付与して直接実行

```bash
chmod +x main.py
./main.py
```

## ファイル構成

```
riosanatea/
├── main.py              # メインエントリーポイント
├── Riosanatea.py        # コアロジックとGUI
├── image_utils.py       # 画像処理ユーティリティ
├── csv_utils.py         # CSV処理と印刷機能
├── requirements.txt     # 必要なライブラリ一覧
├── README.md           # このファイル
└── ReadMe-Orig.pdf     # 天杉 善哉氏のオリジナルREADME
```

## モジュール説明

### main.py
アプリケーションのエントリーポイントです。コマンドライン引数を処理し、GUIを起動します。

### Riosanatea.py
- `atena_image_maker`: 宛名画像生成クラス
- `frame_plus`: メインGUIフレーム
- 各種ダイアログクラス

### image_utils.py
画像処理関連のユーティリティ関数:
- `contraction()`: 画像の縮小処理
- `letter_to_pil_image()`: 文字を画像に変換
- `greyscale_autocrop()`: 余白の自動削除
- `pil_through_paste_greyscale()`: 透過貼り付け
- `maybe_list_natsort()`: 自然順ソート

### csv_utils.py
ファイル処理と印刷関連の機能:
- `csv_to_list()`: CSV読み込み（複数文字コード対応）
- `pil_printing()`: 画像の印刷処理（Ubuntu 22.04/24.04対応）

## CSV形式

住所録は以下のような形式のCSVファイルで管理します:

```csv
郵便番号,住所1,住所2,氏名1,氏名2,敬称
1234567,東京都千代田区霞が関1-1-1,,山田 太郎,,様
```

列の順序は「差出人記述・項目と列の対応」タブで変更できます。

## 印刷設定

1. 「印刷レイアウト」タブで各項目の位置を調整
2. 「差出人記述・敬称等」タブで差出人情報を入力
3. 「住所表編集、宛名印刷」タブで印刷範囲を指定して印刷

## トラブルシューティング

### フォントが見つからない

システムにインストールされているフォントが自動検出されます。
見つからない場合は、フォント選択ダイアログが表示されます。

### 印刷できない

CUPSが正しく設定されているか確認してください:

```bash
lpstat -p -d
```

### Ubuntu 24.04で印刷エラー

このバージョンは Ubuntu 24.04 の変更に対応済みです。
PNG形式で印刷を行うため、両バージョンで動作します。

## ライセンス

BSD 2-Clause License

原著作権:
Copyright 2017 天杉 善哉 (あますぎ ぜんざい Amasugi Zenzai) All rights reserved.

改変:
Copyright 2025 Bureau Mikami

本ソフトウェアは天杉 善哉氏の原著作物をBureau Mikamiが改変したものです。
改変部分についても同じBSD 2-Clause Licenseが適用されます。

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

## バージョン

Version 36 (Ubuntu 22.04/24.04 対応版)

## 開発者向け情報

### コードの分割について

巨大な単一ファイルから、以下のように機能別にモジュール化しました:

- 画像処理 → `image_utils.py`
- ファイル処理・印刷 → `csv_utils.py`
- メインロジック → `Riosanatea.py`
- エントリーポイント → `main.py`

### さらなる改善案

- ダイアログクラスを `gui_dialogs.py` に分離
- カスタムウィジェットを `gui_widgets.py` に分離
- 設定管理を `config.py` に分離
- テストコードの追加

## サポート

問題が発生した場合は、Issueを作成してください。
