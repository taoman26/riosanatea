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
pip install wxPython>=4.1.0 Pillow>=9.0.0 reportlab>=3.6.0
```

### システム要件

- Python 3.6 以上
- Ubuntu 22.04 または 24.04
- CUPS印刷システム
- reportlab (PDF生成用)

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
- `pil_printing()`: 画像の印刷処理
  - PDF形式で印刷（正確なサイズ制御）
  - 自動用紙サイズ検出（Canon/Epson/Brother/HP対応）
  - はがき、封筒、A系列など全サイズ対応
  - Ubuntu 22.04/24.04対応

## CSV形式

住所録は以下のような形式のCSVファイルで管理します:

```csv
郵便番号,住所1,住所2,氏名1,氏名2,敬称
1234567,東京都千代田区霞が関1-1-1,,山田 太郎,,様
```

列の順序は「差出人記述・項目と列の対応」タブで変更できます。

## 印刷設定

### 基本設定

1. 「印刷レイアウト」タブで各項目の位置を調整
2. 「差出人記述・敬称等」タブで差出人情報を入力
3. 「住所表編集、宛名印刷」タブで印刷範囲を指定して印刷

### フォントサイズ調整

印刷レイアウトタブの右側に「フォントサイズ」調整スライダーがあります（50%〜150%）。
印刷した文字が大きすぎる/小さすぎる場合は、ここで調整してください。

### 対応プリンタ

以下のメーカーのプリンタで動作確認済み:
- **Canon** (例: TS8230シリーズ)
- **Epson**
- **Brother**
- **HP**

### 対応用紙サイズ

アプリケーション内で選択可能な全ての用紙サイズに対応:
- **はがき**: 100×148mm
- **往復はがき**: 200×148mm
- **封筒**: 長形3号/4号、角形2号/3号/A4号 など
- **標準用紙**: A3/A4/A5/A6

印刷時は画像サイズから自動的に適切なプリンタ設定が選択されます。

## トラブルシューティング

### フォントが見つからない

システムにインストールされているフォントが自動検出されます。
見つからない場合は、フォント選択ダイアログが表示されます。

### 印刷できない

CUPSが正しく設定されているか確認してください:

```bash
lpstat -p -d
```

デフォルトプリンタが設定されていない場合は、以下のコマンドで設定:

```bash
lpoptions -d プリンタ名
```

### 印刷サイズが合わない

**フォントサイズ調整を使用:**
印刷レイアウトタブの右側にある「フォントサイズ」スライダーで調整してください（50%〜150%）。

**プリンタ設定の確認:**
プリンタがサポートする用紙サイズを確認:

```bash
lpoptions -l | grep -i "pagesize\|media"
```

**PDF印刷方式:**
このバージョンはPDF形式で印刷するため、画像サイズと印刷サイズが正確に一致します。
スケーリングの問題が発生した場合は、プリンタドライバの設定を確認してください。

### Ubuntu 24.04で印刷エラー

このバージョンは Ubuntu 24.04 の変更に対応済みです。
PDF形式で印刷を行うため、両バージョンで動作します。

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

Version 36-20251202 (Ubuntu 22.04/24.04 対応版)

### 主な変更点

**印刷機能の改善:**
- PDF形式での印刷に変更（正確なサイズ制御）
- 全メーカー（Canon/Epson/Brother/HP）のプリンタに対応
- 自動用紙サイズ検出機能
- はがき、封筒、A系列など全サイズ対応

**UI機能追加:**
- フォントサイズ調整機能（50%〜150%）
- SpinButtonのサイズ警告を修正
- ボタンレイアウトの改善

**コード品質向上:**
- .gitignoreファイル追加
- プリンタ検出とエラーハンドリング改善
- 用紙サイズ定義の一元管理

## 開発者向け情報

### コードについて

単一ファイルから、以下のように機能別にモジュール化しました。コードの分割および機能修正はGitHub Copilot Agentを使用して実行し、コードレビューのみ人間が行っています。

- 画像処理 → `image_utils.py`
- ファイル処理・印刷 → `csv_utils.py`
- メインロジック → `Riosanatea.py`
- エントリーポイント → `main.py`

### さらなる改善案（検討事項）

- ダイアログクラスを `gui_dialogs.py` に分離
- カスタムウィジェットを `gui_widgets.py` に分離
- 設定管理を `config.py` に分離
- テストコードの追加

## サポート

問題が発生した場合は、Issueを作成してください。ただし、対応をお約束するものではありません。
