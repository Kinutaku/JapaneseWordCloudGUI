# JapaneseWordCloudGUI

日本語テキストを形態素解析し、WordCloud・共起ネットワーク・頻度グラフを作成できるデスクトップGUIアプリです。MeCab + ipadic を用いた分かち書きに加え、ストップワード編集や連語結合（2〜4語のルールで単語をまとめる）が行えます。

## 主な機能
- テキスト入力（直接入力/ファイル読込/CSV列結合/サンプル）とストップワード管理
- 分かち書き結果をタブ2で確認し、連語結合ルールを追加・プレビュー・適用
- 単語頻度リストの表示・検索・削除・品詞で一括削除
- WordCloud生成（矩形/楕円/カスタム画像マスク、行ごと重複排除オプション）
- 共起ネットワーク生成（スライディング窓 or 行単位、連続重複の圧縮、行内重複ペア抑制、凡例表示）
- 頻度グラフおよび共起頻度表の表示・CSV/画像出力
- 画像保存（PNG/SVG）とCSVエクスポート

## 動作環境
- Python 3.10+ を想定
- 必要ライブラリは `requirements.txt` を参照（主要: tkinter, pillow, wordcloud, mecab-python3, networkx, matplotlib, japanize-matplotlib）
- MeCab は基本的に `pip install mecab-python3 ipadic` で利用可能（バイナリ配布込み）。環境によってはシステム版 MeCab を入れる必要があるため、うまく動かない場合は OS のパッケージ（Windows 配布/`brew install mecab mecab-ipadic` など）も検討してください。
- フォント: Windows の Meiryo を想定。別環境は `C:\Windows\Fonts\meiryo.ttc` を適宜変更

## セットアップ
```bash
pip install -r requirements.txt
```
MeCab が未導入の場合は OS に応じてインストールしてください。

## 使い方
1. `python main.py` を実行すると GUI が起動します。
2. タブ1でテキストを入力または読み込み、「分かち書き実行」で単語抽出します。
3. タブ2で連語結合ルールを追加し、プレビューまたは「適用して編集領域を更新」で反映します。
4. タブ3で単語編集や検索・削除を行い、右側のサブタブから WordCloud/共起ネットワーク/頻度グラフを生成します。
5. 生成結果はボタンから画像またはCSVとして保存できます。

## トラブルシュート
- MeCab が見つからない: MeCab 本体と辞書をインストールし、環境変数 PATH を確認してください。
- フォントが見つからない: `main.py` 冒頭の `font_path` を環境にある日本語フォントに変更してください。
- カスタム画像マスク: 透過PNGなどを推奨し、生成サイズに合わせてリサイズされます。
