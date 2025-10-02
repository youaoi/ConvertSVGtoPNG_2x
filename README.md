# ConvertSVGtoPNG_2x / SVG to High-Quality PNG Converter (VBScript)

SVGファイルをドラッグ＆ドロップするだけで、高画質を保ったまま**2倍サイズのPNGファイル**に変換できるWindows用VBScriptです。(このscriptはAIによって作成されました)

![demo](https://user-images.githubusercontent.com/12345/your-demo-image.gif)
*（ここにリポジトリにアップロードしたデモGIFなどの画像URLを貼り付けてください）*

## ✨ 特徴 (Features)

* **簡単な操作**: 変換したいSVGファイルをスクリプトアイコンにドラッグ＆ドロップするだけです。
* **高品質な変換**: ImageMagickを利用し、SVGのシャープさを損なわずにPNGへ変換します。
* **複数ファイル対応**: 複数のSVGファイルを一度にまとめて処理できます。
* **背景透過**: 生成されるPNGの背景は自動的に透過になります。
* **インストール不要**: スクリプト自体はWindows標準の機能で動作するため、インストールは不要です。（※別途ImageMagickは必要です）

## ⚙️ 動作要件 (Requirements)

* **OS**: Windows
* **依存ソフトウェア**: **[ImageMagick](https://imagemagick.org/)**
    * このスクリプトは内部でImageMagickのコマンドを呼び出しています。**必ず事前にインストールしてください。**

## 🚀 セットアップ (Setup)

### 1. ImageMagickのインストール

このスクリプトを使用するには、まず画像処理ツール「ImageMagick」をインストールする必要があります。

1.  **[ImageMagickの公式サイト](https://imagemagick.org/script/download.php)**にアクセスし、Windows用のインストーラーをダウンロードします。
2.  インストーラーを起動し、ウィザードに従ってインストールを進めます。
3.  **インストールの途中で表示される以下のオプションに、必ずチェックを入れてください。** これらが有効でないと、スクリプトは正しく動作しません。
    * `Add application directory to your system path (PATH)`
    * `Install legacy utilities (e.g., convert)`

    ![ImageMagick Installer Options](https://user-images.githubusercontent.com/12345/your-installer-screenshot.png)
    *（ここにインストーラー画面のスクリーンショットURLを貼り付けてください）*


### 2. VBScriptファイルのダウンロード

1.  このリポジトリから `ConvertSVGtoPNG_2x.vbs` ファイルをダウンロードします。
2.  ダウンロードしたファイルを、デスクトップなどの使いやすい場所に配置してください。

## 使い方 (Usage)

1.  変換したいSVGファイル（複数可）を選択します。
2.  選択したファイルを、`ConvertSVGtoPNG_2x.vbs` のアイコン上にドラッグ＆ドロップします。
3.  処理が完了するとメッセージが表示され、元のSVGファイルと同じフォルダに2倍サイズのPNGファイルが作成されています。

## 🛠️ 仕組み (How it Works)

このVBScriptは、ドラッグ＆ドロップされたSVGファイルのパスを取得し、それらを引数としてImageMagickのコマンドラインプロセスをバックグラウンドで実行します。

実行されるコマンドの要点は以下の通りです。

```bash
magick convert -density 192 -background none "input.svg" "output.png"
```

`-density 192` オプションがこのスクリプトの肝です。これは、SVG（ベクター形式）をPNG（ビットマップ形式）に変換（ラスタライズ）する際の解像度を、標準的な96dpiの2倍である192dpiに指定するものです。

これにより、ビットマップに変換した後に画像を引き伸ばすのではなく、ベクターデータの情報を基に初めから高精細なビットマップを生成するため、シャープな画質が保たれます。

## 📄 ライセンス (License)

このプロジェクトは [MIT License](LICENSE) の下で公開されています。自由にご利用、改変、再配布が可能です。
