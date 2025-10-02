Option Explicit

'==========================================================================
' SVG to High-Quality PNG Converter (2x Size) - v2
'
' ■機能
'   SVGファイルを、高画質を保ったまま元のサイズの縦横2倍のPNGファイルに
'   変換します。背景は透過になります。
'
' ■変更点 (v2)
'   画像がぼやける問題に対応するため、変換処理の方法を変更しました。
'   SVGを読み込む時点の解像度(DPI)を標準の2倍(192 DPI)に指定することで、
'   ベクター情報のディテールを失うことなく高精細なPNGを生成します。
'
' ■使い方
'   1. このスクリプトファイル(.vbs)をデスクトップなどに保存します。
'   2. 変換したいSVGファイルを、このスクリプトのアイコンの上に
'      ドラッグ＆ドロップします。（複数ファイルも可）
'   3. SVGファイルと同じフォルダに、倍サイズのPNGファイルが作成されます。
'
' ■事前準備
'   このスクリプトを実行するには、「ImageMagick」がインストールされている
'   必要があります。
'   公式サイト: https://imagemagick.org/
'
'==========================================================================

Dim objShell, objFSO, objArgs
Dim strSVGPath, strPNGPath, strCommand, fileCount

' --- オブジェクトの生成 ---
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
fileCount = 0

' --- ドラッグ＆ドロップされたかチェック ---
If objArgs.Count = 0 Then
    MsgBox "SVGファイルをこのスクリプトのアイコンにドラッグ＆ドロップしてください。", vbInformation, "使い方"
    WScript.Quit
End If

' --- 各ファイルをループ処理 ---
For Each strSVGPath In objArgs

    ' ファイルが存在し、拡張子がSVGか確認
    If objFSO.FileExists(strSVGPath) And LCase(objFSO.GetExtensionName(strSVGPath)) = "svg" Then

        ' 出力するPNGファイルのパスを生成 (例: icon.svg -> icon.png)
        strPNGPath = objFSO.BuildPath( _
            objFSO.GetParentFolderName(strSVGPath), _
            objFSO.GetBaseName(strSVGPath) & ".png" _
        )

        ' ImageMagickを実行するためのコマンドを組み立てる
        ' 【重要】-density 192: SVGを読み込む際の解像度を標準(96dpi)の2倍に指定します。
        '                      これにより、ラスタライズ(ビットマップ化)時点で2倍のピクセルサイズになります。
        '                      このオプションは入力ファイルより「前」に記述する必要があります。
        ' -background none: 背景を透過に設定
        strCommand = "magick convert -density 192 -background none """ & strSVGPath & """ """ & strPNGPath & """"

        ' コマンドを実行
        ' 第2引数 0: コマンドプロンプトのウィンドウを非表示にする
        ' 第3引数 True: コマンドの処理が完了するまで待機する
        On Error Resume Next
        objShell.Run strCommand, 0, True
        If Err.Number <> 0 Then
            MsgBox "ImageMagickの実行に失敗しました。" & vbCrLf & _
                   "ImageMagickが正しくインストールされ、PATHが通っているか確認してください。" & vbCrLf & vbCrLf & _
                   "実行コマンド: " & strCommand, vbCritical, "実行エラー"
            Err.Clear
            WScript.Quit
        End If
        On Error GoTo 0
        
        fileCount = fileCount + 1

    End If
Next

' --- 完了メッセージ ---
If fileCount > 0 Then
    MsgBox fileCount & "個のSVGファイルを高画質PNGに変換しました。", vbInformation, "処理完了"
Else
    MsgBox "処理対象のSVGファイルが見つかりませんでした。", vbExclamation, "処理完了"
End If

' --- オブジェクトの解放 ---
Set objShell = Nothing
Set objFSO = Nothing
Set objArgs = Nothing