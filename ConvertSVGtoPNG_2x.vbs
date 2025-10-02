Option Explicit

'==========================================================================
' SVG to High-Quality PNG Converter (Advanced Sizing) - v5.2
'
' ■機能
'   SVGファイルを、柔軟なサイズ指定で高画質なPNGファイルに変換します。
'
' ■変更点 (v5.2)
'   - ピクセル指定時 (-resize) の品質を大幅に向上。
'     SVGを一度、高解像度 (600 DPI) で読み込んでからリサイズすることで、
'     ぼやけの発生を防ぎ、常にシャープな画像を生成します。
'
' ■サイズ指定の形式
'   - 倍率指定: x1.5 (1.5倍), x2 (2倍)
'   - % 指定:   150 (150%), 200 (200%)
'   - 幅指定:   w200 (幅200px)
'   - 高さ指定: h200 (高さ200px)
'   - 絶対指定: 200x100 (幅200px, 高さ100px)
'
' ■事前準備
'   「ImageMagick」のインストールが必要です。
'
'==========================================================================

Dim objShell, objFSO, objArgs
Dim strSVGPath, strPNGPath, strCommand, strResizeParam, fileCount, input, msg

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

' --- サイズ指定の入力を受け付ける ---
msg = "出力PNGのサイズを指定してください。" & vbCrLf & vbCrLf & _
      "  x1.5  または  x2    (1.5倍、2倍に拡大)" & vbCrLf & _
      "  150              (150%に拡大)" & vbCrLf & _
      "  w200             (幅を200pxに)" & vbCrLf & _
      "  h200             (高さを200pxに)" & vbCrLf & _
      "  200x100          (幅200, 高さ100pxに)"
      
input = InputBox(msg, "サイズ指定", "x2")

If IsEmpty(input) Or input = "" Then WScript.Quit ' キャンセル

input = LCase(Trim(input)) ' 小文字に変換し、前後のスペースを削除
strResizeParam = ""

' --- 入力値を解析してリサイズパラメータを決定 ---
Dim density ' density は複数の条件で使うためここで宣言

If Left(input, 1) = "x" And IsNumeric(Mid(input, 2)) Then
    ' --- 1. 倍率指定 (x1.5, x2) ---
    Dim multiplier
    multiplier = CDbl(Mid(input, 2))
    If multiplier <= 0 Then
        MsgBox "エラー: 0より大きい倍率を指定してください。", vbCritical, "入力エラー"
        WScript.Quit
    End If
    density = 96 * multiplier
    strResizeParam = "-density " & density

ElseIf Left(input, 1) = "w" And IsNumeric(Mid(input, 2)) Then
    ' --- 2. 幅指定 (高品質化) ---
    strResizeParam = "-density 600 -resize " & Mid(input, 2) & "x"

ElseIf Left(input, 1) = "h" And IsNumeric(Mid(input, 2)) Then
    ' --- 3. 高さ指定 (高品質化) ---
    strResizeParam = "-density 600 -resize x" & Mid(input, 2)

ElseIf InStr(input, "x") > 1 Then ' "x"を含み、かつ先頭ではない
    ' --- 4. 幅x高さ指定 (高品質化) ---
    Dim parts, width, height
    parts = Split(input, "x")
    If UBound(parts) = 1 And IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
        width = parts(0)
        height = parts(1)
        strResizeParam = "-density 600 -resize " & width & "x" & height & "!" ' !でアスペクト比を無視
    End If

ElseIf IsNumeric(input) Then
    ' --- 5. パーセント指定 (150, 200) ---
    Dim scalePercent
    scalePercent = CDbl(input)
    If scalePercent <= 0 Then
        MsgBox "エラー: 0より大きい数値を入力してください。", vbCritical, "入力エラー"
        WScript.Quit
    End If
    density = 96 * (scalePercent / 100)
    strResizeParam = "-density " & density

End If

' パラメータが決定できなかった場合はエラー
If strResizeParam = "" Then
    MsgBox "エラー: 入力された形式が正しくありません。" & vbCrLf & "指定された形式で入力してください。", vbCritical, "入力エラー"
    WScript.Quit
End If

' --- 各ファイルをループ処理 ---
For Each strSVGPath In objArgs
    If objFSO.FileExists(strSVGPath) And LCase(objFSO.GetExtensionName(strSVGPath)) = "svg" Then
        strPNGPath = objFSO.BuildPath(objFSO.GetParentFolderName(strSVGPath), objFSO.GetBaseName(strSVGPath) & ".png")
        
        ' ImageMagickのコマンドを組み立てる
        strCommand = "magick convert -background none " & strResizeParam & " """ & strSVGPath & """ """ & strPNGPath & """"

        ' コマンドを実行
        On Error Resume Next
        objShell.Run strCommand, 0, True
        If Err.Number <> 0 Then
            MsgBox "ImageMagickの実行に失敗しました。" & vbCrLf & _
                   "ImageMagickが正しくインストールされているか確認してください。", vbCritical, "実行エラー"
            Err.Clear
            WScript.Quit
        End If
        On Error GoTo 0
        
        fileCount = fileCount + 1
    End If
Next

' --- 完了メッセージ ---
If fileCount > 0 Then
    MsgBox fileCount & "個のSVGファイルを指定のサイズでPNGに変換しました。", vbInformation, "処理完了"
Else
    MsgBox "処理対象のSVGファイルが見つかりませんでした。", vbExclamation, "処理完了"
End If

' --- オブジェクトの解放 ---
Set objShell = Nothing
Set objFSO = Nothing
Set objArgs = Nothing
