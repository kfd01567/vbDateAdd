'-----------スクリプト　GPXファイルの時間シフト　2023/04/11 v1.1 --------------------------
'タイトルバーにバージョン番号を表示
'年月日の入力画面のメッセージ文字誤記修正( - を / に修正)
'カシミールなどが生成する時間が1桁のGPXファイルを与えるとエラーとなる件の修正
'-----------スクリプト　GPXファイルの時間シフト　2020/04/19 v1.0 --------------------------
'初版

Option Explicit
Const strTitl = "GPX時間シフト v1.1"
Const intTC = -9  '標準時間係数　グリニッジ標準時は-9, 日本ローカル時間は0

Function getFormatNum(inOrg)
  If inOrg < 10 Then
    getFormatNum = "0" & CStr(inOrg)
  Else
    getFormatNum = CStr(inOrg)
  End If
End Function

Function getFormatHour(strOrg)
  '時刻が1桁の場合は2桁に補正
  If InStr(strOrg, ":") = 13 Then
    getFormatHour = Left(strOrg, 11) & "0" & Mid(strOrg, 12, 7)
  Else
    getFormatHour = strOrg
  End If
End Function

Sub checkErr()
  If Err.Number = 0 Then Exit Sub
  MsgBox "エラーが発生しました。処理を中止します。(" & Err.Description & ")", , strTitl
  WScript.Quit  '終了
End Sub

Sub Main()
  Dim strScriptPath, objFSO, strArg, objDrpFile, strBuf
  Dim intI, intJ

  If WScript.Arguments.Count = 0 Then
    MsgBox "GPXファイルをドロップしてください。", , strTitl
    WScript.Quit
  End If
  
  'スクリプトのあるフォルダを作業フォルダにしてFileSystemObjectの取得
  strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  strArg = WScript.Arguments(0)
  If Not objFSO.FileExists(strArg) Then
    MsgBox "ドロップされたのはファイルではありません。", , strTitl
    WScript.Quit  '終了
  End If

  '入力ファイルを全行読み込みバッファに格納
  On Error Resume Next
  Set objDrpFile = CreateObject("ADODB.Stream")
  checkErr()
  With objDrpFile
    .Type = 2
    .Charset = "UTF-8"
    .Open
    .LoadFromFile strArg
    checkErr()
    '改行コードがLf(Stravaなど)またはCrLf(カシミールなど)のいずれにも対応
    'するため全行読み込む
    strBuf = .ReadText(-1)
    checkErr()
    .Close
  End With
  On Error GoTo 0

  '先頭から開始日時<time>を探す
  intJ = 0
  intI = InStr(strBuf, "<trkseg")
  If intI > 0 Then
    intJ = InStr(intI, strBuf, "<time>")  'trksegの直後のtimeが開始日時
  End If
  If intJ = 0 Then
    MsgBox "ドロップされたのはGPXファイルではありません。", , strTitl
    WScript.Quit  '終了
  End If

  '修正後の日時の入力UI表示する 既定値には修正前の値を表示する
  Dim strStrTimInit, longDiffSeconds
  Dim strStrTim, strDate, dateValue, strNewDateTime, strNewBuf

  strStrTimInit = Mid(strBuf, intJ + 6, 19)
  strStrTimInit = getFormatHour(strStrTimInit)
  strStrTimInit = Replace(strStrTimInit, "T", " ")
  strStrTimInit = DateAdd("h", intTC * -1, strStrTimInit)
  strStrTimInit = getFormatHour(strStrTimInit)
  
  strDate = InputBox("修正後の出発日付[半角10文字]を入力 (yyyy/mm/dd)", strTitl, Mid(strStrTimInit, 1, 10))
  If IsEmpty(strDate) Then
    WScript.Quit
  End If
  strStrTim = InputBox("修正後の出発時間[半角8文字]を入力 (hh:mm:ss)", strTitl, Mid(strStrTimInit, 12, 8))
  If IsEmpty(strStrTim) Then
    WScript.Quit
  End If
  strStrTim = strDate & " " & strStrTim
  If InStr(strStrTim, ":") = 0 Then strStrTim = strStrTim & " 00:00:00"
  On Error Resume Next
  longDiffSeconds = DateDiff("s", strStrTimInit, strStrTim)
  checkErr()
  On Error GoTo 0
  If longDiffSeconds = 0 Then
    MsgBox "修正前と修正後は異なる日時を指定してください。", , strTitl
    WScript.Quit  '終了
  End If

  '出力ファイル準備
  Dim strOutFilePath
  With objFSO
    strOutFilePath = strScriptPath + .GetBaseName(strArg) + "_Change.gpx"
    If .FileExists(strOutFilePath) Then
      If MsgBox("同名のファイルがあります。上書きしますか？", vbYesNo, strTitl) = vbNo Then
        WScript.Quit
      End If
      .DeleteFile (strOutFilePath)
    End If
  End With

  '時間シフト処理
  Dim iProcTime
  iProcTime = Timer
  intJ = InStr(strBuf, "<time>")  'ヘッダ部分のtimeも変更するため先頭から検索
  strNewBuf = Left(strBuf, intJ + 5)
  Do While True
    strStrTim = Mid(strBuf, intJ + 6, 19)
    strStrTim = getFormatHour(strStrTim)
    strStrTim = Replace(strStrTim, "T", " ")
    dateValue = DateAdd("s", longDiffSeconds, strStrTim)
    dateValue = getFormatHour(dateValue)
    strNewDateTime = DatePart("yyyy", dateValue) & "-" & getFormatNum(DatePart("m", dateValue)) & "-" & getFormatNum(DatePart("d", dateValue)) & "T" & getFormatNum(DatePart("h", dateValue)) & ":" & getFormatNum(DatePart("n", dateValue)) & ":" & getFormatNum(DatePart("s", dateValue)) & "Z"
    strNewBuf = strNewBuf & strNewDateTime
    intI = InStr(intJ + 26, strBuf, "<time>")
    If intI = 0 Then
      Exit Do
    End If
    strNewBuf = strNewBuf & Mid(strBuf, intJ + 26, intI - (intJ + 20))
    intJ = intI
  Loop
  strNewBuf = strNewBuf & Mid(strBuf, intJ + 26)

  '出力ファイル作成
  Dim objOutFile
  Set objOutFile = CreateObject("ADODB.Stream")
  With objOutFile
    .Type = 2
    .Charset = "UTF-8"
    .Open
    .WriteText strNewBuf, 0
  
    '出力ファイルのBOM除去
    .Position = 0   '先頭にSeek
    .Type = 1      'バイナリ形式に変更
    .Position = 3   '位置を3バイト分移動
    strNewBuf = .Read
    .Position = 0
    .Write strNewBuf
    .SetEOS
    .SaveToFile strOutFilePath, 2
    .Close
  End With
  iProcTime = Timer - iProcTime
  MsgBox "終了しました。[処理時間:" & iProcTime & "秒]", , strTitl
  WScript.Quit
End Sub

Call Main

