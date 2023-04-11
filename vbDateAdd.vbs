'-----------�X�N���v�g�@GPX�t�@�C���̎��ԃV�t�g�@2023/04/11 v1.1 --------------------------
'�^�C�g���o�[�Ƀo�[�W�����ԍ���\��
'�N�����̓��͉�ʂ̃��b�Z�[�W������L�C��( - �� / �ɏC��)
'�J�V�~�[���Ȃǂ��������鎞�Ԃ�1����GPX�t�@�C����^����ƃG���[�ƂȂ錏�̏C��
'-----------�X�N���v�g�@GPX�t�@�C���̎��ԃV�t�g�@2020/04/19 v1.0 --------------------------
'����

Option Explicit
Const strTitl = "GPX���ԃV�t�g v1.1"
Const intTC = -9  '�W�����ԌW���@�O���j�b�W�W������-9, ���{���[�J�����Ԃ�0

Function getFormatNum(inOrg)
  If inOrg < 10 Then
    getFormatNum = "0" & CStr(inOrg)
  Else
    getFormatNum = CStr(inOrg)
  End If
End Function

Function getFormatHour(strOrg)
  '������1���̏ꍇ��2���ɕ␳
  If InStr(strOrg, ":") = 13 Then
    getFormatHour = Left(strOrg, 11) & "0" & Mid(strOrg, 12, 7)
  Else
    getFormatHour = strOrg
  End If
End Function

Sub checkErr()
  If Err.Number = 0 Then Exit Sub
  MsgBox "�G���[���������܂����B�����𒆎~���܂��B(" & Err.Description & ")", , strTitl
  WScript.Quit  '�I��
End Sub

Sub Main()
  Dim strScriptPath, objFSO, strArg, objDrpFile, strBuf
  Dim intI, intJ

  If WScript.Arguments.Count = 0 Then
    MsgBox "GPX�t�@�C�����h���b�v���Ă��������B", , strTitl
    WScript.Quit
  End If
  
  '�X�N���v�g�̂���t�H���_����ƃt�H���_�ɂ���FileSystemObject�̎擾
  strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  strArg = WScript.Arguments(0)
  If Not objFSO.FileExists(strArg) Then
    MsgBox "�h���b�v���ꂽ�̂̓t�@�C���ł͂���܂���B", , strTitl
    WScript.Quit  '�I��
  End If

  '���̓t�@�C����S�s�ǂݍ��݃o�b�t�@�Ɋi�[
  On Error Resume Next
  Set objDrpFile = CreateObject("ADODB.Stream")
  checkErr()
  With objDrpFile
    .Type = 2
    .Charset = "UTF-8"
    .Open
    .LoadFromFile strArg
    checkErr()
    '���s�R�[�h��Lf(Strava�Ȃ�)�܂���CrLf(�J�V�~�[���Ȃ�)�̂�����ɂ��Ή�
    '���邽�ߑS�s�ǂݍ���
    strBuf = .ReadText(-1)
    checkErr()
    .Close
  End With
  On Error GoTo 0

  '�擪����J�n����<time>��T��
  intJ = 0
  intI = InStr(strBuf, "<trkseg")
  If intI > 0 Then
    intJ = InStr(intI, strBuf, "<time>")  'trkseg�̒����time���J�n����
  End If
  If intJ = 0 Then
    MsgBox "�h���b�v���ꂽ�̂�GPX�t�@�C���ł͂���܂���B", , strTitl
    WScript.Quit  '�I��
  End If

  '�C����̓����̓���UI�\������ ����l�ɂ͏C���O�̒l��\������
  Dim strStrTimInit, longDiffSeconds
  Dim strStrTim, strDate, dateValue, strNewDateTime, strNewBuf

  strStrTimInit = Mid(strBuf, intJ + 6, 19)
  strStrTimInit = getFormatHour(strStrTimInit)
  strStrTimInit = Replace(strStrTimInit, "T", " ")
  strStrTimInit = DateAdd("h", intTC * -1, strStrTimInit)
  strStrTimInit = getFormatHour(strStrTimInit)
  
  strDate = InputBox("�C����̏o�����t[���p10����]����� (yyyy/mm/dd)", strTitl, Mid(strStrTimInit, 1, 10))
  If IsEmpty(strDate) Then
    WScript.Quit
  End If
  strStrTim = InputBox("�C����̏o������[���p8����]����� (hh:mm:ss)", strTitl, Mid(strStrTimInit, 12, 8))
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
    MsgBox "�C���O�ƏC����͈قȂ�������w�肵�Ă��������B", , strTitl
    WScript.Quit  '�I��
  End If

  '�o�̓t�@�C������
  Dim strOutFilePath
  With objFSO
    strOutFilePath = strScriptPath + .GetBaseName(strArg) + "_Change.gpx"
    If .FileExists(strOutFilePath) Then
      If MsgBox("�����̃t�@�C��������܂��B�㏑�����܂����H", vbYesNo, strTitl) = vbNo Then
        WScript.Quit
      End If
      .DeleteFile (strOutFilePath)
    End If
  End With

  '���ԃV�t�g����
  Dim iProcTime
  iProcTime = Timer
  intJ = InStr(strBuf, "<time>")  '�w�b�_������time���ύX���邽�ߐ擪���猟��
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

  '�o�̓t�@�C���쐬
  Dim objOutFile
  Set objOutFile = CreateObject("ADODB.Stream")
  With objOutFile
    .Type = 2
    .Charset = "UTF-8"
    .Open
    .WriteText strNewBuf, 0
  
    '�o�̓t�@�C����BOM����
    .Position = 0   '�擪��Seek
    .Type = 1      '�o�C�i���`���ɕύX
    .Position = 3   '�ʒu��3�o�C�g���ړ�
    strNewBuf = .Read
    .Position = 0
    .Write strNewBuf
    .SetEOS
    .SaveToFile strOutFilePath, 2
    .Close
  End With
  iProcTime = Timer - iProcTime
  MsgBox "�I�����܂����B[��������:" & iProcTime & "�b]", , strTitl
  WScript.Quit
End Sub

Call Main

