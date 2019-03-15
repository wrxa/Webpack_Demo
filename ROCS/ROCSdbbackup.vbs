rem ***** バックアップ＆結果メール送信 *****

Option Explicit

Const vbHide = 0             'ウィンドウを非表示
Const vbNormalFocus = 1      '通常のウィンドウ、かつ最前面のウィンドウ
Const vbMinimizedFocus = 2   '最小化、かつ最前面のウィンドウ
Const vbMaximizedFocus = 3   '最大化、かつ最前面のウィンドウ
Const vbNormalNoFocus = 4    '通常のウィンドウ、ただし、最前面にはならない
Const vbMinimizedNoFocus = 6 '最小化、ただし、最前面にはならない

Dim cnvMOTO
Dim strYMD
Dim strHMS
Dim strDate

Dim strFolder
Dim strExec

Dim objWShell

Dim objFSO
Dim objFolder
Dim strSize

Dim objMsg

rem 日付文字列作成 yyyy/mm/dd hh:mm:ssから / や : を削除
cnvMOTO = Now()
strYMD = Replace(FormatDateTime(cnvMOTO,2),"/","")
strHMS = Right("0" & Replace(FormatDateTime(cnvMOTO,3),":",""),6)
strDate = strYMD & strHMS

rem 添付ファイルバックアップ
Set objWShell = CreateObject("WScript.Shell")
strFolder = "\\snd5420\remhope1$\backup\db\rocs\" & strDate
strExec = "xcopy /I /Q /S c:\ROCS\Attachment " & strFolder & "\Attachment"
objWShell.Run strExec, vbNormalFocus, True

rem DBバックアップ
strExec = "cmd /c c:\xampp\mysql\bin\mysqldump.exe --user=root --opt --databases rocs > " & strFolder & "\mysql_T.dmp"
objWShell.Run strExec, vbNormalFocus, True

rem 日曜に全バリシートをバックアップ
if WeekDay(cnvMOTO)=1 then
	strExec = "xcopy /I /Q /S \\snd8990\data-rocs\AllVariation_T_Backup " & strFolder & "\AllVariation_T"
	objWShell.Run strExec, vbNormalFocus, True
End if
Set objWShell = Nothing

rem バックアップサイズ取得
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolder)
strSize = objFolder.Size
Set objFolder = Nothing
Set objFSO = Nothing

rem メール送信
Set objMsg = CreateObject("CDO.Message")

objMsg.From = "tbatrocs@hgt-rocs-app04.jpn.mds.honda.com"
objMsg.To = "hgtgrp_ctrl_rocs_admin@n.t.rd.honda.co.jp"
objMsg.Subject = "ROCS DB backup (hgt-rocs-app05)"
objMsg.TextBody = "ROCS DBが、" & strFolder & " にバックアップされました。"
objMsg.TextBody = objMsg.TextBody & vbCrLf & "サイズは " & FormatNumber(strSize/1024/1024,1,0,0,-1) & "MB です。"
objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "webhope1-t.edp.t.rd.honda.co.jp"
objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "fantasia.t.rd.honda.com"
objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMsg.Configuration.Fields.Update
objMsg.Send
Set objMsg = Nothing

WScript.Echo "Script end."
