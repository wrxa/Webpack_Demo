rem ***** �o�b�N�A�b�v�����ʃ��[�����M *****

Option Explicit

Const vbHide = 0             '�E�B���h�E���\��
Const vbNormalFocus = 1      '�ʏ�̃E�B���h�E�A���őO�ʂ̃E�B���h�E
Const vbMinimizedFocus = 2   '�ŏ����A���őO�ʂ̃E�B���h�E
Const vbMaximizedFocus = 3   '�ő剻�A���őO�ʂ̃E�B���h�E
Const vbNormalNoFocus = 4    '�ʏ�̃E�B���h�E�A�������A�őO�ʂɂ͂Ȃ�Ȃ�
Const vbMinimizedNoFocus = 6 '�ŏ����A�������A�őO�ʂɂ͂Ȃ�Ȃ�

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

rem ���t������쐬 yyyy/mm/dd hh:mm:ss���� / �� : ���폜
cnvMOTO = Now()
strYMD = Replace(FormatDateTime(cnvMOTO,2),"/","")
strHMS = Right("0" & Replace(FormatDateTime(cnvMOTO,3),":",""),6)
strDate = strYMD & strHMS

rem �Y�t�t�@�C���o�b�N�A�b�v
Set objWShell = CreateObject("WScript.Shell")
strFolder = "\\snd5420\remhope1$\backup\db\rocs\" & strDate
strExec = "xcopy /I /Q /S c:\ROCS\Attachment " & strFolder & "\Attachment"
objWShell.Run strExec, vbNormalFocus, True

rem DB�o�b�N�A�b�v
strExec = "cmd /c c:\xampp\mysql\bin\mysqldump.exe --user=root --opt --databases rocs > " & strFolder & "\mysql_T.dmp"
objWShell.Run strExec, vbNormalFocus, True

rem ���j�ɑS�o���V�[�g���o�b�N�A�b�v
if WeekDay(cnvMOTO)=1 then
	strExec = "xcopy /I /Q /S \\snd8990\data-rocs\AllVariation_T_Backup " & strFolder & "\AllVariation_T"
	objWShell.Run strExec, vbNormalFocus, True
End if
Set objWShell = Nothing

rem �o�b�N�A�b�v�T�C�Y�擾
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolder)
strSize = objFolder.Size
Set objFolder = Nothing
Set objFSO = Nothing

rem ���[�����M
Set objMsg = CreateObject("CDO.Message")

objMsg.From = "tbatrocs@hgt-rocs-app04.jpn.mds.honda.com"
objMsg.To = "hgtgrp_ctrl_rocs_admin@n.t.rd.honda.co.jp"
objMsg.Subject = "ROCS DB backup (hgt-rocs-app05)"
objMsg.TextBody = "ROCS DB���A" & strFolder & " �Ƀo�b�N�A�b�v����܂����B"
objMsg.TextBody = objMsg.TextBody & vbCrLf & "�T�C�Y�� " & FormatNumber(strSize/1024/1024,1,0,0,-1) & "MB �ł��B"
objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "webhope1-t.edp.t.rd.honda.co.jp"
objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "fantasia.t.rd.honda.com"
objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMsg.Configuration.Fields.Update
objMsg.Send
Set objMsg = Nothing

WScript.Echo "Script end."
