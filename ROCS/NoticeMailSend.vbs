'*********************************************************************
'*
'* Class Name : NoticeMailSend.vbs
'* author  wang.rui
'* version 1.00 2019/03/13
'*
'* History
'* 1.00 2019/03/13  FXS)wang.rui			initialize release.
'*********************************************************************
Option Explicit

'*******************************************************************************************************************************************************	
	
	'ImportFile函数用来引入其他文件
	Sub ImportFile(sInstFile)
		Dim oFSO, f, s
		On Error Resume Next
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set f = oFSO.OpenTextFile(sInstFile)
		s = f.ReadAll
		If Err.Number <> 0  Then
			WScript.Echo Err.Number & Err.Description & Err.Source
		End If
		f.Close
		Set oFSO = Nothing
		Set f = Nothing
		ExecuteGlobal s
	End Sub 
	
	'********************************************************************************************************************************	

	'ImportFile("UtilService.vbs")
	ImportFile("MySQLController.vbs")
	Dim utils, sqlController
	'Set utils = New UtilService
	Set sqlController = New MySQLController
	
	sqlController.SelectFromAllVariStatus()

	'********************************************************************************************************************************
	
	
	
