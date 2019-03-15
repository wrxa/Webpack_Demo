'*********************************************************************
'*
'* Class Name : test.vbs
'* author  wang.rui
'* version 1.00 2019/02/13
'*
'* History
'* 1.00 2019/02/13  FXS)wang.rui			initialize release.
'*********************************************************************
Option Explicit
'输入并回显你的名字, 使用InputBox和Msgbox函数
'Dim name,msg
'msg="Please input your name:"
'name=Inputbox(msg)
'Msgbox name

'Const strConn="dsn=wrMySQL; driver={MySQL ODBC 5.1 Driver}; server=127.0.0.1; database=mytest; port=3306; uid=root; password=123456"
'Set conn = CreateObject("ADODB.connection")
'conn.Open strConn
'查看是否连接成功，成功状态值为1
'If conn.State = 0 Then
     'msgbox  "连接数据库失败"
	 'WScript.Echo "connect MySQL Fail"
'else
    'msgbox   "连接数据库成功"
	'WScript.Echo "connect MySQL Success"
'End If
'WScript.Echo "Script end."
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

	ImportFile("UtilService.vbs")
	ImportFile("CsvController.vbs")
	Dim utils, csvCon
	Set utils = New UtilService
	Set csvCon = New CsvController
	
	csvCon.TraverseNasFolder(utils.m_NasFolder)
	'utils.SendErrorEmail()

	'********************************************************************************************************************************
	
	
	
