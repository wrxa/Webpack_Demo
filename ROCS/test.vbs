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
	'Include函数用来引入其他文件
	Function  Include(sInstFile)
		Dim oFSO, f, s
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set f = oFSO.OpenTextFile(sInstFile)
		s = f.ReadAll
		f.Close
		Set oFSO = Nothing
		Set f = Nothing
		ExecuteGlobal s
	End Function 
	
	'引入MySQLController.class文件，注意文件路径
	Include("MySQLController.class")
	Dim sqlController, bConnRet, bSelectRet
	Set sqlController = New MySQLController
	bConnRet = sqlController.ConnectDB()
	
	If bConnRet = False Then
		WScript.Echo "connect MySQL Fail"
	Else
		WScript.Echo "connect MySQL Success"
		'bSelectRet = sqlController.SelectFromDB()
		'If bSelectRet = False Then
			'WScript.Echo "Select from MySQL Fail"
		'Else
			'WScript.Echo "Select from MySQL Success"
		'End If	
	End If

	sqlController.CloseDB()
	'********************************************************************************************************************************
	Include("CsvController.class")
	Dim m_NasFolder, m_SuccessFolder, m_FailFolder
	m_NasFolder = "E:\CHANYE\ROCS\NAS"
	m_SuccessFolder = "E:\CHANYE\ROCS\SUCCESS_CSV"
	m_FailFolder = "E:\CHANYE\ROCS\FAIL_CSV"
	
	Dim csvCon
	Set csvCon = New CsvController
	csvCon.TraverseNasFolder(m_NasFolder)
	
	
