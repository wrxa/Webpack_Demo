'*********************************************************************
'*
'* Class Name : MySQLController.class
'* author  wang.rui
'* version 1.00 2019/02/13
'*
'* History
'* 1.00 2019/02/13  FXS)wang.rui			initialize release.
'*********************************************************************
Option Explicit

Class MySQLController
	Dim m_MySQLConn
	Dim m_DSN, m_DataBaseServer, m_DataBaseName, m_DataBasePort, m_DataBaseUID, m_DataBasePassword, m_DataBaseDriver

	Sub Class_Initialize
		m_DSN = "ROCS"
        m_DataBaseDriver = "{MySQL ODBC 5.1 Driver}"
        m_DataBaseServer = "127.0.0.1"
		m_DataBaseName = "fxs_test"
        m_DataBasePort = "3306"
        m_DataBaseUID = "root"
        m_DataBasePassword = ""
	End Sub
	
	'连接数据库
	Public Function ConnectDB()
		Dim strConn
		On Error Resume Next
		'strConn = "DSN=ROCS; DRIVER={MySQL ODBC 5.1 Driver}; SERVER=127.0.0.1; DATABASE=fxs_test; PORT=3306; UID=root; PASSWORD="
		strConn = "DSN=" & m_DSN & ";"
		strConn = strConn & "DRIVER=" & m_DataBaseDriver & ";"
		strConn = strConn & "SERVER=" & m_DataBaseServer & ";"
		strConn = strConn & "DATABASE=" & m_DataBaseName & ";"
		strConn = strConn & "PORT=" & m_DataBasePort & ";"
		strConn = strConn & "UID=" & m_DataBaseUID & ";"
		strConn = strConn & "PASSWORD=" & m_DataBasePassword & ";"
		
		Set m_MySQLConn = CreateObject("ADODB.connection")
		m_MySQLConn.Open strConn
		If Err.Number <> 0 Then
			WScript.Echo Err.Number & Err.Description & Err.Source
			ConnectDB = False
		Else
			'查看是否连接成功，成功状态值为1
			If m_MySQLConn.State = 0 Then
				ConnectDB = False
				'msgbox  "连接数据库失败"
				'WScript.Echo "connect MySQL Fail"
			Else
				ConnectDB = True
				'msgbox   "连接数据库成功"
				'WScript.Echo "connect MySQL Success"
			End If
		End If
	End Function
	
	'释放数据库对象 
	Public Function CloseDB()
		If IsEmpty(m_MySQLConn) = False Then
			If Not m_MySQLConn Is Nothing Then
				m_MySQLConn.Close
				Set m_MySQLConn = Nothing
			End If
		End If
	End Function
	
	'从数据库中检索
	' Public Function SelectFromDB()
	' 	Dim results
	' 	Const selectSQL = "SELECT * FROM table1"
	' 	On Error Resume Next
	' 	If IsEmpty(m_MySQLConn) = False And Not m_MySQLConn Is Nothing Then
	' 		set results = m_MySQLConn.Execute(selectSQL) 
	' 		If Err.Number <> 0  Then
	' 			SelectFromDB = False
	' 		Else
	' 			SelectFromDB = True
	' 			While not results.eof
	' 				WScript.Echo results.Fields.Item("name").Value
	' 				results.moveNext 
	' 			WEND
	' 		End If
	' 	End If	
	' End Function
	
	'插入：全バリステータス ALL_VARI_STATUS 表
	Public Function InsertAllVariStatus(insertSql)
		'Const insertSQL = " INSERT INTO all_vari_status (CAR_DATA_ID, STATUS_ENGINE_INPUT_1_UPDATE_DATE) VALUES (888888, '2019/2/27  16:02:00')"

		If IsEmpty(m_MySQLConn) = False And Not m_MySQLConn Is Nothing Then
			m_MySQLConn.BeginTrans
			On Error Resume Next
			m_MySQLConn.Execute(insertSQL) 
			'InsertAllVariStatus = True

			If Err.Number <> 0  Then
				WScript.Echo Err.Number & Err.Description & Err.Source
				m_MySQLConn.RollbackTrans
				InsertAllVariStatus = False
			Else
				m_MySQLConn.CommitTrans
				InsertAllVariStatus = True
			End If
		Else
			InsertAllVariStatus = False
		End If	
	End Function

	'插入：全バリ入力値履歴 ALL_VARI_INPUT_HISTS 表
	Public Function InsertAllVariInputHists(insertSql)
		Dim results
		On Error Resume Next
		If IsEmpty(m_MySQLConn) = False And Not m_MySQLConn Is Nothing Then
			'set results = m_MySQLConn.Execute(insertSQL) 
			If Err.Number <> 0  Then
				WScript.Echo Err.Number & Err.Description & Err.Source
				InsertAllVariInputHists = False
			Else
				InsertAllVariInputHists = True
			End If
		End If	
	End Function

	Public Function UpdateAllVariStatus(updateSqlList)

		ConnectDB()

		If IsEmpty(m_MySQLConn) = False And Not m_MySQLConn Is Nothing Then
			Dim i, Dim flag
			flag = False

			m_MySQLConn.BeginTrans
			For i = 0 to ubound(updateSqlList)	
				On Error Resume Next
				m_MySQLConn.Execute(updateSql(i)) 				

				If Err.Number <> 0  Then
					WScript.Echo Err.Number & Err.Description & Err.Source					
					flag = False
					Exit For
				Else
					flag = True
				End If
			Next

			If flag = True Then
				m_MySQLConn.CommitTrans
				UpdateAllVariStatus = True
			Else 
				m_MySQLConn.RollbackTrans
				UpdateAllVariStatus = False
			End If
			
		Else
			UpdateAllVariStatus = False
		End If

		CloseDB()
	End Function

	Public Function SelectFromAllVariStatus()
		Dim Res, selectSQL, Cmd
		selectSQL = "SELECT * FROM all_vari_status"
		Set Res = CreateObject("ADODB.Recordset")
		Set Cmd = CreateObject("ADODB.Command")

		On Error Resume Next
		ConnectDB()

		If IsEmpty(m_MySQLConn) = False And Not m_MySQLConn Is Nothing Then
			Cmd.activeconnection = m_MySQLConn
			Cmd.CommandType = 1
			Cmd.CommandText = selectSQL
			'Set Res = Cmd.Execute()

			Res.CursorLocation = 3 
        	Res.Open Cmd

			'Res = m_MySQLConn.Execute(selectSQL) 
			If Err.Number <> 0  Then
				SelectFromAllVariStatus = False
			Else
				SelectFromAllVariStatus = True

				' Dim my_date1
				' my_date1 = DateAdd("d",-14,date)
				' my_date1 = year(my_date1) & "-" & Month(my_date1) & "-" & day(my_date1)

				' Dim date2, date3
				' date2 = CDate("2019-03-15") -20
				' date3 = Date()

				Dim currentDate
				currentDate = Date()

				Dim recID, statusMailDate1, statusMailDate2, statusMailDate3, statusMailDate4
				Dim statusWtInputDate, statusCdxaApprovalDate, statusTireApprovalDate, statusBrakeApprovalDate
				Dim statusHubApprovalDate, statusYopApprovalDate, statusWtApprovalDate, statusEngineApprovalDate
				Dim updateStatusSql1, updateStatusSql2, updateStatusSql3, updateStatusSql4, index
				Dim updateSqlList(), emailList()

				
				While Not Res.eof
					index = 0
					recID = Res.Fields.Item("ID").Value
					statusMailDate1 = Res.Fields.Item("STATUS_MAIL_DATE_1").Value
					statusMailDate2 = Res.Fields.Item("STATUS_MAIL_DATE_2").Value
					statusMailDate3 = Res.Fields.Item("STATUS_MAIL_DATE_3").Value
					statusMailDate4 = Res.Fields.Item("STATUS_MAIL_DATE_4").Value
					statusWtInputDate = Res.Fields.Item("STATUS_WT_INPUT_2_DATE").Value
					statusCdxaApprovalDate = Res.Fields.Item("STATUS_CDXA_APPROVAL_2_DATE").Value
					statusTireApprovalDate = Res.Fields.Item("STATUS_TIRE_APPROVAL_2_DATE").Value
					statusBrakeApprovalDate = Res.Fields.Item("STATUS_BRAKE_APPROVAL_2_DATE").Value
					statusHubApprovalDate = Res.Fields.Item("STATUS_HUB_APPROVAL_2_DATE").Value
					statusYopApprovalDate = Res.Fields.Item("STATUS_YOP_APPROVAL_2_DATE").Value
					statusWtApprovalDate = Res.Fields.Item("STATUS_WT_APPROVAL_2_DATE").Value
					statusEngineApprovalDate = Res.Fields.Item("STATUS_ENGINE_APPROVAL_2_DATE").Value

					If statusMailDate1 Is Nothing OR statusMailDate1 = "" Then
						If NOT statusWtInputDate Is Nothing AND statusWtInputDate <> "" Then
							statusWtInputDate = CDate(statusWtInputDate)
							If currentDate = statusWtInputDate -2  Then
								ReDim Preserve updateSqlList(index)
								ReDim Preserve emailList(index)
								updateStatusSql1 = CreateUpdateSql(recID, "STATUS_MAIL_DATE_1", currentDate)
								updateSqlList(index) = updateStatusSql1
								emailList(index) = 1
								' bUpdateRet = UpdateAllVariStatus(updateStatusSql)
								' If bUpdateRet = False Then
								' 	WScript.Echo "Update MySQL Fail"
								' Else
								' 	WScript.Echo "UpdateMySQL Success"
								' End If	
								index = index + 1
							End If
						End If
					End If

					If statusMailDate2 Is Nothing OR statusMailDate2 = "" Then
						If NOT statusCdxaApprovalDate Is Nothing AND statusCdxaApprovalDate <> "" _
							AND NOT statusTireApprovalDate Is Nothing AND statusTireApprovalDate <> "" _ 
							AND NOT statusBrakeApprovalDate Is Nothing AND statusBrakeApprovalDate <> "" _
							AND NOT statusHubApprovalDate Is Nothing AND statusHubApprovalDate <> "" _
							AND NOT statusYopApprovalDate Is Nothing AND statusYopApprovalDate <> "" Then
							ReDim Preserve updateSqlList(index)
							ReDim Preserve emailList(index)
							updateStatusSql2= CreateUpdateSql(recID, "STATUS_MAIL_DATE_2", currentDate)
							updateSqlList(index) = updateStatusSql2
							emailList(index) = 2
							index = index + 1
						End If
					End If

					If statusMailDate3 Is Nothing OR statusMailDate3 = "" Then
						If NOT statusWtApprovalDate Is Nothing AND statusWtApprovalDate <> "" Then
							ReDim Preserve updateSqlList(index)
							ReDim Preserve emailList(index)
							updateStatusSql3= CreateUpdateSql(recID, "STATUS_MAIL_DATE_3", currentDate)
							updateSqlList(index) = updateStatusSql3
							emailList(index) = 3
							index = index + 1
						End If
					End If

					If statusMailDate4 Is Nothing OR statusMailDate4 = "" Then
						If NOT statusEngineApprovalDate Is Nothing AND statusEngineApprovalDate <> "" Then
							ReDim Preserve updateSqlList(index)
							ReDim Preserve emailList(index)
							updateStatusSql4= CreateUpdateSql(recID, "STATUS_MAIL_DATE_4", currentDate)
							updateSqlList(index) = updateStatusSql4
							emailList(index) = 4
							index = index + 1
						End If
					End If

					'WScript.Echo ubound(updateSqlList)
					Dim updateResult
					If ubound(updateSqlList) >= 0 Then
						updateResult = UpdateAllVariStatus(updateSqlList)

						If updateResult = True Then
							util.AnalyzeNoticeEmailList(emailList)
						End If
					End If

					Res.MoveNext 
				WEND
			End If
		End If

		CloseDB()

	End Function

	Public Sub AnalyzeNoticeEmailList(emailList)
        Dim i, mailAddress,
        For i = 0 to ubound(emailList)
            If emailList(i) = 1 Then
            End If
        Next

    End Sub

	Public Function CreateUpdateSql(recordId, updateItem, updateValue)
		Dim updateSql
		updateSql = "UPDATE all_vari_status SET "
		updateSql = updateSql & updateItem & "="
		updateSql = updateSql & "'" & updateValue & "' WHERE ID = "
		updateSql = updateSql & recordId

		CreateUpdateSql = updateSql
		'WScript.Echo CreateUpdateSql
	End Function
	
End Class

















