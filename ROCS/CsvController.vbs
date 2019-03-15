'*********************************************************************
'*
'* Class Name : CsvController.class
'* author  wang.rui
'* version 1.00 2019/03/08
'*
'* History
'* 1.00 2019/03/08  FXS)wang.rui			initialize release.
'*********************************************************************
Option Explicit

Class CsvController
	Dim sqlController
	Dim utils

	Sub Class_Initialize
		ImportFile("MySQLController.vbs")
		ImportFile("UtilService.vbs")
		Set sqlController = New MySQLController
		Set utils = New UtilService
	End Sub

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

	'*******************************************************************************************************************

	'遍历NAS文件夹
	Public Function TraverseNasFolder(nasFolder)
		Dim oFso, oFolder, oFiles, oFile
		Set oFso = CreateObject("Scripting.FileSystemObject") 
		Set oFolder = oFso.GetFolder(nasFolder)  
		Set oFiles = oFolder.Files  
		
		Dim statusCsvList(), factorCsvList()
		ReDim Preserve statusCsvList(0)
		ReDim Preserve factorCsvList(0)
		'遍历所有csv文件
		Dim index1, index2
		index1 = 0
		index2 = 0
		For Each oFile In oFiles
			If InStr(oFile.Path, "_status.csv") <> 0 Then
				statusCsvList(index1) = oFile.Path
				index1 = index1+1
				ReDim Preserve statusCsvList(index1)
			ElseIf InStr(oFile.Path, "_factor.csv") <> 0 Then
				factorCsvList(index2) = oFile.Path
				index2 = index2+1
				ReDim Preserve factorCsvList(index2)
			End If
		Next 

		If ubound(statusCsvList) = 0 And ubound(factorCsvList) = 0 Then
			WScript.Echo "no csv file! exit"
		Else
			If ubound(statusCsvList) <> 0 Then
				WScript.Echo "start to read status csv files..."
				ReadStatusCsv(statusCsvList)
			End If

			If ubound(factorCsvList) <> 0 Then
				WScript.Echo "start to read factor csv files..."
				ReadFactorCsv(factorCsvList)
			End If
		End If		

		Set oFolder = Nothing  
		Set oFso = Nothing	
	End Function
	
	Public Function ReadFactorCsv(csvList)
		Dim objFSO, objFile, strLine, i
		Dim insertFactorSql, bConnRet, bInsertRet
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		For i = 0 to ubound(csvList)-1
			Dim stm, strArray1
			Set stm = CreateObject("Adodb.Stream") 
			stm.Type = 2 
			stm.mode = 3 
			stm.charset = "utf-8" 
			stm.Open 
			stm.LoadFromFile csvList(i) 
			strArray1 = split(stm.readtext, vbCrLf)
			stm.close 			

			Dim j, StrArray2
			For j = 0 to ubound(strArray1)
				StrArray2 = split(strArray1(j), ",")
				insertFactorSql = CreateInsertFactorSql(strArray2)

				bConnRet = sqlController.ConnectDB()	
				If bConnRet = False Then
					WScript.Echo "connect MySQL Fail"
					'utils.MoveFailCsvFile(csvList(i))
					'utils.SendErrorEmail()
					Exit For
				Else
					WScript.Echo "connect MySQL Success"
					bInsertRet = sqlController.InsertAllVariInputHists(insertFactorSql)
					If bInsertRet = False Then
						WScript.Echo "Insert into MySQL Fail"
						'utils.MoveFailCsvFile(csvList(i))
						'utils.SendErrorEmail()
						Exit For
					Else
						WScript.Echo "Insert into MySQL Success"
					End If	
				End If

				sqlController.CloseDB()
			Next
			
			If bInsertRet = True Then
				'utils.MoveSuccessCsvFile(csvList(i))
			End If
		Next
		Set objFSO = nothing
	End Function


	Public Function ReadStatusCsv(csvList)
		Dim objFSO, objFile, strLine,i
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		For i = 0 to ubound(csvList)-1
			'Set objFile = objFSO.OpenTextFile(csvList(i), 1)
			' Do Until objFile.AtEndOfStream  
			' 	Dim strArray
			' 	strLine = objFile.readline
			' 	'WScript.Echo strLine
			' 	strArray = split(strLine,",")
				
			' 	'For j=0 to ubound(strArray)
			' 		'WScript.Echo strArray(j)
			' 	'Next				
			' Loop
			'objFile.close

			Dim stm, strArray
			Set stm = CreateObject("Adodb.Stream") 
			stm.Type = 2 
			stm.mode = 3 
			stm.charset = "utf-8" 
			stm.Open 
			stm.LoadFromFile csvList(i) 
			strArray = split(stm.readtext,",")
			stm.close 

			Dim insertStatusSql, bConnRet, bInsertRet
			insertStatusSql = CreateInsertStatusSql(strArray)

			bConnRet = sqlController.ConnectDB()	
			If bConnRet = False Then
				WScript.Echo "connect MySQL Fail"
			Else
				WScript.Echo "connect MySQL Success"
				bInsertRet = sqlController.InsertAllVariStatus(insertStatusSql)
				If bInsertRet = False Then
					WScript.Echo "Insert into MySQL Fail"
					utils.MoveFailCsvFile(csvList(i))
					'utils.SendErrorEmail()
				Else
					WScript.Echo "Insert into MySQL Success"
					utils.MoveSuccessCsvFile(csvList(i))
				End If	
			End If

			sqlController.CloseDB()

		Next		
		Set objFSO = nothing

	End Function
	
	Public Function CreateInsertFactorSql(StrArray)
		Dim car_id, factor_id, key1, key2, key3, key4, key5, value1,value2,value3,value4,value5,value6,value7,value8
		Dim value9, value10, value11, value12, value13, value14, value15, value16,value17,value18,value19,value20
		Dim factor_update_user, factor_update_date
		Dim insertSql, factorName

		car_id = Replace(strArray(0),"'","")

		factorName = Replace(strArray(1),"'","")
		factorName = UCase(factorName) 

		If factorName = "TIRE" Then
			factor_id = 0
		ElseIf factorName = "BRAKE" Then
			factor_id = 1
		ElseIf factorName = "HUB" Then
			factor_id = 2
		ElseIf factorName = "CDXA" Then
			factor_id = 3
		End If

		key1 = Replace(strArray(2),"'","")
		key2 = Replace(strArray(3),"'","")
		key3 = Replace(strArray(4),"'","")
		key4 = Replace(strArray(5),"'","")
		key5 = Replace(strArray(6),"'","")
		value1 = Replace(strArray(7),"'","")
		value2 = Replace(strArray(8),"'","")
		value3 = Replace(strArray(9),"'","")
		value4 = Replace(strArray(10),"'","")
		value5 = Replace(strArray(11),"'","")
		value6 = Replace(strArray(12),"'","")
		value7 = Replace(strArray(13),"'","")
		value8 = Replace(strArray(14),"'","")
		value9 = Replace(strArray(15),"'","")
		value10 = Replace(strArray(16),"'","")
		value11 = Replace(strArray(17),"'","")
		value12 = Replace(strArray(18),"'","")
		value13 = Replace(strArray(19),"'","")
		value14 = Replace(strArray(20),"'","")
		value15 = Replace(strArray(21),"'","")
		value16 = Replace(strArray(22),"'","")
		value17 = Replace(strArray(23),"'","")
		value18 = Replace(strArray(24),"'","")
		value19 = Replace(strArray(25),"'","")
		value20 = Replace(strArray(26),"'","")
		factor_update_user = Replace(strArray(27),"'","")
		factor_update_date = Replace(strArray(28),"'","")

		insertSql = "INSERT INTO all_vari_input_hists (CAR_ID,"
		insertSql = insertSql & "FACTOR_ID,"
		insertSql = insertSql & "KEY1,"
		insertSql = insertSql & "KEY2,"
		insertSql = insertSql & "KEY3,"
		insertSql = insertSql & "KEY4,"
		insertSql = insertSql & "KEY5,"
		insertSql = insertSql & "VALUE1,"
		insertSql = insertSql & "VALUE2,"
		insertSql = insertSql & "VALUE3,"
		insertSql = insertSql & "VALUE4,"
		insertSql = insertSql & "VALUE5,"
		insertSql = insertSql & "VALUE6,"
		insertSql = insertSql & "VALUE7,"
		insertSql = insertSql & "VALUE8,"
		insertSql = insertSql & "VALUE9,"
		insertSql = insertSql & "VALUE10,"
		insertSql = insertSql & "VALUE11,"
		insertSql = insertSql & "VALUE12,"
		insertSql = insertSql & "VALUE13,"
		insertSql = insertSql & "VALUE14,"
		insertSql = insertSql & "VALUE15,"
		insertSql = insertSql & "VALUE16,"
		insertSql = insertSql & "VALUE17,"
		insertSql = insertSql & "VALUE18,"
		insertSql = insertSql & "VALUE19,"
		insertSql = insertSql & "VALUE20,"
		insertSql = insertSql & "UPDATE_USER_ID,"
		insertSql = insertSql & "UPDATE_DATE) VALUES ("
		insertSql = insertSql & "'" & car_id & "',"
		insertSql = insertSql & "'" & factor_id & "',"
		insertSql = insertSql & "'" & key1 & "',"
		insertSql = insertSql & "'" & key2 & "',"
		insertSql = insertSql & "'" & key3 & "',"
		insertSql = insertSql & "'" & key4 & "',"
		insertSql = insertSql & "'" & key5 & "',"
		insertSql = insertSql & "'" & value1 & "',"
		insertSql = insertSql & "'" & value2 & "',"
		insertSql = insertSql & "'" & value3 & "',"
		insertSql = insertSql & "'" & value4 & "',"
		insertSql = insertSql & "'" & value5 & "',"
		insertSql = insertSql & "'" & value6 & "',"
		insertSql = insertSql & "'" & value7 & "',"
		insertSql = insertSql & "'" & value8 & "',"
		insertSql = insertSql & "'" & value9 & "',"
		insertSql = insertSql & "'" & value10 & "',"
		insertSql = insertSql & "'" & value11 & "',"
		insertSql = insertSql & "'" & value12 & "',"
		insertSql = insertSql & "'" & value13 & "',"
		insertSql = insertSql & "'" & value14 & "',"
		insertSql = insertSql & "'" & value15 & "',"
		insertSql = insertSql & "'" & value16 & "',"
		insertSql = insertSql & "'" & value17 & "',"
		insertSql = insertSql & "'" & value18 & "',"
		insertSql = insertSql & "'" & value19 & "',"
		insertSql = insertSql & "'" & value20 & "',"
		insertSql = insertSql & "'" & factor_update_user & "',"
		insertSql = insertSql & "'" & factor_update_date & "')"

		CreateInsertFactorSql = insertSql	
	End Function

	Public Function CreateInsertStatusSql(strArray)
		Dim car_id, status_name, status_mail, status_date, status_update_user, status_update_date, status_type
		Dim insertSql

		car_id = Replace(strArray(0),"'","")
		'category = strArray(1)
		'status_type = strArray(2)
		status_name = Replace(strArray(3),"'","")
		status_mail = Replace(strArray(4),"'","")
		status_date = Replace(strArray(5),"'","")
		status_update_user = Replace(strArray(6),"'","")
		status_update_date = Replace(strArray(7),"'","")

		status_type ="STATUS_" & Replace(strArray(1),"'","") & "_" & Replace(strArray(2),"'","") & "_"

		insertSql = "INSERT INTO all_vari_status (CAR_DATA_ID,"
		insertSql = insertSql & status_type & "NAME,"
		insertSql = insertSql & status_type & "MAIL,"
		insertSql = insertSql & status_type & "DATE,"
		insertSql = insertSql & status_type & "UPDATE_USER_ID,"
		insertSql = insertSql & status_type & "UPDATE_DATE) VALUES ("
		insertSql = insertSql & "'" & car_id & "',"
		insertSql = insertSql & "'" & status_name & "',"
		insertSql = insertSql & "'" & status_mail & "',"
		insertSql = insertSql & "'" & status_date & "',"
		insertSql = insertSql & "'" & status_update_user & "',"
		insertSql = insertSql & "'" & status_update_date & "')"

		CreateInsertStatusSql = insertSql
	
	End Function	
	
End Class











