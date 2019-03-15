'*********************************************************************
'*
'* Class Name : UtilService.class
'* author  wang.rui
'* version 1.00 2019/03/11
'*
'* History
'* 1.00 2019/03/11  FXS)wang.rui			initialize release.
'*********************************************************************
Option Explicit

Class UtilService
    Dim m_NasFolder, m_SuccessFolder, m_FailFolder, m_ErrorMailAddress
    ' Dim m_DSN, m_DataBaseServer, m_DataBaseName, m_DataBasePort, m_DataBaseUID, m_DataBasePassword, m_DataBaseDriver

    Sub Class_Initialize
        m_NasFolder = "E:\CHANYE\ROCS\NAS\"
        m_SuccessFolder = "E:\CHANYE\ROCS\SUCCESS_CSV\"
	    m_FailFolder = "E:\CHANYE\ROCS\FAIL_CSV\"
        m_ErrorMailAddress = "HGTGRP_CTRL_ROCS_ADMIN@n.t.rd.honda.co.jp"
        ' m_DSN = "ROCS"
        ' m_DataBaseDriver = "{MySQL ODBC 5.1 Driver}"
        ' m_DataBaseServer = "127.0.0.1"
        ' m_DataBasePort = "3306"
        ' m_DataBaseUID = "root"
        ' m_DataBasePassword = ""
    End Sub

    Public Sub MoveSuccessCsvFile(filePath)
        Dim objFSO
        On Error Resume Next
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        objFSO.CopyFile filePath, m_SuccessFolder
        objFSO.DeleteFile filePath
        If Err.Number <> 0  Then
			WScript.Echo Err.Number & Err.Description & Err.Source
		End If
    End Sub

     Public Sub MoveFailCsvFile(filePath)
        Dim objFSO
        On Error Resume Next
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        objFSO.CopyFile filePath, m_FailFolder
        objFSO.DeleteFile filePath
        If Err.Number <> 0  Then
			WScript.Echo Err.Number & Err.Description & Err.Source
		End If
    End Sub

    Public Sub SendErrorEmailByOutLook()
        Dim outlook, oItem, ns
        On Error Resume Next
        set outlook = WScript.CreateObject("Outlook.Application")
        set ns = outlook.getnamespace("MAPI")
        ns.logon "", "", true, false

        'set oItem = outlook.CreateItem(olMailItem)
        set oItem = outlook.CreateItem(0)

        oItem.SubJect = "VBS Test Email"
        oItem.Body = "This is a Test email"
        oItem.To = "wang.rui@cn.fujitsu.com"
        oItem.Send

        If Err.Number <> 0  Then
            WScript.Echo Err.Number
            WScript.Echo Err.Description
            WScript.Echo Err.Source
        Else
            WScript.Echo "send email success"
        End If

        set outlook = Nothing		
    End Sub

    Public Sub SendErrorEmail()
    End Sub


    Public Sub SendNoticeEmail(emailList)
        Dim i
        For i = 0 to ubound(emailList)

        Next

    End Sub

End Class