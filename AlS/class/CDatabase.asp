<!--#include file="../inc/adovbs.inc"-->
<SCRIPT LANGUAGE="VBScript" RUNAT="Server">


'**************************************************
' Copyright (C) by Atlas Industries Limited
' E-mail: info@atlasindustries.com
'
' CLASS NAME:
'
'		clsDatabase
'
' DESCRIPTION:
'
'
' METHODS:
'
'   Function dbConnect() As Boolean
'   Sub dbDisConnect()
'   Function recConnect() As Boolean
'	Sub recDisconnect()
'   Function openRec() As Boolean
'	Sub closeRec()
'   Function MoveFirst() As Boolean
'   Function MoveNext() As Boolean
'   Function MovePrevious() As Boolean
'   Function MoveLast() As Boolean
'   Function recSort() As Boolean
'   Function runActionQuery()
'   Function runQuery()
'	Function noRecord() AS Boolean
'
' AUTHOR:
' DATE:
'
' NOTE:
'**************************************************

Class clsDatabase 
    Public cnDatabase
    Public rsElement
    Public strMessage
    
'**************************************************
' Function: dbConnect
' Description: Connect to SQL server by connection pooling
' Parameters: - strConn: String 
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

    Public Function dbConnect(ByVal strConn)
		On Error Resume Next
		dbConnect = False
		 
		Set cnDatabase = Server.CreateObject("ADODB.Connection")
		cnDatabase.Open strConn & "APP=Timesheet System - Pooled connection; OLE DB Services=-1;"
		
		If Err.number = 0 then
			dbConnect = True
		Else
			strMessage = Err.Description				
		End If
		Err.Clear
    End function
    
'**************************************************
' Sub: dbDisConnect
' Description: Disconnect to SQL server
' Parameters: None
' Return value: None
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************
    
    Public Sub dbDisConnect()
		cnDatabase.Close()
		Set cnDatabase = nothing
    End Sub	' dbDisConnect

'**************************************************
' Sub: recConnect
' Description: Permits to access data in a disconnected mode.
' Parameters: - strConn: String 
' Return value: None
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************
	
	Public Function recConnect(Byval strConn)
		On Error Resume Next
		recConnect = False

		Set rsElement = Server.CreateObject("ADODB.Recordset")

		rsElement.ActiveConnection = strConn & "APP=Timesheet System - Pooled connection; OLE DB Services=-1;"

' Obligatoiry to get the recordset persisted.
		rsElement.CursorLocation = adUseClient
		
		
		If Err.number = 0 then
			recConnect = True
		Else
			strMessage = Err.Description				
		End If
		Err.Clear
		
	End Function	' recConnect

'**************************************************
' Sub: recDisConnect
' Description: 
' Parameters: None 
' Return value: None
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************
	
	Public Sub recDisConnect()
		Set rsElement.ActiveConnection = Nothing
		'rsElement.MoveFirst
	End Sub	' recDisConnect		
	
'**************************************************
' Function: openRec
' Description: Open recordset in a forward & backward supported mode
' Parameters: - strSQL: String
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

	Public Function openRec(Byval strSQL)
		On Error Resume Next
		openRec = False
		
		If (mid(str,len(strSQL),1)=";") then
			 strSQL = mid(strSQL,1,len(strSQL)-1)
		End If
		
		
		rsElement.Open strSQL, , adOpenStatic, adLockBatchOptimistic, adCmdText

		If Err.number = 0 then
			openRec = True
		Else
			strMessage = Err.Description	
		End If
		Err.Clear
	End Function	' openRec

'**************************************************
' Function: openRec
' Description: Open recordset in a forward & backward supported mode
' Parameters: - strSQL: String
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

	Public Function openRecordset(ByRef myCmd)
		On Error Resume Next
		openRecordset = False
		
		rsElement.Open myCmd, , adOpenStatic, adLockBatchOptimistic

		If Err.number = 0 then
			openRecordset = True
		Else
			strMessage = Err.Description	
		End If
		Err.Clear
	End Function	' openRec
	
'**************************************************
' Sub: closeRec
' Description: Close recordset
' Parameters: None 
' Return value: None
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

	Public Sub closeRec()
		rsElement.Close()
		set rsElement = Nothing
	End Sub	' closeRec
	
'**************************************************
' Function: MovePrevious
' Description: Go to the previous record
' Parameters: None
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************
	
    Public Function MovePrevious()
		On Error Resume Next
		MovePrevious = False
	
		rsElement.MovePrevious
		
		If Err.number = 0 then
			MovePrevious = True
		Else
			strMessage = Err.Description	
		End If
		Err.Clear
	End Function

'**************************************************
' Function: MoveNext
' Description: Go to the next record
' Parameters: None
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

	Public Function MoveNext()
		On Error Resume Next
		MoveNext = False
	
		rsElement.MoveNext
		
		If Err.number = 0 then
			MoveNext = True
		Else
			strMessage = Err.Description	
		End If
		Err.Clear
	End Function

'**************************************************
' Function: MoveFirst
' Description: Go to the first record
' Parameters: None
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

	Public Function MoveFirst()
		On Error Resume Next
		MoveFirst = False
	
		rsElement.MoveFirst
		
		If Err.number = 0 then
			MoveFirst = True
		Else
			strMessage = Err.Description	
		End If
		Err.Clear

	End Function
	
'**************************************************
' Function: MoveLast
' Description: Go to the last record
' Parameters: None
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************
	
	Public Function MoveLast()
		On Error Resume Next
		MoveLast = False
	
		rsElement.MoveLast
		
		If Err.number = 0 then
			MoveLast = True
		Else
			strMessage = Err.Description	
		End If
		Err.Clear
	End Function

'**************************************************
' Function: recSort
' Description: Return a sorted recordset
' Parameters: - FieldName: String
'			  - Order: A _ Ascending; D _ Descending 
' Return value: None
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************
	
	Public Function recSort(Byval FieldName, Byval Order)
		If Order = "A" Then
			rsElement.Sort = left(FieldName,len(FieldName)-1) & " ASC" & chr(34)   
		Else
			rsElement.Sort = left(FieldName,len(FieldName)-1) & " DESC" & chr(34) 
		End If
	End Function

'**************************************************
' Function: runActionQuery
' Description: Execute a query
' Parameters: strSQL
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

    Public Function runActionQuery(ByVal strSQL)
		On Error Resume Next
		runActionQuery = False
		
		if (mid(strSQL,len(strSQL),1)=";") then
			 strSQL = mid(strSQL,1,len(strSQL)-1)
		end if
		cnDatabase.Execute(strSQL)
		
		If Err.number = 0 then
			runActionQuery = True
		Else
			strMessage = Err.Description
		End If
		Err.Clear
    End Function

'**************************************************
' Function: runQuery
' Description: Execute a query and put the result into a recordset
' Parameters: strSQL
' Return value: true if success, false if not
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************

	Public Function runQuery(ByVal strSQL)
		On Error Resume Next
		runQuery = False
		
		If (mid(strSQL,len(strSQL),1)=";") Then
    		strSQL = mid(strSQL,1,len(strSQL)-1)
    	End If   	
  	 		
    	Set rsElement = cnDatabase.Execute(strSQL)
		
		If Err.number = 0 Then
			runQuery = True
		Else
			strMessage = Err.Description	
		End If
		Err.Clear
	End Function

'**************************************************
' Function: noRecord
' Description: Have any record or not
' Parameters: None
' Return value: true if have not any record, false if have
' Author: 
' Date: 20/06/2001
' Note:
'**************************************************

	Public Function noRecord()
		If Not rsElement.EOF Or rsElement.RecordCount > 0 Then
			noRecord = False
		Else
			noRecord = True
		End If
	End Function

'**************************************************
' Function: getColumn_by_name
' Description: Return the value of column
' Parameters: strColName
' Return value: None
' Author: 
' Date: 15/06/2001
' Note:
'**************************************************
	
	Public Function getColumn_by_name(ByVal strColName)
		If Not rsElement.eof Then
			getColumn_by_name = rsElement(strColName)
		Else
			getColumn_by_name = ""
		End If
    End Function

End Class
</SCRIPT>