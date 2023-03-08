
<%    
'**************************************************
' Sub: MakeHTTP
' Description: http://newtimesheet/
' Parameters: userID of login user
' Return value: array
' Author:
' Date: 28/06/2001
' Note:
'**************************************************
private function MakeHTTP_()
	'create http path
	Dim strTmp, strHTTP
	strTmp = Request.ServerVariables("URL")
	strTmp = Mid(strTmp , 1, Instr(2, strTmp, "/")-1)
	strHTTP = "http://" & Request.ServerVariables("SERVER_NAME") & strTmp & "/"
	
	MakeHTTP_=strHTTP
	
End Function

	Response.ContentType = "application/json; charset=utf-8"    
    callback = Request("callback") ' assuming callback contains the callback function name

	intUserID = session("USERID")
'--------------------------------------------------
' Check VIEWALL right
'--------------------------------------------------	
	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "view all" Then
				fgRight = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------
' Check Enter Annual Leave right
'--------------------------------------------------
	fgHREnter=false
	If not isEmpty(session("RightOn")) Then
		varGetRight = session("RightOn")
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "Write Timesheet as HR control" OR intUserID=252 Then
				fgHREnter = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------		

    strSql = "SELECT PersonID, Fullname, JobTitle,Department, FirstNameLeader + ' ' + LastnameLeader as reportTo FROM HR_Employee "
	
	
	if not fgRight then strSql=strSql & " WHERE PersonID IN ( SELECT StaffID FROM UserByReportTo("& intUserID & ") )"
	
	strSql=strSql & " ORDER BY Fullname"
    
    strconn=Application("g_strConnect")	

    set conTem=Server.CreateObject("ADODB.Connection")
    conTem.Open(strconn)

    Set rsElementTem = Server.CreateObject("ADODB.Recordset")
    rsElementTem.Open strSql,conTem,3,3
    strArr=""
	i=0
    if not rsElementTem.EOF then
        rsElementTem.MoveFirst
                
        do while not rsElementTem.EOF
			i=i+1
            if strArr<>"" then strArr=strArr & ","
			
            strArr=strArr & "{"&_
                            """DT_RowId"":""" & rsElementTem("PersonID") & """," & _
							 """Fullname"":""" & rsElementTem("Fullname") & """," & _
							 """JobTitle"":""" & rsElementTem("JobTitle") & """," & _
							 """Department"":""" & trim(rsElementTem("Department")) & """," & _
							  """ReportTo"":""" & rsElementTem("reportTo")  & """" 
			if fgHREnter then
				strArr=strArr & ",""Duration"":""" & _
							"<a href='" & MakeHTTP_() & "aisnet/HREnterTimeSheet/HREnterTimeSheet/" & rsElementTem("PersonID") & "' class='tms'>by Days <span class='glyphicon glyphicon-triangle-right'></a></button>"""

			end if
							  
			strArr=strArr & "}" 
                                      
            rsElementTem.MoveNext
        loop
    end if
	Response.Write "{"
	Response.Write """draw"": 1,"
	Response.Write """recordsTotal"":"&  i & ","
	Response.Write """recordsFiltered"":" & i & ","
	Response.Write """data"": "
    Response.Write "[" & strArr & "]}"
%> 