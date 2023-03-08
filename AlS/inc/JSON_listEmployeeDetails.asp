
<%    
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
    strSql = "SELECT PersonID,IDNumber, Fullname,Birthday, JoinDate, JobTitle,Department, (FirstNameLeader + ' ' + LastnameLeader) as LeaderName FROM HR_Employee" '& SearchPhrase()"
    
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
							 """Birthday"":""" & day(rsElementTem("Birthday")) &"/"& month(rsElementTem("Birthday")) &"/"& year(rsElementTem("Birthday")) & """," & _
							 """StartDate"":""" & day(rsElementTem("JoinDate")) &"/"& month(rsElementTem("JoinDate")) &"/"& year(rsElementTem("JoinDate")) & """," & _
							 """JobTitle"":""" & rsElementTem("JobTitle") & """," & _
							 """Department"":""" & trim(rsElementTem("Department")) & """," & _
							  """ReportTo"":""" & rsElementTem("LeaderName")  & """," & _
							  """StaffID"":""" & rsElementTem("IDNumber") & " ("& rsElementTem("PersonID") & ")" & """" & _
							 "}" 
                                      
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