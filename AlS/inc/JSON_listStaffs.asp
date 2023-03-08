
<%    
    Response.ContentType = "application/json; charset=utf-8"
    
    callback = Request("callback") ' assuming callback contains the callback function name

	dateToday=date()
	
	'the last Monday from today
	dateFrom=dateToday-(Weekday(dateToday,2) + 6)			
	'the last Sunday from today
	dateTo=dateFrom + 6
	
    strSql = "SELECT PersonID, Fullname, JobTitle,Department, FirstNameLeader + ' ' + LastnameLeader as reportTo,ISNULL(b.staffID,'0') as approved FROM HR_Employee a " & _
				" LEFT JOIN (SELECT staffID FROM ATC_TimesheetApproval WHERE DateFrom='" & dateFrom & "' AND DateTo='"& dateTo & "') b ON a.PersonID=b.StaffID" &_
				" WHERE FirstName<>'Managers' ORDER BY a.FirstName"
    
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
			strApproval=""
			if rsElementTem("approved")>0 then strApproval="<img src='../../images/yes.gif'>"	
            strArr=strArr & "{"&_
                            """DT_RowId"":""" & rsElementTem("PersonID") & """," & _
							 """Fullname"":""" & rsElementTem("Fullname") & """," & _
							 """JobTitle"":""" & rsElementTem("JobTitle") & """," & _
							 """Department"":""" & trim(rsElementTem("Department")) & """," & _
							  """ReportTo"":""" & rsElementTem("reportTo")  & """," & _
							  """approved"":""" & strApproval & """" & _
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