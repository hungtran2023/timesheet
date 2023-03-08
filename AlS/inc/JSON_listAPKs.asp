
<%    
    Response.ContentType = "application/json; charset=utf-8"
    
    callback = Request("callback") ' assuming callback contains the callback function name

	strSql=	"SELECT b.ProjectID AS ProjectKey,Projectkey2, ProjectName, ISNULL(CSOFilename,'') as CSOFilename, DateTransfer, (CASE WHEN CHARINDEX('___',a.ProjectID,7) > 1 THEN 'New' ELSE 'Issued' END) AS fgStatus, " & _
				"CSOApproval,SignContract,CSOCompleted,ManagerID,billable,(c.FirstName + ' ' + c.LastName) as Manager , e.Department, f.Fullname as BDM FROM ATC_ProjectStage a INNER JOIN ATC_Projects b ON a.ProjectID=b.ProjectID " & _ 
				"INNER JOIN ATC_Department e ON b.DepartmentID=e.DepartmentID LEFT JOIN ATC_Companies d ON b.CompanyID=d.CompanyID " & _
				"LEFT JOIN ATC_PersonalInfo c ON ManagerID=c.PersonID LEFT JOIN HR_BDM f ON f.BDMID=b.BDMID " & _
				"WHERE b.fgDelete = 0  ORDER BY SortDate DESC  "
    
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
			strSignContract=""	
			if rsElementTem("SignContract")=1 then 
				strSignContract="<img src='../../images/notyet.gif'>"
			elseif rsElementTem("SignContract")=2 then
				strSignContract="<img src='../../images/icon_doc_download.gif' border=0>"
			end if
			
			if rsElementTem("CSOFileName")<>"" then strSignContract="<a href='#' path='" & strServerPath & rsElementTem("CSOFileName") & "'  class='cso'>" & strSignContract & "</a>"
			
			strArr=strArr & "{"&_
                           	 """APK"":""" & rsElementTem("ProjectKey") & """," & _
							 """ProjectName"":""" & rsElementTem("ProjectName") & """," & _
							 """RegisterDate"":""" & trim(rsElementTem("DateTransfer")) & """," & _
							  """Department"":""" & rsElementTem("Department")  & """," & _
							  """ProjectBMD"":""" & rsElementTem("BDM") & """," & _
							  """ProjectManager"":""" & rsElementTem("Manager") & """," & _
							  """InvoiceLink"":""" & "<a href='#' class='inv c'>--</a>" & """," & _
							  """CSODownload"":""" & strSignContract & """" & _
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