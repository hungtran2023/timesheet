<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%
'*********************************************************
'Generate report
'*********************************************************
Function ATSsql(byval strF,byval strT)
	dim strATS
	If year(strF)<> year(strT) then
		strATS="(SELECT ISNULL(Hours,0) as Hours,ISNULL(OverTime,0) as OverTime,AssignmentID FROM "
		For ii=year(strF) To year(strT)
			strATS=strATS & selectTable(ii)
			If ii<>	year(strT) then
				strATS=strATS & " UNION ALL SELECT ISNULL(Hours,0) as Hours,ISNULL(OverTime,0) as OverTime,AssignmentID FROM "
			else
				strATS=strATS & ")"
			end if
		Next
	else
		strATS=selectTable(year(strT))
	end if
	ATSsql=strATS
End function
'*********************************************************
'Generate report
'*********************************************************
Function GenerateReport()
	dim strSql, rsPro,strReturn
	strSql="SELECT DISTINCT(PRO.ProjectID),REPLACE(UPPER(RIGHT(PRO.ProjectID,LEN(PRO.ProjectID)-CHARINDEX('_',PRO.ProjectID))),'_',' ') + ' - ' + " & _
				"ProjectName AS CurrentJob,ISNULL(CSOMainHours,0) as CSOMainHours,ManHours " & _
			"FROM ATC_Projects PRO INNER JOIN ATC_ProjectStage PROSTAGE ON PRO.ProjectID=PROSTAGE.ProjectID " & _
				"INNER JOIN (SELECT c.ProjectID, SUM((Hours + OverTime)) as ManHours " & _
						"FROM " & ATSsql("1/1/2000",date()) & " a INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
								"INNER JOIN ATC_Tasks c ON b.SubtaskID=c.SubtaskID " & _
						"GROUP BY c.ProjectID) ATS ON PRO.ProjectID=ATS.ProjectID " & _
			"WHERE fgActivate=1 and fgDelete=0 AND LEFT(PRO.ProjectID,3)<>'ATL' ORDER BY PRO.ProjectID"
			
'Response.Write strSql
'Response.End
			
	Call GetRecordset(strSql,rsPro)
	
	if gMessage="" then
		intNo=0
		if rsPro.RecordCount>0 then
			rsPro.MoveFirst
			strBgcolor="#FFFFFF"
			Do while not rsPro.EOF
				intNo=intNo + 1
				strReturn=strReturn & "<tr valign='top' bgcolor='" & strBgcolor & "' class='blue-normal'> " & vbCrLf
				strReturn=strReturn & " <td align='center'>" & intNo & "</td> " & vbCrLf
				strReturn=strReturn & " <td>" & rsPro("CurrentJob") & "</td> " & vbCrLf
				strReturn=strReturn & " <td align='right'>" & rsPro("CSOMainHours") & "</td> " & vbCrLf
				strReturn=strReturn & " <td align='right'>" & rsPro("ManHours") & "</td> " & vbCrLf
				if cdbl(rsPro("CSOMainHours"))>0 then
					dblEstimate=FormatNumber(((cdbl(rsPro("ManHours"))/cdbl(rsPro("CSOMainHours"))) * 100),2) & "%"
				else
					dblEstimate="-"
				end if
				strReturn=strReturn & " <td align='right'>" & dblEstimate & "</td> " & vbCrLf
				strReturn=strReturn & " <td>&nbsp;</td> " & vbCrLf
				strReturn=strReturn & " <td>&nbsp;</td> " & vbCrLf
				strReturn=strReturn & "</tr>" & vbCrLf
				rsPro.MoveNext
			Loop		
		end if
	end if
	GenerateReport=strReturn
End Function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If
'-------------------------------
' Calculate pagesize
'-------------------------------
	if not isEmpty(session("Preferences")) then
		arrPre = session("Preferences")
		if arrPre(1, 0)>0 then PageSize = arrPre(1, 0) else PageSize = PageSizeDefault
		set arrPre = nothing
	else
		PageSize = PageSizeDefault
	end if
	
'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.ServerVariables("URL") 
	while Instr(tmp, "/")<>0
		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
	Wend
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../welcome.asp")
	end if
'----------------------------------
' Get report
'----------------------------------	
Dim intMonth,intYear,arrlstDay(2),strprintdate,strfromto,strfrom,strTo
Dim gMessage,intnumMonth,intTypeOfSearch

strprintdate=FormatDateTime(cdate(month(date) & "/" & Day(date) & "/" & year(date)),1)

intMonth=Request.Form("")
intMonth=month(Date())
intYear=year(date())

'0: For all activated projects
'1: For projects in month
intTypeOfSearch=Request.Form("optType")
if intTypeOfSearch="" then intTypeOfSearch=1 

strLast = GenerateReport()
'----------------------------------
' Get Company Information
'----------------------------------
if isEmpty(session("arrInfoCompany")) then
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
		strQuery = "SELECT a.CompanyName, isnull(Address,'') Address, isnull(City,'') City, isnull(b.CountryName,'') Country, " &_
					"isnull(Phone,'') Phone, isnull(Fax,'') Fax, isnull(c.Logo,'') Logo FROM ATC_Companies a " &_
					"LEFT JOIN ATC_Countries b On a.CountryID = b.CountryID " &_
					"LEFT JOIN ATC_CompanyProfile c ON a.CompanyID = c.CompanyID " &_
					"WHERE a.CompanyID = " & session("Inhouse")
		If objDb.runQuery(strQuery) Then
			If not objDb.noRecord then
				arrInfoCompany = objDb.rsElement.getRows
				session("arrInfoCompany") = arrInfoCompany
				objDb.closerec
			end if
		Else
			gMessage = objDb.strMessage
		end if
		objDb.dbDisconnect
	Else
		gMessage = objDb.strMessage
	End if
	set objDb = nothing
end if
'----------------------------------
' Get Full Name and Job Title
'----------------------------------

	Set objEmployee = New clsEmployee
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'><a href='../welcome.asp?choose_menu=B' class='c' onMouseOver='self.status=&quot;Return Main menu&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				"<a class='c' href='javascript:_print();' onMouseOver='self.status=&quot;Print report&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Print</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing	
	
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_report.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
if not isEmpty(session("arrInfoCompany")) then
	arrTmp = session("arrInfoCompany")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", arrTmp(0, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", arrTmp(1, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", arrTmp(2, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", arrTmp(3, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", arrTmp(4, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", arrTmp(5, 0))
	if arrTmp(6, 0)<>"" then
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "<img src='../images/" & arrTmp(6, 0) & "' border='0'>" )
	else
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;" )
	end if
	set arrTmp = nothing
else
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;")
end if
%>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
<script>
var objWindowSumPro;

function _print() { //v2.0
var str1 = "<%=strfromto%>";
	str1 = escape(str1);
var str2 = "<%=strprintdate%>";
	str2 = escape(str2);
var fgprint = 10;
if (fgprint!=0) {
	window.status = "";
	strFeatures = "top="+(screen.height/2-275)+",left="+(screen.width/2-390)+",width=800,height=550,toolbar=no," 
	            + "menubar=yes,location=no,resizable=yes,directories=no,scrollbars=yes,status=yes";
	if ((objWindowSumPro) && (!objWindowSumPro.closed)) {
		objWindowSumPro.focus();
	
	} else {
		objWindowSumPro = window.open("p_DailyUtilisation.asp?fromto=" + str1 + "&printdate=" + str2 + "&num=" + 7, "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";
  }
else
	alert("No data for your request.")
}

function window_onunload() {
	if((objWindowSumPro) && (!objWindowSumPro.closed))
		objWindowSumPro.close();
}

function _submit() {
	document.frmreport.action = "OngoingProjects.asp";
	document.frmreport.target = "_self" ;
	document.frmreport.submit();
	
}


</script>
</head>


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE=javascript onunload="return window_onunload()">
<form name="frmreport" method="post">
	<%
	'--------------------------------------------------
	' Write the header of HTML page
	'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
	%>	
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="445" style=height:"76%"  align="center" >
    <tr> 
      <td bgcolor="#FFFFFF" valign="top">
      
        <table width="100%" border="0" cellspacing="0" cellpadding="0">

          <tr> 
            <td ><img src="../IMAGES/dot1px.gif" width="1" height="10"></td>
          </tr>
        </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="60%">
				<table width="100%" border="0" cellpadding="2" cellspacing="0" >
                <tr> 
				  <td width="5%"><input type="radio" name="optType" value="0" <%if cint(intTypeOfSearch)=0 then%>checked<%end if%>></td>
                  <td width="45%" class="blue-normal" colspan="2" >&nbsp;All actived projects</td>
                  
				<td width="40%" class="blue-normal" align="left"><img src="../IMAGES/dot1px.gif" width="1" height="10"></td>                  
                </tr>
                
                <tr> 
				  <td><input type="radio" name="optType" value="1" <%if cint(intTypeOfSearch)=1 then%>checked<%end if%>></td>
                  <td class="blue-normal">&nbsp;Project in Month&nbsp;</td>
                  <td>
					<select name='lstmonth' size='1' height='26px' width='50px' style='width:50px;height=24px; background-color: #ffffff; border-style:1px; border-color: #A0AEA4' class='blue-normal'>
						<%For iM=1 to 12%>
						<option value='<%=iM%>' <%if cint(iM)=cint(intMonth) then%>selected<%end if%>><%=MonthName(iM,true)%></option>
						
						<%  next%>
					</select>
					<select name='lstyear' size='1' height='26px' width='50px' style='width:50px;height=24px; background-color: #ffffff; border-style:1px; border-color: #A0AEA4' class='blue-normal'>
					<%For ii=Year(Date)-1 To Year(Date)%>
						<option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
					<%Next%>
					</select>                   
                  </td>
                  <td class="blue-normal" align="left">
					<table width="60" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
						<tr> 
							<td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
								<a href="javascript:_submit();" class="b">Submit</a> </td>
						</tr>
					</table>
				</td>
				<td class="blue-normal" align="left"><img src="../IMAGES/dot1px.gif" width="1" height="10"></td>
                  
                </tr>
              </table></td>
            <td width="40%"></td>
          </tr>
          
        </table>
      
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor=<%If gMessage="" Then%>"FFFFFF"<%Else%>"#E7EBF5"<%End If%>>
		  <td class="red" height="20" align="left" width="100%"> &nbsp;<b><%=gMessage%></b></td>
	    </tr>
        <tr> 
          <td bgcolor="#8CA0D1"><img src="../../IMAGES/DOT-01.GIF" width="1" height="1"></td>
        </tr>
        <tr> 
          <td>&nbsp; </td>
        </tr>
      </table>
      
      
      
<%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "2D PROJECTS MAN HOUR - ONGOING PROJECTS")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", "")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
%>

			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr> 
					<td bgcolor="#617DC0"> 
						<table width="100%" border="0" cellspacing="1" cellpadding="5">
							<tr> 
								<td class="blue" align="center" width="4%" bgcolor="#E7EBF5">No.</td>
								<td class="blue" align="center" width="32%" bgcolor="#E7EBF5">Current Project</td>
								<td class="blue" align="center" width="8%" bgcolor="#E7EBF5">Estimated total MH</td>
								<td class="blue" align="center" width="9%" bgcolor="#E7EBF5">Actual MH to date</td>
								<td class="blue" align="center" width="10%" bgcolor="#E7EBF5">Estimated %<br> from MD</td>
								<td class="blue" align="center" width="13%"bgcolor="#E7EBF5">Work % from<br>program</td>
								<td class="blue" align="center" width="24%" bgcolor="#E7EBF5">Notes</td>
							</tr>
			<%=strLast%>
						</table>
					</td>
				</tr>
			</table>
      </td>
    </tr>
</table>
<%			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
%>
<input type="hidden" name="txtBillable" value="">
</form>
</body>
</html>