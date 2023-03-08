<!-- #include file = "../../class/CDatabase.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					
	
	gMessage = ""
		
	varAct = Request.QueryString("act")
	if varAct = "CLEAR" then Session("filteremp") = empty
	filter_criteria = ""
	if varAct = "APPLY" then
	  criteria = Request.Form("txtname")
	  if criteria<>"" then
	  	criteria = replace(criteria, "%", "")
		criteria = replace(criteria, "#", "")
		if trim(criteria) <> "" then
			if Instr(criteria, "'")>0 then
				criteria = "#" & criteria & "#"
			else
				criteria = "'%" & criteria & "%'"
			end if
			filter_criteria = "Fullname LIKE " & criteria
		end if
	  end if	  
	  criteria = Request.Form("lstShortList")
	  'if criteria<>"" then filter_criteria = filter_criteria & " AND StaffID IN (SELECT StaffID FROM ATC_ShortListDetails WHERE ShortlistID = " & criteria & ")"
	  criteria = Request.Form("lstjobtitle")
	  if criteria<>"" then filter_criteria = filter_criteria & " AND JobTitleID = " & criteria
	  criteria = Request.Form("lstdepartment")
	  if criteria<>"" then filter_criteria = filter_criteria & " AND DepartmentID = " & criteria
	  criteria = Request.Form("lstreportto")
	  if criteria<>"" then filter_criteria = filter_criteria & " AND DirectLeaderID = " & criteria

	  if Instr(filter_criteria, " AND ") = 1 then
		filter_criteria = Mid(filter_criteria, 6, len(filter_criteria))
	  end if
	  Session("filteremp") = filter_criteria
	end if

if filter_criteria <> "" or varAct = "CLEAR" then
	
%>
<SCRIPT LANGUAGE=javascript>
<!--
	var tmp = window.opener.document.location;
	tmp = tmp.toString();
	var i2 = tmp.indexOf("?");
	if(i2==-1) { 
		i2 = tmp.length;
	}
	var i1 = tmp.lastIndexOf("/");
	scriptname = tmp.substring(0, i2);//tmp.substring(i1 + 1, i2);
	window.opener.document.forms[0].action = scriptname + "?filter=1";
	window.opener.document.forms[0].submit();
	//-->
</SCRIPT>

<%	
else
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		ret = objDb.runQuery("SELECT * FROM ATC_Shortlists WHERE OwnerID=" & Session("UserID") & " ORDER BY Shortlist")
		strOut1 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut1 = "<select name='lstShortList' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
			strOut1 = strOut1 & "<option value=''>All</option>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
			    strOut1 = strOut1 & "<option value='" & objDb.rsElement(0) & "'>" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut1 = strOut1 & "</select>"
		end if
		
		ret = objDb.runQuery("SELECT * FROM ATC_JobTitle WHERE fgActivate=1 ORDER BY JobTitle")
		strOut3 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else		
			strOut3 = "<select name='lstjobtitle' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
			strOut3 = strOut3 & "<option value=''>All</option>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
			    strOut3 = strOut3 & "<option value='" & objDb.rsElement(0) & "'>" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut3 = strOut3 & "</select>"
		end if

		ret = objDb.runQuery("SELECT * FROM ATC_Department WHERE fgActivate=1 ORDER BY Department")
		strOut4 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut4 = "<select name='lstdepartment' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
			strOut4 = strOut4 & "<option value=''>All</option>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
			    strOut4 = strOut4 & "<option value='" & objDb.rsElement(0) & "'>" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut4 = strOut4 & "</select>"
		end if
		
		strQuery = "SELECT DISTINCT a.UserID, e.Firstname + ' ' + ISNULL(e.LastName, '') + ' ' + ISNULL(e.MiddleName, '') as Fullname " &_
					"FROM ATC_UserGroup a LEFT JOIN ATC_Group b ON a.GroupID = b.GroupID " &_
					"LEFT JOIN ATC_Permissions c ON b.GroupID = c.GroupID " &_
					"LEFT JOIN ATC_Functions d ON c.FunctionID = d.FunctionID " &_
					"LEFT JOIN ATC_PersonalInfo e ON a.UserID = e.PersonID " &_
					"WHERE d.Description = 'Receive Report' AND e.fgDelete = 0 ORDER BY Fullname"
		ret = objDb.runQuery(strQuery)
		strOut5 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else		
			strOut5 = "<select name='lstreportto' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
			strOut5 = strOut5 & "<option value=''>All</option>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
			    strOut5 = strOut5 & "<option value='" & objDb.rsElement(0) & "'>" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut5 = strOut5 & "</select>"
			objDb.CloseRec
		end if
		objDb.dbdisConnect
	else
		'error in connection
		gMessage = objDb.strMessage
	end if
	set objDb = nothing
end if
%>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>
<script>
function apply() {
var tmp = document.empfilter.txtname.value
	document.empfilter.txtname.value = alltrim(tmp);	
	document.empfilter.action = "empfilter.asp?act=APPLY";
	document.empfilter.target = "_self";
	document.empfilter.submit();
}

function clear() {
	document.empfilter.action = "empfilter.asp?act=CLEAR";
	document.empfilter.target = "_self";
	document.empfilter.submit();
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="252" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" bgcolor="#003399" height="260">
  <tr> 
    <td valign="middle"> 
      <table width="250" border="0" cellspacing="0" cellpadding="0" align="center">
        <form name="empfilter" method="post">
		  <tr bgcolor=<%if gMessage="" then%>"#C0CAE6"<%else%>"#E7EBF5"<%end if%>>
            <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
          </tr>
          <tr bgcolor="C0CAE6" align="center"> 
            <td colspan="2" height="50" class="title">Filter Criteria</td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="30%" class="blue-normal" height="26" align="right"> 
              Name&nbsp;
            </td>
            <td width="70%" class="blue-normal" bgcolor="C0CAE6"> 
              <input type="text" name="txtname" class="blue-normal" size="10" maxlength="15">
            </td>
          </tr>

          <tr bgcolor="C0CAE6"> 
            <td width="30%" class="blue-normal" align="right" height="26">Short List&nbsp;</td>
            <td width="70%" class="blue-normal" bgcolor="C0CAE6"> 
<%	Response.Write strOut1%>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="30%" class="blue-normal" align="right" height="26">Job 
              Title&nbsp;</td>
            <td width="70%" class="blue-normal" bgcolor="C0CAE6"> 
<%	Response.Write strOut3%>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="30%" class="blue-normal" align="right" height="26">Department&nbsp;</td>
            <td width="70%" class="blue-normal" bgcolor="C0CAE6"> 
<%	Response.Write strOut4%>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="30%" class="blue-normal" align="right" height="26">Reports 
              To&nbsp;</td>
            <td width="70%" class="blue-normal" bgcolor="C0CAE6"> 
<%	Response.Write strOut5%>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td height="60" colspan="2"> 
              <table width="180" border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa">
                <tr> 
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
                      <a href="javascript:apply();" class="b" onMouseOver="self.status='Apply filter'; return true;" onMouseOut="self.status=''">Apply</a>
                  </td>
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center"> 
                    <a href="javascript:clear();" class="b" onMouseOver="self.status='Clear filter'; return true;" onMouseOut="self.status=''">Clear</a></td>
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
					<a href="javascript:void(0);" class="b" onClick="window.close();" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
                </tr>
              </table>
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
</body>
</html>