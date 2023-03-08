<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' function: OutBody
' Description:
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc, ByVal psize, Byval whichpage)
	Dim strOut
	
	strOut = ""
	arrTmp = session("arrASS")

	topofpage = (whichpage-1)*psize
	if not rsSrc.EOF then
		cnt = 0
		For i = 1 to psize
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"

			strCHK = ""
			'if arrTmp(0, topofpage + i - 1) = 1 then strCHK = "checked"

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
					"<td valign='top' class='blue'>&nbsp;" & Showlabel(rsSrc("Username")) & "</td>" & chr(13) &_
					"<td valign='top' class='blue'>&nbsp;" & Showlabel(rsSrc("Fullname")) & "</td>" & chr(13) &_
                    "<td valign='top' class='blue-normal'>&nbsp;" & Showlabel(rsSrc("JobTitle")) & "</td>" & chr(13) &_
                    "<td valign='top' class='blue-normal' align='center'>"

			strOut = strOut & "<input type='radio' name='chkass' value='" & rsSrc("PersonID") & "'" & " " & strCHK & "></td>" & chr(13)
			strOut = strOut & "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	
	Outbody = strOut
end function

'--------------------------------------
Dim gMessage, PageSize
Dim arrlstFrom(2),arrlongmon
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then	Response.Redirect("../../message.htm")

'-----------------------------------
'Check VIEWALL right
'-----------------------------------
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = "view all" then
				fgRight=true
				exit for
			end if
		next
		set getRight = nothing
	end if
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
	
'----------------------------------	
' Starting
'----------------------------------	
	strAct = Request.QueryString("act")
    intUsertype= Request.form("radUser")
	strURLBack = "AtlasComputer.asp?ID=-1"

    if cint(intUsertype)=1 then
			
	    strSql= "SELECT PersonID,Fullname,Username ,Jobtitle " & _
			    "FROM HR_Employee INNER JOIN ATC_Users ON HR_Employee.PersonID = ATC_Users.UserID "' & _
			    '"UNION ALL SELECT TPUserID as PersonID, Fullname, UserName, '' as Jobtitle FROM dbo.HR_TPStaff"
    else
        strSql= "SELECT TPUserID as PersonID, Fullname, UserName, '' as Jobtitle FROM dbo.HR_TPStaff"
    end if
		
	strSearch=Request.Form("txtsearch")	
	
	if trim(strSearch)<>"" then		
		strSearch=Request.Form("lstType") & " like '%" & strSearch & "%'"
		strSql =strSql & " WHERE " & strSearch
	end if
	
	strSql=strSql &	" ORDER BY Fullname"

	call GetRecordset(strSql,rsSrc)

	strLast=Outbody(rsSrc,200,1)

'------------------------------------	
' Get Full Name
'------------------------------------
	If IsEmpty(Session("strHTTP")) Then	Call MakeHTTP

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing

	'Make list of menu
	If isEmpty(session("Menu")) then 
		getRes = getarrMenu(session("USERID"))
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End if
	
	'current URL
	if Request.ServerVariables("QUERY_STRING")<>"" then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	else
		strURL = Request.ServerVariables("URL")
	end if
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AF"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

	arrlstFrom(0) = selectmonth("lstmonthF",month(Date()) , -1)
	arrlstFrom(1) = selectday("lstdayF", day(date()), -1)
	arrlstFrom(2) = selectyear("lstyearf", year(date()), 1999, year(date())+2, 0)
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if
%>	
<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>
<script>
<!--
function search() {
	var tmp = document.selectemployee.txtsearch.value
	if (tmp != "") {
		document.selectemployee.action = "selectemployee_ass.asp?search=" + tmp;
		document.selectemployee.target = "_self";
		document.selectemployee.submit();
	}
}


function showall() {
	document.selectemployee.txtsearch.value=""
	document.selectemployee.action = "selectemployee_ass.asp"
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
	
}

function assignment() {
	
	var userID = getCheckedValue(document.selectemployee.elements['chkass'])
	
	if (userID=="")
		alert ("Please choose a staff for assignment.");
	else
		{
			document.selectemployee.txtUserID.value=userID
			document.selectemployee.action = "AtlasComputer.asp?act=out";
			document.selectemployee.target = "_self";
			document.selectemployee.submit();
		}
	
}

function BackPrevious(strURL) {

	document.selectemployee.action = strURL;
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
	
}

// return the value of the radio button that is checked
// return an empty string if none are checked, or
// there are no radio buttons
function getCheckedValue(radioObj) {
	if(!radioObj)
		return "";
	var radioLength = radioObj.length;
	if(radioLength == undefined)
		if(radioObj.checked)
			return radioObj.value;
		else
			return "";
	for(var i = 0; i < radioLength; i++) {
		if(radioObj[i].checked) {
			return radioObj[i].value;
		}
	}
	return "";
}


-->
</script>
</head>


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onunload="return window_onunload();">
<form name="selectemployee" method="post">
<% If gMessage<>"" OR strAct<>"SAVE" then %>
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			'begin of @@Conntent
			%>

  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
    <tr> 
      <td height="90"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
		    <td class="red" colspan="4" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		  </tr>
			<tr> 
			  <td class="blue" width="11%" valign="middle">&nbsp;&nbsp;
					<a href="javascript:BackPrevious('<%=strURLBack%>');" onMouseOver="self.status='Go to Previous page'; return true;" onMouseOut="self.status=''">Back</a></td>
			  <td class="blue-normal" align="right" width="22%" valign="middle"> Search for&nbsp;&nbsp; </td>
			  <td align="right" width="27%" valign="middle"> 
			    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:150" value="">
			  </td>
			 <td align="right" width="27%" valign="middle"> 
			    <select name='lstType' size='1' height='26px' width='70px' style='width:95%;height=24px;' class='blue-normal'>
						<option value='UserName'>User Name</option>
						<option value='Fullname'>Full name</option>
						<option value='Firstname'>First name</option>
					</select>
                    
			  </td>
			  <td class="blue" align="right" width="40%" valign="middle"> 
			    <table width="160" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
			      <tr> 
			        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
			            <a href="javascript:search();" class="b" onMouseOver="self.status='Search for Fullname'; return true;" onMouseOut="self.status=''">Search</a></td>
			        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
						<a href="javascript:showall();" class="b" onMouseOver="self.status='Show all of employees'; return true;" onMouseOut="self.status=''">Show All</a></td>
			       </tr>
			    </table>
			  </td>
			</tr>        
          
          <tr> 
            <td class="title" height="50" align="center" colspan="5"> Select Employees </td>
          </tr>
          
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr> 
           <td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" height="200">
			 <tr>
               <td bgcolor="#FFFFFF" valign="top">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
	             <tr>
	              <td bgcolor="#617DC0"> 
	              <table width="100%" border="0" cellspacing="1" cellpadding="5">
	                <tr bgcolor="8CA0D1"> 
						<td class="blue" bgcolor="8CA0D1" width="20%">&nbsp;
						User Name</td>	                
	                  <td class="blue" bgcolor="8CA0D1" width="35%">&nbsp;
						Full Name</td>
	                  <td class="blue" width="35%">&nbsp;Job Title </td>
	                  <td class="blue" align="center" width="10%">&nbsp;</td>
	                </tr>
	<%
				Response.Write strLast
	%>
	              </table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">

			  <tr> 
			    <td bgcolor="#FFFFFF" height="20" class="blue-normal" align="center"> 
			      <table width="120" border="0" cellspacing="5" cellpadding="0" height="20">
			        <tr> 
			          <td align="center" class="blue" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
			            <a href="javascript: assignment();" class="b" onMouseOver="self.status='Assign'; return true;" onMouseOut="self.status=''">Assign</a>
			          </td>
			          <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
			          <a href="javascript:BackPrevious('<%=strURLBack%>');" class="b" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
			        </tr>
			      </table>
			    </td>
			  </tr>

			</table>
	              </td>
	             </tr>
	            </table>
	           </td>
	          </tr>
	         </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td align="right" valign="bottom" bgcolor="#E7EBF5">
		      <table width="100%" border="0" cellspacing="1" cellpadding="0" height="20">
		        <tr class="black-normal"> 
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page </td>
		          <td align="center" valign="middle" width="13%" class="blue-normal"> 
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPagesee")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPagesee")%>/<%=session("NumPagesee")%>&nbsp;&nbsp;</td>
		          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();" onMouseOver="self.status='Go to previous page'; return true;" onMouseOut="self.status=''">Previous</a> /
		          <a href="javascript:next();" onMouseOver="self.status='Go to next page'; return true;" onMouseOut="self.status=''"> Next</a>&nbsp;&nbsp;&nbsp;</td>
		        </tr>
		      </table>
      </td>
    </tr>
  </table>


<%'end of @@content
  Response.Write(arrTmp(1))
%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))
			
end if%>			
	<input type="hidden" name="txtComputerName" value="<%=Request.Form("txtComputerName")%>">
	<input type="hidden" name="lbType" value="<%=Request.Form("lbType")%>">
	<input type="hidden" name="txtIP" value="<%=Request.Form("txtIP")%>">
	<input type="hidden" name="txtUserID" value="<%=Request.Form("txtUserID")%>">
	<input type="hidden" name="radUser" value="<%=Request.Form("radUser")%>">
		
	<input type="hidden" name="txtURLBack" value="<%=strURLBack%>">
	<input type="hidden" name="txtID" value="<%=Request.Form("txtID")%>">
	<input type="hidden" name="txtAtlasPCID" value="<%=Request.Form("txtAtlasPCID")%>">
	
</form>  
</table>

</body>
</html>

 