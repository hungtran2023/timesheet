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
	    
	if not rsSrc.EOF then
		cnt = 0
		For i = 1 to psize
		
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
			
			strCHK = ""

		
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
					"<td valign='top' width='55%' class='blue'>&nbsp;" & Showlabel(rsSrc("Fullname")) & "</td>" & chr(13) &_
                    "<td valign='top' width='35%' class='blue-normal'>&nbsp;" & Showlabel(rsSrc("JobTitle")) & "</td>" & chr(13) &_
                    "<td valign='top' width='10%' class='blue-normal' align='center'>"

			strOut = strOut & "<input type='checkbox' name='chkass' value='" & rsSrc("PersonID") & "'" & " " & strCHK & "></td>" & chr(13)
			strOut = strOut & "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
		
	end if
	Outbody = strOut
end function

'--------------------------------------
Dim gMessage, PageSize, intAtlasCourseID
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

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
	intAtlasCourseID=request.Form("txtAtlasCourseID")
	strURLBack=request.Form("txtpreviouspage")
	
	strAct=request.QueryString("act")
	
	if strAct="SAVE" then
		
		strCheck=Request.Form("chkass")
						
		
		strQuery="INSERT INTO ATC_AtlasCourseParticipant ([StaffID],[AtlasCourseID])" & _
					" SELECT StaffID, " & intAtlasCourseID & " FROM ATC_Employees WHERE StaffID IN (" & strCheck & ")"
		
		strCnn = Application("g_strConnect")	
		Set objDatabase = New clsDatabase     
		strError=""
		If objDatabase.dbConnect(strCnn) Then              
			if not objDatabase.runActionQuery(strQuery) then 
			   gMessage = objDatabase.strMessage
			else
				Response.Redirect "ListOfAttendeesCourse.asp?id=" & intAtlasCourseID
			end if
		end if
		
	end if
	
	strSearch = trim(Request.QueryString("search"))
	if strSearch<>""  then
			strSearch = replace(strSearch, "%", "")
			strSearch = replace(strSearch, "#", "")
			strSearch = " AND Fullname Like '%" & trim(strSearch) & "%'"
	end if
	
	strQuery = "SELECT * FROM HR_Employee WHERE PersonID not in (Select StaffID From ATC_AtlasCourseParticipant Where AtlasCourseID = " & intAtlasCourseID & _
				" ) " & strSearch & " ORDER BY Firstname"

	Call GetRecordset(strQuery,rsUser)
	curpage=1
	strLast = Outbody(rsUser, pageSize, curpage)
	
	
'------------------------------------	
' Get Full Name
'------------------------------------
	If IsEmpty(Session("strHTTP")) Then
		Call MakeHTTP
	End if
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
	if strChoseMenu = "" then strChoseMenu = "AH"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

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
var objEMFIWindow;


function window_onunload() {
	if((objEMFIWindow)&&(!objEMFIWindow.closed))
		objEMFIWindow.close();
}

//-->
function closeemp() {
	if((objEMFIWindow)&&(!objEMFIWindow.closed))
		objEMFIWindow.close();
	window.close();
}

function sort(type) {
	document.selectemployee.action = "selectemployee_ass.asp?sorttype=" + type; //1: fullname, 2: jobtitle
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
}

function search() {
	var tmp = document.selectemployee.txtsearch.value
	if (tmp != "") {
		document.selectemployee.action = "selectemployee_ass.asp?search=" + tmp;
		document.selectemployee.target = "_self";
		document.selectemployee.submit();
	}
}

function showall() {

		document.selectemployee.action = "selectemployee_ass.asp";
		document.selectemployee.target = "_self";
		document.selectemployee.submit();

}

function setchecked(val) {
  with (document.selectemployee) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].type == "checkbox") {
			elements[ii].checked = val;
		}
	}
  }
}

function checkass () {
  selection = false;
  with (document.selectemployee) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].type == "checkbox") && (elements[ii].checked==true)) {
			selection = true;
			break;
		}
	}
  }
  return(selection)
}

function assignment() {

	document.selectemployee.action = "selectemployee_ass.asp?act=SAVE";
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
	
}

function BackPrevious(strURL) {
	
	document.selectemployee.action = strURL;
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
	
}
</script>
</head>

<body style="background-color:'#FFFFFF';margin: 0px 0px 0px 0px">

<form name="selectemployee" method="post">
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
<% If gMessage<>"" OR strAct<>"SAVE" then %>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
    <tr> 
      <td height="90"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
		    <td class="red" colspan="4" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		  </tr>
			<tr> 
			  <td class="blue" width="11%" valign="middle">&nbsp;&nbsp;&nbsp;
					<a href="javascript:BackPrevious('<%=strURLBack%>');" onMouseOver="self.status='Go to Previous page'; return true;" onMouseOut="self.status=''">Back</a></td>
			  <td class="blue-normal" align="right" width="22%" valign="middle"> Search for&nbsp; </td>
			  <td  width="27%" valign="middle"> 
			    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:100%" value="">
			  </td>
			  <td class="blue"  width="40%" valign="middle"> 
			    <table width="240" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
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
            <td class="title" height="50" align="center" colspan="4"> List of Employees</td>
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
	                  <td class="blue" bgcolor="8CA0D1" width="194">&nbsp;
						<a href="javascript:sort(1);" class="c" onMouseOver="self.status='Order by Fullname'; return true;" onMouseOut="self.status=''">Full 
	                    Name</a> </td>
	                  <td class="blue" width="189">&nbsp;<a href="javascript:sort(2);" class="c" onMouseOver="self.status='Order by Job Title'; return true;" onMouseOut="self.status=''">Job 
	                    Title</a> </td>
	                  <td class="blue" align="center" width="8%">&nbsp;</td>
	                </tr>
	<%
				Response.Write strLast
	%>
	              </table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td bgcolor="#FFFFFF" height="20" class="blue" align="right"><a href="javascript:setchecked(1);" class="c" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
			      All</a>&nbsp;&nbsp;&nbsp;<a href="javascript:setchecked(0);" class="c" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''"> Clear All</a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
			  </tr>
			  <tr> 
			    <td bgcolor="#FFFFFF" height="20" class="blue-normal" align="center"> 
			      <table width="120" border="0" cellspacing="5" cellpadding="0" height="20">
			        <tr> 
			          <td align="center" class="blue" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
			            <a href="javascript: assignment();" class="b" onMouseOver="self.status='Assign'; return true;" onMouseOut="self.status=''">Assign</a>
			          </td>
			          <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
			          <a href="javascript:closeemp();" class="b" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
			        </tr>
			      </table>
			    </td>
			  </tr>
			  <tr>
			    <td bgcolor="#FFFFFF" height="20" class="blue-normal">&nbsp;&nbsp;*Click 
			      on each column header to sort the list by alphabetical order.
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
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
		          </td>
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
<%end if%>

<%'end of @@content
  Response.Write(arrTmp(1))
%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))
			%>	
	<input type="hidden" name="txtAtlasCourseID" value="<%=intAtlasCourseID%>">
	<input type="hidden" name="txtURLBack" value="<%=strURLBack%>">
</form>
</body>
</html>
<% If gMessage="" and strAct="SAVE" then %>
				<SCRIPT LANGUAGE=javascript>
				<!--
					document.selectemployee.action = "<%=strURLBack%>";
					document.selectemployee.target = "_self";
					document.selectemployee.submit();
					//-->
				</SCRIPT>
<%end if%>