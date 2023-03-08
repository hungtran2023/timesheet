<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../inc/library.asp"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If					

'----------------------------------
' Get Full Name and Job Title
'----------------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
'----------------------------------	
' Make list of menu
'----------------------------------
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
	if strChoseMenu = "" then strChoseMenu = "C"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../")
'-----------------------------------
' Analyse query string
'-----------------------------------

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_menu.htm")


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

<link rel="stylesheet" href="../timesheet.css">

<script language="javascript" src="../library/library.js"></script>
<script>


function next() {
var curpage = <%=session("CurPageshort")%>
var numpage = <%=session("NumPageshort")%>
	if (curpage < numpage) {
		document.frmshort.action = "shortlist.asp?navi=NEXT";
		document.frmshort.target = "_self";
		document.frmshort.submit();
	}
}

function prev() {
var curpage = <%=session("CurPageshort")%>
var numpage = <%=session("NumPageshort")%>
	if (curpage > 1) {
		document.frmshort.action = "shortlist.asp?navi=PREV";
		document.frmshort.target = "_self";
		document.frmshort.submit();
	}
}

function go() {
	var numpage = <%=session("NumPageshort")%>
	var curpage = <%=session("CurPageshort")%>
	var intpage = document.frmshort.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmshort.action = "shortlist.asp?Go=" + intpage;
		document.frmshort.target = "_self";
		document.frmshort.submit();		
	}
}

function setchecked(val) {
  with (document.frmshort) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].name == "chkstaff") {
			elements[ii].checked = val;
		}
	}
  }
}

function chkremove() {
  fg = false;
  with (document.frmshort) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].name == "chkstaff") && (elements[ii].checked)) {
			fg = true;
			break;
		}
	}
  }
 if (fg == false) alert("No participant selected.")
 return(fg)
}

function remove() {
  if (chkremove()==true) {
  	document.frmshort.action = "shortlist.asp?act=REMOVE";
	document.frmshort.target = "_self";
	document.frmshort.submit();
  }
}

function gopage() {
	var tmp = document.frmshort.lstshort.options[document.frmshort.lstshort.selectedIndex].text;
	document.frmshort.txtshortname.value = tmp;
  	document.frmshort.action = "shortlist.asp?act=GOPAGE";
	document.frmshort.target = "_self";
	document.frmshort.submit();
}

function addshort() {
	var tmp  = document.frmshort.txtshortlistname.value;
	document.frmshort.txtshortlistname.value = alltrim(tmp);
	if (document.frmshort.txtshortlistname.value=="") { alert("Please enter a shortlist name.") }
	else {
		document.frmshort.action = "shortlist.asp?act=ADDSHORT";
		document.frmshort.target = "_self";
		document.frmshort.submit();
	}
}

function delete_() {
  	document.frmshort.action = "shortlist.asp?act=DELETE";
	document.frmshort.target = "_self";
	document.frmshort.submit();
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmshort" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>
			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			'Response.Write(arrPageTemplate(1))
			Response.Write(arrTmp(0))
			'begin of @@Conntent
			%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor=<%if gMessage="" then%>"#FFFFFF"<%else%>"#E7EBF5"<%end if%>>
		 <td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
        <tr valign="middle"> 
          <td class="title" height="50" align="center"> Form List</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" align="center" width="5%">No.</td>
                              <td class="blue" align="center" width="35"></td>
                              <td class="blue" align="center" width="37%">Title</td>
                              <td class="blue" align="center" >Description</td>
                            </tr>
                            
                           <tr bgcolor="#FFF2F2" height="40px" valign='middle'>
                                <td align='center' class='blue-normal' >1.</td>
							    <td class='blue-normal'><img src="http://ais.atlasindustries.com/timesheet/images/word.png" /></td>
							    <td class='blue'><a href='http://ais.atlasindustries.com/timesheet/Data/HR%20Documents/OT Form_Staff_2021.docx' class='c'>OVERTIME FORM - REGULAR</a></td>
							    <td class='blue-normal'>Used for all OT requests</td>
							    </tr>
                            <tr bgcolor="#E7EBF5" height="40px" valign='middle'>
                                <td class='blue-normal' align="center">3.</td>
                                <td class='blue-normal'><img src="http://ais.atlasindustries.com/timesheet/images/word.png" /></td>
                                <td class='blue'><a href='http://ais.atlasindustries.com/timesheet/Data/HR%20Documents/OT Form_TP_2020.docx' class='c'>OVERTIME FORM - THIRD PARTY</a></td>
                                <td  class='blue-normal'>Used for TP staff</td>
                            </tr>
                            <tr bgcolor="#E7EBF5" height="40px" valign='middle'>
                                <td class='blue-normal' align="center">4.</td>
                                <td class='blue-normal'><img src="http://ais.atlasindustries.com/timesheet/images/word.png" /></td>
                                <td class='blue'><a href='http://ais.atlasindustries.com/timesheet/Data/HR%20Documents/Leave Form.docx' class='c'>LEAVE APPLICATION</a></td>
                                <td  class='blue-normal'></td>
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
<%if strLast<>"" then%>
  <tr> 
    <td> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
		  <tr> 
		    <td align="right" bgcolor="#E7EBF5"> 
		      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
		        <tr class="black-normal"> 
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
		          </td>
		          <td align="center" valign="middle" width="13%" class="blue-normal"> 
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageShort")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPageShort")%>/<%=session("NumPageshort")%>&nbsp;&nbsp;</td>
		          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();" onMouseOver="self.status='Go to previous page'; return true;" onMouseOut="self.status=''">Previous</a> /
		          <a href="javascript:next();" onMouseOver="self.status='Go to next page'; return true;" onMouseOut="self.status=''"> Next</a>&nbsp;&nbsp;&nbsp;</td>
		        </tr>
		      </table>
		    </td>
		  </tr>
		</table>
    </td>
  </tr>
<%end if%>
</table>
<%'end of @@content
  Response.Write(arrTmp(1))
%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))
			%>
<input type="hidden" name="txtshortid" value="<%=strShortID%>">
<input type="hidden" name="txtshortname" value="<%=strShortlistname%>">
</form>
</body>
</html>