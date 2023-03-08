<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim intUserID, varPre, getRes, strURL, intMonth, intYear

'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	intMonth = Month(Date)
	intYear = Year(Date)

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID	= session("USERID")

'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strFunction = "<a class='c' href='javascript:gopage();' onMouseOver='self.status=&quot;Preferences&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Preferences</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#'>Help</a>&nbsp;&nbsp;&nbsp;"
	Set objEmployee = Nothing
	
'--------------------------------------------------
' Make list of menu
'--------------------------------------------------

	If isEmpty(session("Menu")) Then 
		getRes = getarrMenu(intUserID)
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End If	

'--------------------------------------------------
' Get current URL
'--------------------------------------------------
	
	If Request.ServerVariables("QUERY_STRING") <> "" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
'--------------------------------------------------
' Get current menu that user is choosing
'--------------------------------------------------
	
	strChoseMenu = Request.QueryString("choose_menu")
	If strChoseMenu = "" Then strChoseMenu = "B"

	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))

	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1) <> "" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if

%>
<html>
<head>
<title>Atlas industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../timesheet.css">

<script language="javascript" src="../../library/library.js"></script>
<script language="javascript">
<!--
var ns, ie;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	var url;
	url = "../../logout.asp";
	if (ns)
		document.location = url;
	else
	{
		window.document.frmreport.action = url;
		window.document.frmreport.submit();
	}	
}

function gopage()
{
	document.frmreport.action = "../../tools/preferences.asp";
	document.frmreport.submit();
}

function checkdata()
{
	if (document.frmreport.rdotype[0].checked)
	{
		if (isnull(document.frmreport.txtFrom.value)==true)
		{
			alert("Please enter startdate before click here.")
			document.frmreport.txtFrom.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtFrom.value)==false)
			{			
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtFrom.focus();
				return false;
			}
		}
		
		if (isnull(document.frmreport.txtTo.value)==true)
		{
			alert("Please enter enddate before click here.")
			document.frmreport.txtTo.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtTo.value)==false)
			{
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtTo.focus();
				return false;
			}
		}
		
		if (comparedate(document.frmreport.txtFrom.value,document.frmreport.txtTo.value)==false)
		{
			alert("The startdate must be less than the finishdate.")
			document.frmreport.txtFrom.focus();
			return false;
		}
	}	
	return true;
}

function viewtms()
{
	if (checkdata() == true)
	{
		document.frmreport.action = "rpt_invalid_tms.asp";
		document.frmreport.submit();
	}	
}

function document_onkeypress() 
{
var keycode = event.keyCode;
	if (keycode == 13) 
	{
		event.keyCode = 0;
		viewtms();
	}
}

//-->
</script>

<script LANGUAGE="javascript" FOR="document" EVENT="onkeypress">
<!--
 document_onkeypress()
//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
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
	Response.Write(arrTmp(0))
%>		

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td class="blue" height="10" align="left" width="37%">&nbsp; </td>
          <td class="blue" height="30" align="right" width="63%">&nbsp;</td>
        </tr>
        <tr> 
          <td class="title" height="50" align="center" colspan="2"> Invalid 
            Timesheet</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign="top">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365" >
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#617DC0"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr bgcolor="#FFFFFF"> 
                      <td valign="top" width="16%" class="blue">&nbsp;</td>
                      <td valign="top" width="21%" class="blue">&nbsp;</td>
                      <td valign="bottom" class="blue" width="17%">Select<b> 
                        </b>Date(s)</td>
                      <td valign="bottom" width="46%" class="blue-normal">&nbsp; 
                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td valign="top" width="16%" class="blue">&nbsp;</td>
                      <td valign="top" width="21%" class="blue">&nbsp;</td>
                      <td valign="bottom" class="blue-normal" width="17%">From 
                      </td>
                      <td valign="bottom" width="46%" class="blue-normal"> 
                        To </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td valign="top" width="16%" class="blue">&nbsp;</td>
                      <td valign="top" width="21%" class="blue" align="right"> 
                        <input type="radio" name="rdotype" value="D" language="javascript" onClick="document.frmreport.txtFrom.focus()">
                      </td>
                      <td valign="top" class="blue-normal" width="17%"> 
                        <input type="text" name="txtFrom" size="10" class="blue-normal" language="javascript" onClick="document.frmreport.rdotype[0].checked=true">
                      </td>
                      <td valign="top" width="46%" class="blue-normal"> 
                        <input type="text" name="txtTo" size="10" class="blue-normal" language="javascript" onClick="document.frmreport.rdotype[0].checked=true">
                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td valign="top" width="16%" class="blue">&nbsp;</td>
                      <td width="21%" class="blue">&nbsp;</td>
                      <td valign="bottom" class="blue-normal" width="17%">Month</td>
                      <td valign="bottom" width="46%" class="blue-normal">Year </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="16%" class="blue">&nbsp;</td>
                      <td align="right" width="21%" class="blue"> 
                        <input type="radio" name="rdotype" value="M" checked language="javascript" onClick="document.frmreport.lbmonth.focus()">
                      </td>
                      <td class="blue-normal" width="17%"> 
						<select name="lbmonth" size="1" class="blue-normal" language="javascript" onFocus="document.frmreport.rdotype[1].checked=true">
						  <option <%If CInt(intMonth)=1 Then%>selected<%End If%> value="1">January</option>
						  <option <%If CInt(intMonth)=2 Then%>selected<%End If%> value="2">February</option>
						  <option <%If CInt(intMonth)=3 Then%>selected<%End If%> value="3">March</option>
						  <option <%If CInt(intMonth)=4 Then%>selected<%End If%> value="4">April</option>
						  <option <%If CInt(intMonth)=5 Then%>selected<%End If%> value="5">May</option>
						  <option <%If CInt(intMonth)=6 Then%>selected<%End If%> value="6">June</option>
						  <option <%If CInt(intMonth)=7 Then%>selected<%End If%> value="7">July</option>
						  <option <%If CInt(intMonth)=8 Then%>selected<%End If%> value="8">August</option>
						  <option <%If CInt(intMonth)=9 Then%>selected<%End If%> value="9">September</option>
						  <option <%If CInt(intMonth)=10 Then%>selected<%End If%> value="10">October</option>
						  <option <%If CInt(intMonth)=11 Then%>selected<%End If%> value="11">November</option>
						  <option <%If CInt(intMonth)=12 Then%>selected<%End If%> value="12">December</option>
						</select>
                      </td>
                      <td width="46%" class="blue-normal"> 
					    <select name="lbyear" size="1" class="blue-normal" language="javascript" onFocus="document.frmreport.rdotype[1].checked=true">
						<%For ii=Year(Date)-1 To Year(Date)%>
					      <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
						<%Next%>
						</select>
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                    <tr> 
                      <td height="50"> 
                        <table width="60" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                          <tr> 
                            <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" width="59" height="20" > 
                              <div align="center" class="blue"> 
                                <a href="javascript:viewtms();" class="b" onMouseOver="self.status='Clich here to view invalid timesheet';return true" onMouseOut="self.status='';return true">Submit</a> 
                              </div>
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
</table>   

<%
	Response.Write(arrTmp(1))
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))
%>
    		
</form>
</body>
</html>
