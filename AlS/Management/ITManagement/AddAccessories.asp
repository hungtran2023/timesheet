<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%

'****************************************
' Function: Outbody
' Description: 
' Parameters: source recordset, number of lines on one page
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc)
	dim intStart,intFinish
	dim strChecked
	strOut = ""

	if not rsSrc.EOF then
		

		For i = 1 to rsSrc.Recordcount
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
			strChecked =""
			if cint(rsSrc("PCAccessoriesID"))<>-1 then strChecked="checked"

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
							"<td valign='top' class='blue'>" & i & "</td>" &_
			         "<td valign='top' class='blue'>" & Showlabel(rsSrc("Description")) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("NoteAccessories")) & "</td>" &_
			         "<td valign='top' class='blue-normal' align='center'><input type='checkbox' name='chkass' "& _ 
			         			" value='" & rsSrc("AccessoriesID") & "'" & strChecked & "></td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function
'------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage, PageSize, fgUpdate, fgRight
	Dim intAtlasPCID,strURLBack,strLicenceType

	strURLBack="AtlasComputer.asp"
	intAtlasPCID=Request.Form("txtAtlasPCID")
	

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
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
	
'-------------------------------
' Get Fullname and Job Title
'-------------------------------
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
'-----------------------------
' Make list of menu
'-----------------------------
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
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
'Get list of data
'--------------------------------------------------
intGroupID=0
strAct=Request.QueryString("act")

	if strAct="SAVE" then
	
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		gMessage = ""
		if ret then
			objDb.cnDatabase.BeginTrans

			objDb.runActionQuery("DELETE ATC_PCAccessories WHERE AtlasPCID=" & intAtlasPCID)

			arrAss=split(Request.Form("chkass"),",")
			
			for ii=0 to UBound(arrAss)			
			
				intLicence= Request.Form("lstLicenceType" & trim(arrAss(ii)))					
				
				strSql="INSERT INTO ATC_PCAccessories (AtlasPCID ,AccessoriesID) VALUES (" & intAtlasPCID & "," & arrAss(ii) & ")"

				ret = objDb.runActionQuery(strSql)	  	 
				if not ret then 
				  gMessage = objDb.strMessage
				  Exit For
				end if
			next
			
			if gMessage<>"" then 
				objDb.cnDatabase.RollbackTrans
			else
				objDb.cnDatabase.CommitTrans
				'gMessage = "Assigned successfully."	  	
			end if
			objDb.dbdisConnect
		else
			gMessage = objDb.strMessage
		end if
		
		set objDb = nothing
	end if

	strSQL="SELECT a.*, ISNULL(b.PCAccessoriesID,-1) as PCAccessoriesID FROM  ATC_ITAccessories a LEFT JOIN " & _
					"(SELECT * FROM ATC_PCAccessories  WHERE AtlasPCID=" & intAtlasPCID & ") b ON a.AccessoriesID=b.AccessoriesID WHERE fgActivate=1"
	 
	Call GetRecordset(strSQL,rsData)
	

	strLast=Outbody(rsData)
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_pro.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intPageCount)	
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../../timesheet.css">

<script language="javascript" src="../../library/menu.js"></script>
<script language="javascript" src="../../library/library.js"></script>
<script>
<!--

function checkedAll (own) {

	var checkboxes
	
	checkboxes = document.getElementsByName('chkass');

	 for(var i=0, n=checkboxes.length;i<n;i++) {
    checkboxes[i].checked = own.checked;
  }
	
}

function BackPrevious() {

	document.navi.action = "<%=strURLBack%>";
	document.navi.target = "_self";
	document.navi.submit();
	
}

function AddAccessories() {

	document.navi.action = "AddAccessories.asp?act=SAVE";
	document.navi.target = "_self";
	document.navi.submit();
	
}



-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="navi" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			Response.Write(arrTmp(0))
			%>
          <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
					<td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
				</tr>
                
                <tr> 
                  <td class="title" height="50" align="center" colspan="5">Choose the Accessories for </td>
                </tr>
                
                 
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="8CA0D1"> 
                            	<td class="blue" align="center" with="10%"  >No.</td>
                              <td class="blue" align="center" with="40%" >Description</td>
                              <td class="blue" align="center" with="40%" >Note</td>
							  <!--<td class="blue" align="center" width="10%">Licence Type<input type='checkbox' name='chkAllLicence' value='1' onclick='checkedAll(this);'></td>-->
	                          <td class="blue" align="center" valign="bottom" width="10%">
								<input type='checkbox' name='chkAllAss' value='1' onclick='checkedAll(this);' ></td>
                            </tr>
                            
<%
	Response.Write(strLast)
%>
                          </table>
						  <table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr> 
			    <td bgcolor="#FFFFFF" height="20" class="blue-normal" align="center"> 
			      <table width="120" border="0" cellspacing="5" cellpadding="0" height="20">
			        <tr> 
			          <td align="center" class="blue" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
			            <a href="javascript: AddAccessories();" class="b" onMouseOver="self.status='Assign'; return true;" onMouseOut="self.status=''">Submit</a>
			          </td>
			          <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
			          <a href="javascript:BackPrevious();" class="b" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
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
			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(1))

			%>		

			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
<input type="hidden" name="txthidden" value="">
<input type="hidden" name="txtAtlasPCID" value="<%=intAtlasPCID%>">
<input type="hidden" name="txtID" value="<%=Request.Form("txtID")%>">
<input type="hidden" name="txtURLBack" value="<%=strURLBack%>">
</form>
<SCRIPT language=JavaScript1.2>
var hotkey=13
if (document.layers)
document.captureEvents(Event.KEYPRESS)
function backhome(e){
	if (document.layers){
		if (e.which==hotkey)
			search();}
	else if (document.all){
		if (event.keyCode==hotkey){
			event.keyCode = 0;
			search();}
	}
}
document.onkeypress=backhome
</SCRIPT>
</body>
<% 
	If gMessage="" and strAct="SAVE" then %>
				<SCRIPT language="javascript">
				<!--				
				document.navi.action = '<%=strURLBack%>';
				document.navi.target = "_self";
				document.navi.submit();
				-->
				</SCRIPT>
<%end if%>
</html>