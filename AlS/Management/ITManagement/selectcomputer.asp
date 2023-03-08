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
function Outbody(ByRef rsSrc, ByVal psize)
Dim strOut
	strOut = ""
	
	'topofpage = (whichpage-1)*psize
	if not rsSrc.EOF then

		cnt = 0
		For i = 1 to rsSrc.Recordcount
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
'Response.Write "works"	& UBound(arrTmp,1)
			strCHK = ""
			'if arrTmp(0, topofpage + i - 1) = 1 then
			'	strCHK = "checked"
			'end if
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
					"<td valign='top' class='blue'>&nbsp;" & Showlabel(rsSrc("PC_Code")) & "</td>" & chr(13) &_
                    "<td valign='top' class='blue-normal'>&nbsp;" & Showlabel(rsSrc("ComputerName")) & "</td>" & chr(13) &_
                    "<td valign='top' class='blue-normal'>&nbsp;" & Showlabel(rsSrc("UserName")) & "</td>" & chr(13) &_
                    "<td valign='top' class='blue-normal'>&nbsp;" & Showlabel(rsSrc("Fullname")) & "</td>" & chr(13) &_
                    "<td valign='top' class='blue-normal' align='center'>" & _
						"<select name='lstLicenceType" & rsSrc("AtlasPCID") & "' class='blue-normal' style='width:95%'>" & _
							strLicenceType & "</select>" & _
                    "</td>" & chr(13) &_
                    "<td valign='top' class='blue-normal' align='center'>"

			strOut = strOut & "<input type='checkbox' name='chkass' value='" & rsSrc("AtlasPCID") & "'" & " " & strCHK & "></td>" & chr(13)
			strOut = strOut & "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	

	Outbody = strOut
end function

'--------------------------------------
Dim gMessage, PageSize
Dim arrlstFrom(2),arrlongmon,strLicenceType

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

'---------------------------------
' Get ruleID
'---------------------------------
intSoftwareID=Request.Form("txtID")
strSoftwareName=Request.Form("txtName")

strURLBack="SoftwareDetail.asp"
strAct=Request.QueryString("act")
	if strAct="SAVE" then
	
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		gMessage = ""
		if ret then
			objDb.cnDatabase.BeginTrans
			
			arrAss=split(Request.Form("chkass"),",")
			for ii=0 to UBound(arrAss)			
				
				intLicence= Request.Form("lstLicenceType" & trim(arrAss(ii)))						
				
				strSql="INSERT INTO ATC_PCSoftware (AtlasPCID ,SoftwareID ,Short_Lived) VALUES (" & arrAss(ii) & "," & intSoftwareID & "," & intLicence & ")"
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

'-------------------------------
' get list of Computer
'-------------------------------
	
	strSql="SELECT * FROM ATC_SoftwareLicenceType WHERE fgActivate=1"
	Call GetRecordset(strSql,rsLicenceType)
	
	if rsLicenceType.recordCount>0 or not rsLicenceType.EOF then
		strLicenceType=""
		rsLicenceType.MoveFirst
		do while not rsLicenceType.EOF
	
			strLicenceType=strLicenceType & "<option value='" & rsLicenceType("LicenceTypeID") & "'>" & rsLicenceType("LicenceTypeDescription") & "</option>" 
			rsLicenceType.MoveNext
		loop
	end if

	strSearch=Request.Form("txtSearch")
	
	strSQL ="SELECT PC_Code,AtlasPCID,ComputerName,ISNULL(UserName,space(1)) as UserName, ISNULL((d.FirstName + space(1) + d.LastName),PublicName) as Fullname  " & _
				"FROM ATC_AtlasPC a INNER JOIN ATC_Computers c ON a.PCID=c.PCID  " & _
				"LEFT JOIN ATC_Users b ON a.UserID=b.UserID " & _
				"LEFT JOIN ATC_PersonalInfo d ON a.UserID=d.PersonID " & _
				"WHERE AtlasPCID NOT IN (SELECT AtlasPCID FROM ATC_PCSoftware WHERE SoftwareID= " & intSoftwareID & ") " 

	if trim(strSearch<>"") then
		intSearchType=Request.Form("lstType")
		if CInt(intSearchType)=1 then 
			strSearch=" PC_Code like '%" & trim(strSearch) & "%'"
		elseif CInt(intSearchType)=2 then
			strSearch=" ComputerName like '%" & trim(strSearch) & "%'"
		else
			strSearch=" Username like '%" & trim(strSearch) & "%'"
		end if
		strSQL=strSQL & " AND " & strSearch 
	end if
	
	strSQL=strSQL & " ORDER BY PC_Code"
		
Call GetRecordset(strSQL,rsData)


strLast=Outbody(rsData,PageSize)

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

function next() {
//var curpage = <%=session("CurPagesee")%>
//var numpage = <%=session("NumPagesee")%>
	if (curpage < numpage) {
		document.selectcomputer.action = "selectcomputer.asp?navi=NEXT";
		document.selectcomputer.target = "_self";
		document.selectcomputer.submit();
	}
}

function prev() {
//var curpage = <%=session("CurPagesee")%>
//var numpage = <%=session("NumPagesee")%>
	if (curpage > 1) {
		document.selectcomputer.action = "selectcomputer.asp?navi=PREV";
		document.selectcomputer.target = "_self";
		document.selectcomputer.submit();
	}
}

function go() {
	//var numpage = <%=session("NumPagesee")%>
	//var curpage = <%=session("CurPagesee")%>
	var intpage = document.selectcomputer.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.selectcomputer.action = "selectcomputer.asp?Go=" + intpage;
		document.selectcomputer.target = "_self";
		document.selectcomputer.submit();	
	}
	else
		alert("Enter another number please.")
}

function sort(type) {
	document.selectcomputer.action = "selectcomputer.asp?sorttype=" + type; //1: fullname, 2: jobtitle
	document.selectcomputer.target = "_self";
	document.selectcomputer.submit();
}

function search() {

	document.selectcomputer.action = "selectcomputer.asp"
	document.selectcomputer.target = "_self";
	document.selectcomputer.submit();

}

function checkedAll (own) {

	var aa= document.getElementById('selectcomputer');
	var chkName
	
	if (own.name=='chkAllLicence')
		chkName="chkLicence"		
	else
		chkName="chkass"
		
	for (var i =0; i < aa.elements.length; i++) 
	{
		strName=String(aa.elements[i].name)
		
		if (aa.elements[i].type == "checkbox" && strName.indexOf(chkName)>-1)
			aa.elements[i].checked = own.checked;
	}
}

function assignment() {

		document.selectcomputer.action = "selectcomputer.asp?act=SAVE";
		document.selectcomputer.target = "_self";
		document.selectcomputer.submit();
}

function BackPrevious(strURL) {

	document.selectcomputer.action = strURL;
	document.selectcomputer.target = "_self";
	document.selectcomputer.submit();
	
}
-->
</script>
</head>


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onunload="return window_onunload();">
<form name="selectcomputer" method="post">
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
                  <td class="blue-normal" align="right" width="15%" valign="middle"> 
                    Search for&nbsp; </td>
                  <td align="right" width="25%" valign="middle"> 
                    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:150" value="<%=Showvalue(varSearch)%>">
                  </td>
                  <td align="right" width="20%" valign="middle"> 
                    <select name='lstType' size='1' height='26px' width='70px' style='width:95%;height=24px;' class='blue-normal'>
						<option value='1'>PC code</option>
						<option value='2'>Computer Name</option>
						<option value='3' selected>User name</option>
					</select>
                    
                  </td>
                  <td class="blue" width="30%" valign="middle"> 
                    <table width="100" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
                      <tr> 
                        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
                            <a href="javascript:search();" class="b">Search</a></td>
                      </tr>
                    </table>
                  </td>
                </tr>
          
          <tr> 
            <td class="title" height="50" align="center" colspan="4"> Select Computer for </td>
          </tr>
          <tr> 
			<td bgcolor="#FFFFFF" valign="top" colspan="4">
					<table width="55%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#003399">
                      <tr> 
                        <td > <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" >
                            <tr> 
                              <td bgcolor="#C0CAE6" >                              
								<table width="100%" border="0" cellspacing="5" cellpadding="0">
                                  <tr> 
                                    <td valign="middle" class="blue-normal" width="30%">&nbsp;&nbsp;Software name</td>
                                    <td valign="middle" width="70%" class="blue"> <%=strSoftwareName%> 
                                    </td>
                                  </tr>              
                                </table>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                    </table> </td>
                </tr>
        </table>
      </td>
    </tr>
    <tr> 
            <td align="center" colspan="4"> &nbsp;&nbsp;</td>
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
	                  <td class="blue" align="center" bgcolor="8CA0D1" width="15%">PC Code</td>
	                  <td class="blue" align="center" width="20%">Computer name</td>
					<td class="blue" align="center" width="15%">User name</td>
					<td class="blue" align="center" width="30%">Fullname</td>	
	                  <td class="blue" align="center" width="15%">Licence Type</td>
	                  <td class="blue" align="center" valign="bottom" width="5%"><input type='checkbox' name='chkAllAss' value='1' onclick='checkedAll(this);' ></td>
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
	<input type="hidden" name="txtkind" value="<%=kindact%>">
	<input type="hidden" name="txtURLBack" value="<%=strURLBack%>">
	<input type="hidden" name="txtName" value="<%=strSoftwareName%>">
	<input type="hidden" name="txtID" value="<%=intSoftwareID%>">
	
  <%if kindact<=2 then%>
	
	
	<input type="hidden" name="txthiddenstrproName" value="<%=Request.Form("txthiddenstrproName")%>">
	<input type="hidden" name="txthiddenstrproID" value="<%=Request.Form("txthiddenstrproID")%>">
	<input type="hidden" name="txtpreviouspage" value="<%=Request.Form("txtpreviouspage")%>">
	
  <%end if%>
  

		      </table>
</form>
</body>
<% 
	If gMessage="" and strAct="SAVE" then %>
				<SCRIPT language="javascript">
				<!--				
				document.selectcomputer.action = '<%=strURLBack%>';
				document.selectcomputer.target = "_self";
				document.selectcomputer.submit();
				-->
				</SCRIPT>
<%end if%>
</html>

 