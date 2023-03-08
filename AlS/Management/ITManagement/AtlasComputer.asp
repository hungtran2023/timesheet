<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim strPCCode,strComputerName,intTypeOfPC,strIPAddress,intUserIDAssign,strUserAssign,strPublicUser
	dim intPCID,intAtlasPCID
	Dim objEmployee, objDatabase, strError,rsData
	Dim arrlstFrom(2),arrlongmon

'***************************************************************
'
'***************************************************************
function OutBody(rsSrc)
	dim strOut
	dim i
	
	strOut=""
	i=0
	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"
			
			strLicence="<img src='../../images/" & rsSrc("Short_Lived")-1 & ".gif' border=0>"
			if cint(rsSrc("Short_lived"))=1 then strLicence="<a href='javascript:UpdateLincence(" & rsSrc("PCSoftwareID") & ")'>" & strLicence & "</a>"
			
			strOut=strOut & "<tr bgcolor='" & strColor & "'>" 
			strOut=strOut & "<td valign='top' class='blue-normal'>&nbsp;" & i+1 & "</td>" 
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("SoftwareName") & "</td>" 
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("Description") & "</td>" 
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("Vendor") & "</td>" 
			strOut=strOut & "<td valign='top' class='blue-normal' align='center'>" & strLicence & "</td>" 
			strOut=strOut & "<td valign='top' class='blue-normal' align='center'><input type='checkbox' name='chkRemove' value='" & rsSrc("PCSoftwareID") & "'></td>" 		
			strOut=strOut & "</tr>"
			i=i+1
			
			rsSrc.MoveNext
		loop
		
	end if
	
	OutBody=strOut
End Function

'***************************************************************
'
'***************************************************************

function GetTypeOfListBox(rsSrc,intTypeID)
	dim strOut
	
	strOut=""
	
	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
									
			strSelect=""
			if cint(rsSrc("AtlasPCTypeID")) =cint(intTypeID) then strSelect="selected"
			
			strOut=strOut & "<option value='" & rsSrc("AtlasPCTypeID") & "' " & strselect & " >" & rsSrc("Description")  & "</option>"
			rsSrc.MoveNext
		loop
		
	end if

	GetTypeOfListBox=strOut
end function


'***************************************************************
'
'***************************************************************
function ConnectAComputerToNetwork(strComputerName,intTypeOfPC,	strIPAddress,intPCID, intUserIDAssign,strPublicName)
	
'Response.Write "ComputerName: " & strComputerName & "<br>"
'Response.Write "intTypeOfPC: " & intTypeOfPC & "<br>"
'Response.Write "strIPAddress: " & strIPAddress & "<br>"
'Response.Write "intPCID: " & intPCID & "<br>"
'Response.Write "intUserIDAssign: " & intUserIDAssign & "<br>"
'Response.Write "strPublicName: " & strPublicName & "<br>"

'Response.end
 


	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "ConnectAComputerToAtlasNetwork"
	
		Set myParam = myCmd.CreateParameter("computername", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("typeOfPC",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("IPAddress", adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam			
		Set myParam = myCmd.CreateParameter("PCID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam				
		Set myParam = myCmd.CreateParameter("userID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("PublicName", adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("AtlasPCID", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam

		myCmd("computername")	= strComputerName
		myCmd("typeOfPC")		= intTypeOfPC
		myCmd("IPAddress")		= strIPAddress
		myCmd("PCID")			= intPCID
		myCmd("userID")			= intUserIDAssign
		myCmd("PublicName")		= strPublicName
		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
		Else
			intAtlasPCID=myCmd("AtlasPCID")
	
			If cint(intAtlasPCID)<>-2 Then
				strError = "Assign computer successfull"
			else
				strError = "The computer name is not avalable for using."
				intAtlasPCID=-1
			end if
			
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	ConnectAComputerToNetwork=strError
	
end function

'***************************************************************
'
'***************************************************************
function UpdateAssignment(strComputerName,intTypeOfPC,	strIPAddress, intUserIDAssign, strPublicName,intAtlasPCID)
	dim strErr, strSql 
	
		strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "UpdateComputerInAtlasNetwork"
	
		Set myParam = myCmd.CreateParameter("computername", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("typeOfPC",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("IPAddress", adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam			
		Set myParam = myCmd.CreateParameter("userID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("PublicName", adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam			
		Set myParam = myCmd.CreateParameter("AtlasPCID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("Err", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam

		myCmd("computername")	= strComputerName
		myCmd("typeOfPC")		= intTypeOfPC
		myCmd("IPAddress")		= strIPAddress
		myCmd("userID")			= intUserIDAssign
		myCmd("PublicName")		= strPublicName
		myCmd("AtlasPCID")		=intAtlasPCID

		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
		Else
			intErr=myCmd("Err")
	
			If cint(intErr)=0 Then
				strError = "Update computer successfull"
			else
				strError = "The computer name is not avalable for using."
			end if
						
			
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	UpdateAssignment=strError
end function
'***************************************************************
'
'***************************************************************
function RevomeComputerOutNetwork(intAtlasPCID)
	dim strErr, strSql ,rsTest
	
	strCnn = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	
	If objDatabase.dbConnect(strCnn) Then
	
        strSql="SELECT Count(*) as numOfSorftware FROM ATC_PCSoftware WHERE AtlasPCID =" &  intAtlasPCID
        
        Call GetRecordset(strSql,rsTest)
        
        if rsTest("numOfSorftware")=0 then
            
		    objDatabase.runActionQuery("DELETE FROM ATC_AtlasPC WHERE AtlasPCID =" &  intAtlasPCID)
    		
		    If Err.number > 0 Then
			    strError= Err.Description
		    Else
			    objDatabase.runActionQuery("DELETE FROM ATC_AtlasPC WHERE AtlasPCID =" &  intAtlasPCID)
			    strError = "This computer has been removed out Atlas network."
    			
			    intAtlasPCID=-1
		    End If
		    Err.Clear
		 else
		    strError = "Please renove all software install before disjoin network."
		 end if
    else
	    strError=objDatabase.strMessage
    end if
	    
	set objDatabase=nothing	
	
	RevomeComputerOutNetwork=strError
end function

'***************************************************************
'
'***************************************************************
function ExecuteSQL(strSql)

	dim objDatabase
	dim strCnn
	dim blnReturn
	
	blnReturn = false	
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	
	If objDatabase.dbConnect(strCnn) then		
		blnReturn= (objDatabase.runActionQuery(strSql))	
		strError="Update successfull."
		if not blnReturn then strError=objDatabase.strMessage		
	else
		strError=objDatabase.strMessage
	end if
	
	Set objDatabase = nothing
	ExecuteSQL=strError
	
end function

'***************************************************************
'
'***************************************************************
function GetAccessories()
	dim strAccess

	strAccess=""

	strSQL="SELECT * FROM [ATC_PCAccessories] a inner join [ATC_ITAccessories] b ON a.AccessoriesID=b.AccessoriesID WHERE AtlasPCID=" & intAtlasPCID
	Call GetRecordset(strSQL,rsAccessories)
	if not rsAccessories.EOF then

		do while not rsAccessories.EOF
				
				if strAccess="" then 
					strAccess=rsAccessories("Description")
				else
					strAccess=strAccess & ", " & rsAccessories("Description")
				end if

				rsAccessories.Movenext
		loop

	end if

GetAccessories=strAccess
end function
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If

	intUserID = session("USERID")

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	'strConnect = Application("g_strConnect")
	'Set objDatabase = New clsDatabase
	intPCID = Request.Form("txtID")
	intAtlasPCID=Request.Form("txtAtlasPCID")
	fgDel=Request.Form("fgstatus")

	if Request.QueryString("act") = "save" then


		strComputerName=Request.Form("txtComputerName")
		intTypeOfPC= Request.Form("lbType")
		strIPAddress=Request.Form("txtIP")

		if trim(strIPAddress)="" then strIPAddress=null	
		
		if Request.Form("txtUserID")="" then
			intUserIDAssign=null		
		else
			intUserIDAssign=Request.Form("txtUserID")
			if cint(intUserIDAssign)=0 then intUserIDAssign=null
		end if
		
		intTypeUser=Request.Form("radUser")
				
		if cint(intTypeUser)=0 then
			intUserIDAssign=null
			strPublicUser=Request.Form("txtOtherUser")
			if strPublicUser="" then strPublicUser=null
		else
			strPublicUser=null
		end if

		if fgDel<>"D" then
						
				
				if Cint(intAtlasPCID)=-1 then
					'Add new	

					strError=ConnectAComputerToNetwork(strComputerName,intTypeOfPC,strIPAddress ,intPCID , intUserIDAssign,strPublicUser )
				else
					intAtlasPCID=Request.Form("txtAtlasPCID")
					'Update
					strError=UpdateAssignment(strComputerName,intTypeOfPC,strIPAddress ,intUserIDAssign,strPublicUser,intAtlasPCID)
				end if
		else
			strError=RevomeComputerOutNetwork(intAtlasPCID)
		end if	
	elseif 	Request.QueryString("act") = "remove" then
		
		arrComputer=Request.Form("chkRemove")
		
		if trim(arrComputer)<>"" then
			strSql="DELETE FROM ATC_PCSoftware WHERE PCSoftwareID IN (" & arrComputer & ")"
			strError= ExecuteSQL(strSql)
		end if
		
	End If
'--------------------------------------------------
' 
'--------------------------------------------------
	
	strSql="SELECT ISnull(AtlasPCID,-1) as AtlasPCID,ComputerName,PublicName,isnull(TypeOfPC,0) as TypeOfPC,IP_Address,a.PCID,isnull(b.UserID,0) as UserID, a.PC_Code , c.UserName, ISNULL(d.UserType,0) as UserType " & _
				"FROM ATC_Computers a LEFT JOIN ATC_AtlasPC b ON a.PCID=b.PCID " & _
					"LEFT JOIN ATC_Users c ON c.UserID=b.UserID " & _
					"LEFT JOIN ATC_PersonalInfo d ON c.UserID=PersonID " & _
					"WHERE AtlasPCID=" & intAtlasPCID
	
	Call GetRecordset(strSql,rsData)	
	
	if rsData.RecordCount>0 then
		strPCCode=rsData("PC_Code")
		strComputerName=rsData("ComputerName")	
		intTypeOfPC =rsData("TypeOfPC")
		strIPAddress=rsData("IP_Address")
	
		intUserIDAssign=rsData("UserID")
		strUserAssign= rsData("UserName")
		intTypeUser=rsData("UserType")
		
		strPublicUser=rsData("PublicName")
		if strPublicUser<>"" then intTypeUser=0
		
		intPCID=rsData("PCID")
		intAtlasPCID=rsData("AtlasPCID")
		
	end if
	

	if Request.QueryString("act") = "out" then
	
		intUserIDAssign=Request.Form("txtUserID")	
		
		
		Call GetRecordset("SELECT UserName FROM ATC_Users WHERE UserID=" & intUserIDAssign,rsData)				
		if not rsData.EOF then strUserAssign=rsData("UserName")
		
		strComputerName=Request.Form("txtComputerName")
					
		intTypeOfPC =Request.Form("lbType")
		if intTypeOfPC="" then intTypeOfPC=0
		
		strIPAddress=Request.Form("txtIP")		
		intAtlasPCID=Request.Form("txtAtlasPCID")
		
		intUserIDAssign=Request.Form("txtUserID")
		if cint(intUserIDAssign)=0 then intUserIDAssign=null
				
		intTypeUser=Request.Form("radUser")
				
		strPublicUser=null	
		
	end if
		
	
	strSql="SELECT * FROM ATC_ComputerType WHERE fgActivate=1"
	Call GetRecordset(strSql,rsComputerType)
	
	strTypeList=GetTypeOfListBox(rsComputerType,intTypeOfPC)
	
	strSql= "SELECT PCSoftwareID,softwarename,short_Lived, vendor,d.Description FROM ATC_AtlasPC a " & _
				"INNER JOIN ATC_PCSoftware b ON a.AtlasPCID=b.AtlasPCID " & _
				"INNER JOIN ATC_Softwares c ON b.SoftwareID=c.SoftwareID " & _
				"LEFT JOIN ATC_SoftwareType d ON d.SoftTypeID=c.SoftTypeID " & _
			"WHERE b.AtlasPCID= " & intAtlasPCID


	Call GetRecordset(strSql,rsData)
	strLast=OutBody(rsData)

'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)					
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
	
	'current URL
	If Request.ServerVariables("QUERY_STRING")<>"" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
	strChoseMenu = Request.QueryString("choose_menu")
	If strChoseMenu = "" Then strChoseMenu = "AF"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

	arrlstFrom(0) = selectmonth("lstmonthF",month(dateBuying) , -1)
	arrlstFrom(1) = selectday("lstdayF", day(dateBuying), -1)
	arrlstFrom(2) = selectyear("lstyearF", year(dateBuying), 1999, year(date())+2, 0)
		
			
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intCurPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intTotalPage)	
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">


</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
	      <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%		If strError <> "" Then%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red" colspan="2">&nbsp;<b><%=strError%></b></td>
				</tr>
<%		End If%>				
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;<a href="ComputerList.asp" onMouseOver="self.status='';return true">Computer List</a></td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;</td>
                  <td class="blue" height="30" align="right" width="77%">
					<table width="360" border="0" cellspacing="2" cellpadding="0" align="right" height="20" name="aa">
                      <tr> 
                       
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue"><a href="javascript:ComputerDetail()" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" class="b">Computer Detail </a></div>
                        </td>
						<td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue">
<%if cint(intAtlasPCID)<>-1 then%> 
<td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue">             
	<div align="center" class="blue">            
							<a href="javascript:InstallSoftwares()" onMouseOver="self.status='Please click here to view Expired rules.';return true" onMouseOut="self.status='';return true" class="b">	Add Softwares</a></div></div></td>
<td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue">							
							<div align="center" class="blue">            
							<a href="javascript:AddAccessories()" onMouseOver="self.status='Please click here to view Expired rules.';return true" onMouseOut="self.status='';return true" class="b">	Add Accessories</a></div></td>
<%end if%>
                                                
                      </tr>
                    </table>
                  </td>
                </tr>                
                <tr align="center"> 
                  <td class="title" height="30" align="center" colspan="2"></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="2">
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="25%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="20%"></td>
                              <td valign="middle" width="35%" class="blue">
								<%=strPCCode%></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            

                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Computer name *</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtComputerName" maxlength="20" class="blue-normal" size="20" style="width:95%" value="<%=strComputerName%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                             <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Type *</td>
                              <td valign="middle" class="blue">
                              <select name='lbType' size='1' class='blue-normal' style="width:95%">
                              <option value='0'>&nbsp;</option>
								<%=strTypeList%></select>
								
							  </td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">IP Address</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtIP" maxlength="50" class="blue-normal" size="20" style="width:95%" value="<%=strIPAddress%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">User Assignment </td>
                              <td valign="middle" class="blue"><table width="40%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td><input name="radUser" type="radio" value="1" <%if cint(intTypeUser)=1 then%>checked="checked"<%end if%> onclick="javascript:showhide('AtlasUser',1);showhide('OtherUser',0);"></td>
                                  <td class="blue-normal">Atlas User</td>
                                  <td><input name="radUser" type="radio" value="3" <%if cint(intTypeUser)=3 then%>checked="checked"<%end if%> onclick="javascript:showhide('AtlasUser',1);showhide('OtherUser',0);"></td>
                                  <td class="blue-normal">Contract User</td>
                                 
                                </tr>
                              </table></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal"> </td>
                              <td valign="middle" class="blue">
								<div id="AtlasUser" style="display:block;">
									<table width="95%" border="0" cellspacing="0" cellpadding="0">
										<tr><td width="60%">
											<table width="100%" border='0' bordercolor='#333333'>
												<tr><td bgcolor='#333333'>
													<table width="100%" border='0' cellspacing='1'>
														<tr><td class='blue-normal' bgcolor='#FFFFFF'>&nbsp;<%=strUserAssign%></td></tr>
													</table></td></tr>
											</table>
										</td>
										<td width="40%">&nbsp;&nbsp;<a href="javascript:selectuser()">Select User... </a></td>
										</tr>                           
									</table> </div> 
								                           
                              </td>
                              <td valign="top" class="blue-normal" align="center"></td>

                            </tr>            
<tr bgcolor="#FFFFFF">
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal"></td>
                              <td valign="middle" class="blue"><table border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                   <td><input name="radUser" type="radio" value="0" <%if cint(intTypeUser)=0 then%>checked="checked"<%end if%> onclick="javascript:showhide('OtherUser',1);showhide('AtlasUser',0);"></td>
                                  <td class="blue-normal">Other</td>
                                 
                                </tr>
                              </table></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td></tr>
<tr bgcolor="#FFFFFF">
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal"> </td>
                              <td valign="middle" class="blue">
								
								 <div id="OtherUser" style="display:block;">

									<input type="text" name="txtOtherUser" maxlength="50" class="blue-normal" size="20" style="width:95%" value="<%=strPublicUser%>"></div>                           
                              </td>
                              <td valign="top" class="blue-normal" align="center">

                              	<input type="hidden" name="txtUserID" value="<%=intUserIDAssign%>"></td>
                            </tr>  
<tr bgcolor="#FFFFFF">
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal"> </td>
                              <td valign="middle" class="blue"> 
																	       <input type="text" name="txtAccessories" class="blue-normal" size="20" style="width:95%" value="<%=GetAccessories()%>">                  
                              </td>
                              <td valign="top" class="blue-normal" align="center"></td>
                            </tr> 

                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table width="150" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" >
                                      <div align="center" class="blue"><a href="javascript:savedata()" onMouseOver="self.status='Please click here to save changes';return true" onMouseOut="self.status='';return true" class="b"><%if cint(intAtlasPCID)=-1 then%>Join Network<%else%>Update<%end if%></a></div>
                                    </td>
<%if cint(intAtlasPCID)<>-1 then%>  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" >
                                      <div align="center" class="blue"><a href="javascript:deletedata()" onMouseOver="self.status='Please click here to delete this record';return true" onMouseOut="self.status='';return true" class="b">Disjoin Network</a></div>
                                    </td>
<%end if%>                                    
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1">
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="5%">No.</td>
                              <td class="blue" align="center" width="30%">Software</td>  
                              <td class="blue" align="center" width="20%">Type</td>  
                              <td class="blue" align="center" width="20%">Vendor</td>  
                              <td class="blue" align="center" width="10%">Licence</td>
                              <td class="blue" align="center" width="15%"><input type='checkbox' name='chkAll' value='1' onclick='checkedAll(this);' ></td>
                            </tr>
<%Response.Write strLast%>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue" align="right"><a href="javascript:uninstallSoftware()">Remove Software</a>&nbsp;&nbsp;</td>
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
<input type="hidden" name="fgstatus" value="<%=fgDel%>">
<input type="hidden" name="txtID" value="<%=intPCID%>">
<input type="hidden" name="txtAtlasPCID" value="<%=intAtlasPCID%>">
<input type="hidden" name="txtURLBack" value="<%=strLink%>">
<input type="hidden" name="txtPCSoftwareID" value="">

</form>
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--

function ComputerDetail()
{
		window.document.frmreport.action = "ComputerDetail.asp?act=EDIT"
		window.document.frmreport.submit();
}

function selectuser(){
        
		window.document.frmreport.action = "selectemployee_ass.asp"
		window.document.frmreport.submit();
}


function savedata()
{
	if (checkdata())
	{
		window.document.frmreport.action = "AtlasComputer.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	var answer = confirm("Do you want to remove <%=strComputerName%> out Atlas Network?")
	if (answer){
		window.document.frmreport.fgstatus.value = "D"
		window.document.frmreport.action = "AtlasComputer.asp?act=save"			
		window.document.frmreport.submit();
	}	
}

function checkdata()
{
	if (window.document.frmreport.txtComputerName.value=="")
	{
		alert("Please enter Computer Name.");
		document.frmreport.txtComputerName.focus();
		return false	
	}	
	
	if (window.document.frmreport.lbType.value==0)
	{
		alert("Please select type of computer.");
		document.frmreport.lbType.focus();
		return false	
	}

	return true	
}

/*function showhide(layer_ref,val) { 
	var state = 'none'; 

	if (val == 0) { 
		state = 'none'; 
	} 
	else { 
		state = 'block'; 
	} 
	
	if (document.all) { //IS IE 4 or 5 (or 6 beta) 
		eval( "document.all." + layer_ref + ".style.display = state"); 
	} 
	if (document.layers) { //IS NETSCAPE 4 or below 
		document.layers[layer_ref].display = state; 
	} 
	if (document.getElementbyId &&!document.all) { 
		hza = document.getElementbyId(layer_ref); 
		hza.style.display = state; 
	}	

	alert(layer_ref);
} */

function showhide(layer_ref,val) { 
	var state = 'none'; 

	if (val == 0) { 
		state = 'none'; 
	} 
	else { 
		state = 'block'; 
	} 
	
	if (document.getElementbyId &&!document.all) { 
		hza = document.getElementbyId(layer_ref); 
		hza.style.display = "none"; 
	}	
	

	//alert(layer_ref + "--" + val + "--" + state);

} 

	
function checkedAll (own) {

	 var checkboxes = document.getElementsByName('chkRemove');
	 for(var i=0, n=checkboxes.length;i<n;i++) {
    checkboxes[i].checked = own.checked;
  }
}

function InstallSoftwares()
{
	window.document.frmreport.action = "InstallSoftwares.asp"
	window.document.frmreport.submit();
}

function AddAccessories()
{
	window.document.frmreport.action = "AddAccessories.asp"
	window.document.frmreport.submit();
}

function uninstallSoftware()
{
	window.document.frmreport.action = "AtlasComputer.asp?act=remove"			
	window.document.frmreport.submit();
}

function UpdateLincence(id)
{
	window.document.frmreport.txtPCSoftwareID.value=id
	window.document.frmreport.action = "LicenceSoftware.asp?fr=com"		
	window.document.frmreport.submit();
}

//-->
</script>
</body>
</html>
