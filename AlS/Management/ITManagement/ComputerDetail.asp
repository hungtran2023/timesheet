<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim strPCCode,strSupplier,	dateBuying,strNote,strCPU,strRAM,strHDD,strVideo,intStatus, strSN
	dim strPCCodeOld
	dim intPCID,intDetailPCID
	Dim objEmployee, objDatabase, strError,rsData
	Dim arrlstFrom(2),arrlongmon, strArrayDisk1,strArrayDisk2
	Dim arrDisk1, arrDisk2

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
			
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & i+1 & "</td>"
			strOut=strOut & "<td valign='top' class='blue'>" & _
						"<a href='javascript:AtlasNetwork(" & rsSrc("AtlasPCID") & ");' " &_
						"class='c'>" & rsSrc("computername") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("description") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("UserAssignment") & "</td>"		
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
function ExecuteSQL(strSql)

	dim strConnect,ret,strMessage
	dim objDb	

	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
		
	If objDb.dbConnect(strConnect) then
			
		ret = objDb.runActionQuery(strSql)
				
		if ret=false then				
			strMessage = objDb.strMessage
		else
			strMessage="Update successfully."
		end if
			  
	else
		strMessage=objDb.strMessage
	end if
	
	ExecuteSQL=strMessage
end function

'***************************************************************
'
'***************************************************************
sub TrackingPCCode(strAction)
	dim strsql
	strsql="INSERT INTO [ATC_ComputerCodeTracking] ([StaffID],[Action],[PC_codeOld],[PC_codeNew]) " & _
			" VALUES( " & intUserID & ",'" & strAction & "','" & strPCCodeOld & "','" & strPCCode & "')"
	
	strError=ExecuteSQL(strsql)
	
end sub

'***************************************************************
'
'***************************************************************
function AddComputer(strPCCode,strSupplier,	dateBuying,blnoutdated,strNote,strCPU,strRAM,strHDD,strVideo,strArrayDisk1,strArrayDisk2,strSN)

On Error Resume next	

	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "InsertAComputer"
		
		Set myParam1 = myCmd.CreateParameter("PC_code", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam1		
		Set myParam2 = myCmd.CreateParameter("Supplier", adVarChar,adParamInput,200)
		myCmd.Parameters.Append myParam2
		Set myParam3 = myCmd.CreateParameter("BuyDate", adDate,adParamInput)
		myCmd.Parameters.Append myParam3
		Set myParam = myCmd.CreateParameter("outdated", adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam4 = myCmd.CreateParameter("ComputerNote", adVarChar,adParamInput,300)
		myCmd.Parameters.Append myParam4		
		Set myParam5 = myCmd.CreateParameter("CPUInformation",adVarChar,adParamInput,100)
		myCmd.Parameters.Append myParam5
		Set myParam6 = myCmd.CreateParameter("RAM",adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam6
		Set myParam7 = myCmd.CreateParameter("SystemMemory",adVarChar,adParamInput,150)
		myCmd.Parameters.Append myParam7
		Set myParam8 = myCmd.CreateParameter("VideoAdapter", adVarChar,adParamInput,150)
		myCmd.Parameters.Append myParam8
		
		Set myParam = myCmd.CreateParameter("Array1", adVarChar,adParamInput,400)
		myCmd.Parameters.Append myParam
		
		Set myParam = myCmd.CreateParameter("Array2", adVarChar,adParamInput,400)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("SeriNum", adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam
		
		Set myParam9 = myCmd.CreateParameter("PCID", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam9
		Set myParam = myCmd.CreateParameter("DetailPCID", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam

		myCmd("PC_code")		= strPCCode
		myCmd("Supplier")		= strSupplier
		myCmd("BuyDate")		= dateBuying
		myCmd("outdated")		= blnoutdated
		myCmd("ComputerNote")	= strNote
		myCmd("CPUInformation")	= strCPU
		myCmd("RAM")			= strRAM
		myCmd("SystemMemory")	= strHDD
		myCmd("VideoAdapter")	= strVideo
		myCmd("Array1")	= strArrayDisk1
		myCmd("Array2")	= strArrayDisk2
		myCmd("SeriNum")= strSN
		
		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
		Else
			strError = "New computer adding successfull"
			intPCID=myCmd("PCID")
			intDetailPCID=myCmd("DetailPCID")
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	AddComputer=strError
	
end function


'***************************************************************
'
'***************************************************************
function UpdateComputer(strPCCode,strSupplier,dateBuying,blnoutdated, strNote,strCPU,strRAM,strHDD,strVideo,strArrayDisk1,strArrayDisk2,intPCID,intDetailPCID)
		
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "UpdateAComputer"
 
		Set myParam = myCmd.CreateParameter("PC_code", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("Supplier", adVarChar,adParamInput,200)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("BuyDate", adDate,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("outdated", adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("ComputerNote", adVarChar,adParamInput,300)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("CPUInformation",adVarChar,adParamInput,100)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("RAM",adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("SystemMemory",adVarChar,adParamInput,150)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("VideoAdapter", adVarChar,adParamInput,150)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("Array1", adVarChar,adParamInput,400)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("Array2", adVarChar,adParamInput,400)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("SeriNum",adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam	
			
		Set myParam = myCmd.CreateParameter("PCID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("DetailPCID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam

		myCmd("PC_code")		= strPCCode
		myCmd("Supplier")		= strSupplier
		myCmd("BuyDate")		= dateBuying
		myCmd("outdated")		= blnoutdated
		myCmd("ComputerNote")	= strNote
		myCmd("CPUInformation")	= strCPU
		myCmd("RAM")			= strRAM
		myCmd("SystemMemory")	= strHDD
		myCmd("VideoAdapter")	= strVideo
		myCmd("SeriNum")= strSN
		
		myCmd("Array1")	= strArrayDisk1
		myCmd("Array2")	= strArrayDisk2
		
		myCmd("PCID")			= intPCID
		myCmd("DetailPCID")		= intDetailPCID
		

		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
		Else
			strError = "Update successfull"
			if strPCCode<>strPCCodeOld then call TrackingPCCode("UPD")
			strPCCodeOld=strPCCode
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	UpdateComputer=strError
	
end function

'***************************************************************
'
'***************************************************************
function DelComputer(intPCID)
		
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "DeleteAComputer"
	
		Set myParam = myCmd.CreateParameter("PCID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("count", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam

		myCmd("PCID")	= intPCID
		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
		Else
			IF myCmd("Count")>0 then
				strError="Please remove this Computer out of Atlas Network first"
			else
				call TrackingPCCode("DEL")				
				Response.Redirect("ComputerList.asp")			
				
			end if
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	DelComputer=strError
	
end function
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
	strPCCodeOld=""
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	'strConnect = Application("g_strConnect")
	'Set objDatabase = New clsDatabase
	intPCID = Request.Form("txtID")
	fgDel=Request.Form("fgstatus")
	
	if Request.QueryString("act") = "save" then
				
		strPCCode=Request.Form("txtPCCode")
		strPCCodeOld=Request.Form("txtPCCodeOld")
		strSupplier=Request.Form("txtSupplier")	
		if strSupplier="" then strSupplier=null
		dateBuying=cdate(Request.Form("lstMonthF") & "/" & Request.Form("lstDayF") & "/" & Request.Form("lstYearF"))
		intStatus=Request.Form("lstStatus")
		if intStatus="" then intStatus=1

		strNote=Request.Form("txtNote")
		if strNote="" then strNote=null
		strCPU=Request.Form("txtCPU")
		if strCPU="" then strCPU=null
		strSN=Request.Form("txtSN")
		if strSN="" then strSN=null
		
		strRAM=Request.Form("txtRAM")
		if strRAM="" then strRAM=null
		strHDD=Request.Form("txtHDD")
		if strHDD="" then strHDD=null
		strVideo=Request.Form("txtVideo")
		if strVideo="" then strVideo=null
		
		
		
		strArrayDisk1=Request.Form("txtArr10")& "@" & Request.Form("txtArr11")& "@" & Request.Form("txtArr12")& "@" & Request.Form("txtArr13")
		strArrayDisk2=Request.Form("txtArr20")& "@" & Request.Form("txtArr21")& "@" & Request.Form("txtArr22")& "@" & Request.Form("txtArr23")
		
		arrDisk1=Split(strArrayDisk1,"@")
		arrDisk2=Split(strArrayDisk2,"@")


		if fgDel<>"D" then
			
			if Cint(intPCID)=-1 then
				'Add new				

				strError=AddComputer(strPCCode,strSupplier,	dateBuying,intStatus,strNote,strCPU,strRAM,strHDD,strVideo,strArrayDisk1,strArrayDisk2,strSN)
			else
				
				intDetailPCID=Request.Form("txtDetailID")
				'Update
				strError=UpdateComputer(strPCCode,strSupplier,dateBuying,intStatus, strNote,strCPU,strRAM,strHDD,strVideo,strArrayDisk1,strArrayDisk2,intPCID,intDetailPCID)

			end if
		else
		
			strError=DelComputer(intPCID)
			fgDel=""
			
		end if

		'strError=ExecuteSQL(strQuery)
	else
		strPCCode=""
		strSupplier=""	
		dateBuying=Date()
		strNote=""
	
		strCPU=""
		strSN=""
		strRAM=""
		strHDD=""
		strVideo=""
		
		arrDisk1=Split("@@@","@")
		arrDisk2=Split("@@@","@")
		
		intStatus=1
	End If
'--------------------------------------------------
' 
'--------------------------------------------------
	
	strSql="SELECT a.PCID,a.PC_code,a.Supplier,a.BuyDate,a.Outdated,a.ComputerNote,b.DetailID," & _
				"b.CPUInformation,b.RAM,b.SystemMemory,b.VideoAdapter, a.SeriNum, Array1,Array2 FROM ATC_Computers a " & _
				"INNER JOIN ATC_ComputerDetails b ON a.PCID=b.PCID " & _
				"WHERE a.PCID=" & intPCID

	Call GetRecordset(strSql,rsData)
	
Response.Write strSql	
	
	If Request.QueryString("act") = "EDIT" Then			
		if rsData.RecordCount>0 then
			strPCCode=rsData("PC_Code")
			strPCCodeOld=strPCCode
			strSupplier=rsData("Supplier")	
			dateBuying=rsData("BuyDate")
			strNote=rsData("ComputerNote")
	
			strCPU=rsData("CPUInformation")
			strSN=rsData("SeriNum")
			strRAM=rsData("RAM")
			strHDD=rsData("SystemMemory")
			strVideo=rsData("VideoAdapter")
			
			strArrayDisk1=rsData("Array1")			
			arrDisk1=Split(strArrayDisk1,"@")
			
			strArrayDisk2=rsData("Array2")
			arrDisk2=Split(strArrayDisk2,"@")
			
			'lstStatus
			intStatus=  rsData("Outdated")		
			intDetailPCID=rsData("DetailID")
		end if			
	end if
	
	strSql ="SELECT AtlasPCID,Computername,TypeOfPC,c.description,ISNULL(PublicName,Username) as userAssignment " & _
			"FROM dbo.ATC_AtlasPC a	LEFT JOIN ATC_Users b ON a.UserID=b.UserID " & _
			" INNER JOIN ATC_ComputerType c ON a.TypeOfPC=c.AtlasPCTypeID " & _
			"WHERE PCID=" & intPCID & " ORDER BY Computername"
'Response.Write strSQl			
	Call GetRecordset(strSql,rsSrc)
	
	strLast=OutBody(rsSrc)
	
	strSql="SELECT * FROM ATC_ComputerStatus WHERE fgActivate=1 ORDER BY StatusDescription"
	Call GetRecordset(strSql,rsStatus)

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

<script type="text/javascript" src="../../library/library.js"></script>
<link href="../../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>

<script type="text/javascript" src="../../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../../jQuery/jquery-ui.min.js"></script>

<link href="../../jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>

<style type="text/css">

#DiskArrays
{
text-align:center;
background-color:White
}
#DiskArrays input
{
	width:100%;

</style>

<script type="text/javascript">


    $(document).ready(function() {
        $("#DiskArrays").toggle();

        $("#tongleDiskArrays").click(function() {
            $("#DiskArrays").toggle();
        });
    })

       
    
</script>

<script language="javascript">
<!--

    function AtlasNetwork(id) {
        window.document.frmreport.txtAtlasPCID.value = id
        window.document.frmreport.action = "AtlasComputer.asp?ID=" + id
        window.document.frmreport.submit();
    }

    function savedata() {
        if (checkdata()) {
            window.document.frmreport.action = "ComputerDetail.asp?act=save"
            window.document.frmreport.submit();
        }
    }

    function deletedata() {
        window.document.frmreport.fgstatus.value = "D"
        window.document.frmreport.action = "ComputerDetail.asp?act=save"
        window.document.frmreport.submit();
    }

    function checkdata() {
        if (window.document.frmreport.txtPCCode.value == "") {
            alert("Please enter PC code.");
            document.frmreport.txtPCCode.focus();
            return false
        }

        var dateFrom = document.frmreport.lstdayF.value + "/" + document.frmreport.lstmonthF.value + "/" + document.frmreport.lstyearF.value

        if (isdate(dateFrom) == false) {
            alert("The date (" + dateFrom + ") is invalid.");
            document.frmwh.lstdayF.focus();
            return false;
        }

        return true
    }
	
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
					<table width="120" border="0" cellspacing="2" cellpadding="0" align="right" height="20" name="aa">
                      <tr> 
                       
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue"><a href="javascript:AtlasNetwork(-1)" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" class="b">Join Atlas Network </a></div>
                        </td>
						                     
                      </tr>
                    </table>
                  </td>
                </tr>                
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Hardware Information</td>
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
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue" colspan="2"><img src="../../images/dot.gif"> General Information</td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="25%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="20%">PC Code *</td>
                              <td valign="middle" width="35%" class="blue">
								<input type="text" name="txtPCCode" maxlength="20" class="blue-normal" style="width:95%;" value="<%=strPCCode%>"></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Serial Number</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtSN" maxlength="50" class="blue-normal" size="20" style="width:95%" value="<%=strSN%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Supplier</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtSupplier" maxlength="200" class="blue-normal" style="width:95%;" value="<%=strSupplier%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Purchased date</td>
                              <td valign="middle" class="blue"><%
														Response.Write arrlstFrom(1)
														Response.Write arrlstFrom(0)
														Response.Write arrlstFrom(2)%></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr> 
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Status </td>
                              <td valign="middle" class="blue">
									<select name="lstStatus" style="width:95%;height=24px; background-color: #ffffff; border-style:1px; border-color: #A0AEA4" class="blue-normal">
										<%if rsStatus.RecordCount>0 then
											rsStatus.MoveFirst
											Do while not rsStatus.EOF%>						
												<option value="<%=rsStatus("StatusID")%>" <%if cint(intStatus)=rsStatus("StatusID") then%>selected<%end if%>><%=rsStatus("StatusDescription")%></option>		
										<%		rsStatus.MoveNext
										  loop
										end if%>
									</select>
							  </td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>    
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtNote" maxlength="200" class="blue-normal" size="20" style="width:95%" value="<%=strNote%>">
							</td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>                                
                            </tr>                                                     
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue" colspan="2"><img src="../../images/dot.gif"> Specification </td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">CPU</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtCPU" maxlength="100" class="blue-normal" size="20" style="width:95%" value="<%=strCPU%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            
                             <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">HDD</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtHDD" maxlength="50" class="blue-normal" size="20" style="width:95%" value="<%=strHDD%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">RAM</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtRAM" maxlength="150" class="blue-normal" size="20" style="width:95%" value="<%=strRAM%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Video Adapter</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtVideo" maxlength="150" class="blue-normal" size="20" style="width:95%" value="<%=strVideo%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr> 
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue" colspan="2"><img src="../../images/dot.gif">
                                <a class="blue c" href="#" id="tongleDiskArrays">Disk Arrays<img src="../../images/DownArrow-Icon.gif" style="border:0;" /></a>
                                <p></p></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>                                                                                                                              
                          </table>
  
                           <div id="DiskArrays">
                               <table id="tblDisks" style="width:70%;margin: 0 auto">
                                    <tr>
                                        <th style="width:13%"></th>
                                        <th class="blue-normal" style="width:14%">Disk Qty</th>
                                        <th class="blue-normal" style="width:40%">Disk type</th>
                                        <th class="blue-normal" style="width:18%" >RAID level</th>
                                        <th class="blue-normal" style="width:15%">Capacity</th>
                                    </tr>
                                    <tr>
                                        <td class="blue-normal">Array 1</td>
                                        <td><input type="text" id="txtArr10" name="txtArr10" class="blue-normal" maxlength="99"  value="<%=showlabel(arrDisk1(0)) %>" /></td>
                                        <td><input type="text" id="txtArr11" name="txtArr11" class="blue-normal" maxlength="99" value="<%=showlabel(arrDisk1(1))%>" /></td>
                                        <td><input type="text" id="txtArr12" name="txtArr12" class="blue-normal" maxlength="99" value="<%=showlabel(arrDisk1(2))%>" /></td>
                                        <td><input type="text" id="txtArr13" name="txtArr13" class="blue-normal" maxlength="99" value="<%=showlabel(arrDisk1(3))%>" /></td>
                                    </tr>
                                    <tr>
                                        <td class="blue-normal">Array 2</td>
                                        <td><input type="text" id="txtArr20" name="txtArr20" class="blue-normal" maxlength="99"  value="<%=showlabel(arrDisk2(0))%>" /></td>
                                        <td><input type="text" id="txtArr21" name="txtArr21" class="blue-normal" maxlength="99" value="<%=showlabel(arrDisk2(1))%>" /></td>
                                        <td><input type="text" id="txtArr22" name="txtArr22" class="blue-normal" maxlength="99" value="<%=showlabel(arrDisk2(2))%>" /></td>
                                        <td><input type="text" id="txtArr23" name="txtArr23" class="blue-normal" maxlength="99" value="<%=showlabel(arrDisk2(3))%>" /></td>

                                    </tr>
                                </table>
                        </div>
                       
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table width="120" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:savedata()" onMouseOver="self.status='Please click here to save changes';return true" onMouseOut="self.status='';return true" class="b">Save</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:deletedata()" onMouseOver="self.status='Please click here to delete this record';return true" onMouseOut="self.status='';return true" class="b">Delete</a></div>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
<%if strLast<>"" then %>                          
						  <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="10%">No.</td>
                              <td class="blue" align="center" width="30%">Computer name</td>  
                              <td class="blue" align="center" width="30%">Type</td>  
                              <td class="blue" align="center" width="30%">User Assignment</td>
                            </tr>
<%Response.Write strLast%>
                          </table>   
<%end if%>                                                 
                          
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
<input type="hidden" name="txtAtlasPCID" value="">
<input type="hidden" name="txtID" value="<%=intPCID%>">
<input type="hidden" name="txtDetailID" value="<%=intDetailPCID%>">
<input type="hidden" name="txtPCCodeOld" value="<%=strPCCodeOld%>">
</form>

</body>
</html>
