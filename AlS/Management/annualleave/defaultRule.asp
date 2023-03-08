<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' Function: Outbody1
' Description: holiday
' Parameters: source recordset
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************
function Outbody1(ByRef rsSrc)
	dim strHName
	strOut = ""
	i = 0
	arrlongmon  = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	
	while not rsSrc.EOF
		strColor = "#FFF2F2"
		if i mod 2 = 0 then	strColor = "#E7EBF5"
	
		strHName="<a class='c' href='javascript:ruleDetail("& rsSrc("RuleID") & ");' " &_
				"OnMouseOver = 'self.status=&quot;Update&quot; ; return true' OnMouseOut = 'self.status = &quot;&quot;'>" & formatnumber(rsSrc("RatePerYear"),2) & "</a>"
				
		
		strOut = strOut & "<tr bgcolor='" & strColor & "'>" &_
				"<td valign='top' class='blue'>&nbsp;" & i + 1 & "</td>" &_
				"<td valign='top' class='blue'>&nbsp;" & strHName & "</td>" &_
				"<td valign='top' class='blue'>" & showlabel(rsSrc("Rulenote")) & "</td>" &_
				"</tr>"
				
		rsSrc.MoveNext
		i = i + 1
	wend	
	Outbody1 = strOut
end function

'*****************************************
'AddNewRule
'*****************************************
Function AddNewRule
	dim strConnect,objDb,ret,idxDate
	ret=true
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then	
		
		strSql="INSERT INTO ATC_AnnualLeaveDefaultRule (RatePerYear,RuleNote,fgLongService) " & _
				"VALUES(" & dblRate & ",'" & Replace(strNote,"'","''") & "'," & IIF(fgLongService,1,0)& ")"
					
		ret = objDb.runActionQuery(strSql)
		  
	else
		gErrMessage=objDb.strMessage
	end if
End Function
'*****************************************
'UpdateHoliday
'*****************************************
Function UpdateRule
	dim strConnect,objDb,ret,idxDate
	ret=true
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	
	If objDb.dbConnect(strConnect) then	
		
		strQuery = "UPDATE ATC_AnnualLeaveDefaultRule " & _
			"SET RatePerYear = " & dblRate & _
		     ",RuleNote = '" & Replace(strNote,"'","''") & "'" & _
		     ",fgActivate = " & IIF(fgActivate,1,0)  & _
			 ",fgLongService = " & IIF(fgLongService,1,0)  & _
		" WHERE RuleID=" & intRuleID
		
		ret = objDb.runActionQuery(strQuery)
				
		if ret=false then				
			gErrMessage = objDb.strMessage
		else
			gErrMessage="Update successfully."
		end if
		  
	else
		gErrMessage=objDb.strMessage
	end if
End Function

'==================================================================
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage,gErrMessage
	Dim dFromOld,dToOld,dFromNew,dToNew
	Dim arrlstFrom(2),arrlstTo(2),arrlongmon
	Dim strHolidayName,intRatio,rsHoliday,strAct,strStatus
	
	arrlongmon  = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	gMessage=""

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If

'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.ServerVariables("URL") 

	while Instr(tmp, "/")<>0
		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
	Wend
	strFilename = tmp
	fgRight=not isEmpty(session("Righton"))
	
	if not isEmpty(session("Righton")) then
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if lcase(getRight(0, ii)) = lcase(tmp) then
				fgRight=true
				fgUpdate = (getRight(1, ii) = 1)
				exit for
			end if
		next
		set getRight = nothing
	end if	

	if fgRight = false then Response.Redirect("../../welcome.asp")
		
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
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
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
	if strChoseMenu="" then strChoseMenu = "AE"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Get data from holiday detail
'--------------------------------------------------
	strAct=Request.QueryString("act")
		
	if strAct="EDIT" Then
		intRuleID=Request.Form("ruleID")
	Else 
		intRuleID=-1	
	End If
	
	
	dblRate=Request.Form("txtRatePerYear")
		
	strNote=Request.Form("txtNote")
	fgActivate= (Request.Form("optActivate")="1")
	fgLongService=(Request.Form("optLongService")="1")
	
	strStatus=strAct
'--------------------------------------------------
' Perform saving data to atc_holiday
'--------------------------------------------------
	If strAct="SAVE" then
		intRuleID=Request.Form("txtRuleID")
		if intRuleID="" then intRuleID=-1
		if cint(intRuleID)=-1 then
			call AddNewRule
			strAct="EDIT"
		else
			call UpdateRule
		end if
		
	Elseif strAct="DEL" then
		call DeleteRule
	End If
	

'--------------------------------------------------
' Get data from ATC_AnnualLeaveDefaultRule
'--------------------------------------------------

if strview="ALL" then
	strSQL="SELECT RuleID,RatePerYear,RuleNote,fgActivate,fgLongService FROM ATC_AnnualLeaveDefaultRule ORDER BY RuleID DESC"
else
	strSQL="SELECT RuleID,RatePerYear,RuleNote,fgActivate,fgLongService FROM ATC_AnnualLeaveDefaultRule WHERE fgActivate=1 ORDER BY RuleID DESC"
end if

Call GetRecordset(strSQL,rsData)

if gMessage="" then 

	strlist = OutBody1(rsData)	

	if strAct="EDIT" then
		intRuleID=Request.Form("txtRuleID")
			
		rsData.MoveFirst
		rsData.Filter="RuleID=" & intRuleID
	
		if rsData.RecordCount>0 then	
			
			dblRate=FormatNumber(rsData("RatePerYear"),2)
	
			strNote=rsData("RuleNote")
			fgActivate=	rsData("fgActivate")
			fgLongService=	rsData("fgLongService")
		end if
	end if
	rsData.close
end if
set rsData=nothing

	arrlstFrom(0) = selectmonth("lstmonthF",intMonth , -1)
	arrlstFrom(1) = selectday("lstdayF", intday, -1)
'--------------------------------------------------
' Read teamplate
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if

gMessage=gMessage & gErrMessage
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<script language="javascript" src="../../library/library.js"></script>
<link rel="stylesheet" href="../../timesheet.css">
<script>

function ruleDetail(id) {

	document.frmwh.txtRuleID.value=id;	
	
	document.frmwh.action = "defaultRule.asp?act=EDIT";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function Add() {
	document.frmwh.txtRatePerYear.value = "";
	document.frmwh.txtNote.value = "";
	document.frmwh.txtRuleID.value=-1
	
	document.frmwh.action = "defaultRule.asp";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function Delete(){
	document.frmwh.action = "defaultRule.asp?act=DEL";
	document.frmwh.target = "_self";
	//document.frmwh.submit();
}

function Details()
{	
	document.frmwh.action = "DetailOfRule.asp";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}


function SetupExpiredRules()
{	
	document.frmwh.action = "ExpiredRule.asp";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function CheckData(){
	
	//var blnCheckDay="<%=(arrlstTo(2)<>"")%>"
	//var dToday="<%=day(Date()+1) & "/" & month(Date()) & "/" & Year(Date())%>"

	if (document.frmwh.txtRatePerYear.value == "") {
		alert("Please enter Rate Per Year.");
		document.frmwh.txtRatePerYear.focus();
		return false;
		}
	else
		if (isNaN(document.frmwh.txtRatePerYear.value)==true) {
			alert("Please enter a number.");
			document.frmwh.txtRatePerYear.focus();
			return false;
		}
		else if (document.frmwh.txtRatePerYear.value<0) {
			alert("The Rate value must be greater than 0.");
			document.frmwh.txtRatePerYear.focus();
			return false;			
		}
	
	
	return true;
}

function Save(){
	if (CheckData()==true){
		
		document.frmwh.action = "defaultRule.asp?act=SAVE";
		document.frmwh.target = "_self";
		document.frmwh.submit();
	}
}

//--> 

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmwh" method="post">
    	<%	'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))%>		
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                              <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
                  <td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
                </tr>
              
				<tr align="center"> 
                  
                  <td class="blue" height="30" align="right" width="77%">
					<table width="150" border="0" cellspacing="2" cellpadding="0" align="right" height="20" name="aa">
                      <tr> 
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" >
                          <div align="center" class="blue"><a href="javascript:SetupExpiredRules()" onMouseOver="self.status='Please click here to view staff Annual Leave.';return true" onMouseOut="self.status='';return true" class="b">Setup Expired Rule</a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>              
                <tr valign="middle">
                  <td class="title" height="50" align="center">Setup Annual Leave Rules</td>
                </tr>
<%if fgUpdate then%>              
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top">
					<table width="55%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#003399">
                      <tr> 
                        <td > 
							<table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" >
								<tr>
								  <td bgcolor="#C0CAE6" >
									<table width="100%" border="0" cellspacing="5" cellpadding="0">
									  <tr>
										<td valign="middle" class="blue-normal" width="30%">&nbsp;&nbsp;Rate per year *</td>
										<td valign="middle" width="70%" class="blue-normal"> 
										  <input type="text" name="txtRatePerYear" maxlength="30" class="blue-normal" size="20" style='width:95%' value="<%=dblRate%>">
										</td>
									  </tr>
									  <tr> 
										<td valign="middle" class="blue-normal">&nbsp;&nbsp;Note</td>
										<td valign="middle" class="blue-normal"> 
												<input type="text" name="txtNote" maxlength="30" class="blue-normal" size="20" style='width:95%' value="<%=showlabel(strNote)%>"></td>
									  </tr>
									  <tr> 
										<td valign="middle" class="blue-normal">&nbsp;&nbsp;</td>
										<td valign="middle" class="blue-normal"> 
										<input type="checkbox" name="optActivate" value="1" <%if fgActivate then%>checked<%end if%>>Activate
										<input type="checkbox" name="optLongService" value="1" <%if fgLongService then%>checked<%end if%>>Extra leave for 
long service
										</td>
									  </tr>
									 
									  <tr> 
										<td valign="middle" class="blue-normal">&nbsp;</td>
										<td valign="middle" class="blue-normal"><table border="0" cellspacing="5" cellpadding="0" align="right" height="20" name="aa">
											<tr> 
											  <td width="70" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center"> 
												<a href="javascript:Save();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Save</a></td>
	<%if strAct="EDIT" then%>                                            
											  <td width="70" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
												<a href="javascript:Delete();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Delete</a></td>
											 <td width="70" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
												<a href="javascript:Details();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Detail</a></td>
	<%End if%>											
											</tr>
										  </table></td>
									  </tr>
									</table>
								  </td>
								</tr>
                          </table></td>
                      </tr>
                    </table> </td>
                </tr>     
<%end if%>                      
                <tr> 
                  <td class="blue" height="20" align="left">&nbsp;&nbsp;
<%if fgUpdate then%><a href="javascript:Add();" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add New</a><%end if%></td>
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
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="4">
                            <tr bgcolor="8CA0D1"> 
                              <td class="blue" bgcolor="8CA0D1" align="center" width="10%">No.</td>
                              <td class="blue" align="center" width="30%">Rate per year</td>
                              <td class="blue" align="center" width="60%">Note </td>
                            </tr>
<%Response.Write strList%>
                          </table>
                          
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
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
<input type="hidden" name="txtRuleID" value=<%=intRuleID%>>
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>