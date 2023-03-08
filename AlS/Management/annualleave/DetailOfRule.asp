<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' Function: Outbody
' Description: List of employees
' Parameters: source recordset
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc)
	dim strApplyDate
	strOut = ""
	i = 0
		
	while not rsSrc.EOF
		strColor = "#FFF2F2"
		if i mod 2 = 0 then	strColor = "#E7EBF5"
		
		strApplyDate=day(rsSrc("Applydate")) & "-" & MonthName(month(rsSrc("Applydate")),true) & "-" & year(rsSrc("Applydate"))
		
		strOut = strOut & "<tr bgcolor='" & strColor & "'>" &_
				"<td valign='top' class='blue'>&nbsp;" & i + 1 & "</td>" &_
				"<td valign='top' class='blue'>&nbsp;" & showlabel(rsSrc("Fullname")) & "</td>" &_
				"<td valign='top' class='blue'>" &  showlabel(rsSrc("Jobtitle")) & "</td>" &_
				"<td valign='top' class='blue'>" & strApplyDate & "</td>" &_
				"<td valign='top' class='blue'><input type='checkbox' name='chkass' value='" & rsSrc("StaffAnnualLeaveID") & "'" & " " & strCHK & "></td>" &_
				"</tr>"
				
		rsSrc.MoveNext
		i = i + 1
	wend	
	Outbody = strOut
end function

'==================================================================
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage,gErrMessage
	Dim arrlstFrom(2),arrlstTo(2),arrlongmon
	Dim rsData,strAct,strStatus
	
	gMessage=""

	arrlongmon  = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
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
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = (getRight(1, ii) = 1)
				exit for
			end if
		next
		set getRight = nothing
	end if	
	
	fgRight=true
	
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
	
	intRuleID=Request.Form("txtRuleID")
	
'--------------------------------------------------
' Get data from ATC_AnnualLeaveDefaultRule
'--------------------------------------------------

strSQL="SELECT RuleID,RatePerYear,RuleNote,fgActivate FROM ATC_AnnualLeaveDefaultRule WHERE RuleID=" & intRuleID

Call GetRecordset(strSQL,rsData)

if gMessage="" then 
	if rsData.RecordCount>0 then		
		dblRate=rsData("RatePerYear")			
	
		strNote=rsData("RuleNote")
		fgActivate=	rsData("fgActivate")
	end if

	rsData.close
end if
set rsData=nothing

'--------------------------------------------------
'Deassign Staff
'--------------------------------------------------

strAct=Request.QueryString("act")

if strAct="de" then
	countU = Request.Form("chkass").Count

	if countU>0 then
		strUpdate=""
		strDelete=""
		
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		if ret then
			objDb.cnDatabase.BeginTrans
			For i = 1 to countU
				varBook = int(Request.Form("chkass")(i))
				'If CheckDel(varBook) then 'delete
					strDelete = strDelete & varBook & ","
				'End if
			Next
			
			if strDelete<>"" then 
				strDelete="DELETE FROM ATC_EmployeeAnnualLeave WHERE StaffAnnualLeaveID IN (" & Left(strDelete,len(strDelete)-1) & ")"
				if not objDb.runActionQuery(strDelete) then gMessage = objDb.strMessage
			end if
						
			if gMessage<>"" then 
				objDb.cnDatabase.RollbackTrans
			else
				objDb.cnDatabase.CommitTrans
			end if
			objDb.dbdisConnect
		else
			gMessage = objDb.strMessage 'error in connection
		end if
		set objDb = nothing
	end if 

end if
	
'--------------------------------------------------
' Get employees who are current in this rule
'--------------------------------------------------

strSQL="SELECT a.*, b.Fullname, b.Jobtitle FROM ATC_EmployeeAnnualLeave a INNER JOIN HR_Employee b ON b.PersonID=a.StaffID WHERE RuleID=" & intRuleID & " ORDER BY Fullname"

Call GetRecordset(strSQL,rsData)

strList=""

if gMessage="" then 
	if rsData.RecordCount>0 then		
		strList=Outbody(rsData)
		
	end if

	rsData.close
end if
set rsData=nothing
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
<title>AIS System</title>
<script language="javascript" src="../../library/library.js"></script>
<link rel="stylesheet" href="../../timesheet.css">
<script>
<!--
function showhide(layer_ref,val) { 
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
} 

function Assign()
{
	
	document.frmwh.action = "selectemployee_ass.asp?outside=1";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function DeAssign()
{
	
	document.frmwh.action = "DetailOfRule.asp?act=de";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function setchecked(val) {
  with (document.frmwh) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].type == "checkbox") {
			elements[ii].checked = val;
		}
	}
  }
}

-->
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
                <tr valign="middle">
                  <td class="title" height="50" align="center">Detail of Annual Leave Rule</td>
                </tr>
        
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top">
					<table width="55%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#003399">
                      <tr> 
                        <td > <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" >
                            <tr> 
                              <td bgcolor="#C0CAE6" >
                              
								<table width="100%" border="0" cellspacing="5" cellpadding="0">
                                  <tr> 
                                    <td valign="middle" class="blue-normal" width="30%">&nbsp;&nbsp;Rate per year *</td>
                                    <td valign="middle" width="70%" class="blue"> <%=dblRate%> (days)
                                    </td>
                                  </tr>
                                   
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;&nbsp;</td>
                                    <td valign="middle" class="blue-normal"> 
											<%=showlabel(strNote)%>                             </td>
                                  </tr>
                                  
                                </table>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                    </table> </td>
                </tr>     
                 
                <tr> 
                  <td class="blue" height="20" align="right">&nbsp;&nbsp;
					<table border="0" cellspacing="1" cellpadding="4">
                            <tr> 
                              <td class="blue" align="center" ><a href="javascript:Assign();"  onMouseOver="self.status='Assign Staff'; return true;" onMouseOut="self.status=''" >Assign</a></td>
                              <td class="blue" align="center" ><a href="javascript:DeAssign();"  onMouseOver="self.status='DeAssign'; return true;" onMouseOut="self.status=''" >Deassign</a></td>
                            </tr>
                      </table>
					</td>
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
                              <td class="blue" bgcolor="8CA0D1" align="center" width="5%">No.</td>
                              <td class="blue" align="center" width="35%">Full Name</td>
                              <td class="blue" align="center" width="35%">Job Title</td>
                              <td class="blue" align="center" width="20%">Apply date </td>
                              <td class="blue" align="center" width="5%"></td>
                            </tr>
<%Response.Write strList%>
                          </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td bgcolor="#FFFFFF" height="20" class="blue" align="right"><a href="javascript:setchecked(1);" class="c" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
			      All</a>&nbsp;&nbsp;&nbsp;<a href="javascript:setchecked(0);" class="c" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''"> Clear All</a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
			  </tr>
			  <tr>
			    <td bgcolor="#FFFFFF" height="20" class="blue-normal">&nbsp;&nbsp;</td>
			  </tr>
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