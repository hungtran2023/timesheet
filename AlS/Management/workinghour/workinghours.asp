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
		strdayfrom = Day(rsSrc("Dfrom")) & "-" & arrlongmon(month(rsSrc("Dfrom"))-1) & "-" & CStr(year(rsSrc("Dfrom")))
		strdayto =  Day(rsSrc("DTo")) & "-" & arrlongmon(month(rsSrc("DTo"))-1) & "-" & CStr(year(rsSrc("DTo")))
		
		strHName="<a class='c' href='javascript:holidaydetail(&quot;"& rsSrc("Dfrom") & "&quot;,&quot;" & rsSrc("DTo") & "&quot;);' " &_
				"OnMouseOver = 'self.status=&quot;Update&quot; ; return true' OnMouseOut = 'self.status = &quot;&quot;'>" & showlabel(rsSrc("Holiday")) & "</a>"
		
		strOut = strOut & "<tr bgcolor='" & strColor & "'>" &_
				"<td valign='top' width='18%' class='blue'>&nbsp;" & strdayfrom & "</td>" &_
				"<td valign='top' width='18%' class='blue'>&nbsp;" & strdayto & "</td>" &_
				"<td valign='top' width='46%' class='blue'>" & strHName & "</td>" &_
				"<td valign='top' width='18%' class='blue' align='center'>" & showlabel(rsSrc("ratio")) & "</td>" &_
				"</tr>"
		rsSrc.MoveNext
		i = i + 1
	wend	
	Outbody1 = strOut
end function
'*****************************************
'Check Data
'**************************************
Function CheckHolidayData
	Dim blnReturn,strSql,rsTemp
	
	blnReturn=true

	strSql="SELECT * FROM ATC_Holiday WHERE  (Holiday='" & Replace(strHolidayName,"'","''") & "' OR " & _
				"(CONVERT(Datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay))>='" & dFromNew & "' " & _
				"AND CONVERT(Datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay))<='" & dToNew & "')) AND syear>=year(Getdate()) " & _
				"AND (CONVERT(Datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay))<'" & dFromOld & "' " & _
						"OR CONVERT(Datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay))>'" & dToOld & "' )"
'Response.Write strSql	
	call GetRecordset(strSql,rsTemp)
	
	if gMessage="" then 
		if rsTemp.Recordcount>0 then 
			gErrMessage="This '" & Request.Form("txtname") & "' or holidays from " & dFromNew & " To " & dToNew & " has already been inputted."
		end if
	else
		gErrMessage=gMessage	
	End if
	
	gMessage=""
	CheckHolidayData=blnReturn and gErrMessage="" 
End Function
'*****************************************
'AddNewHoliday
'*****************************************
Function AddNewHoliday
	dim strConnect,objDb,ret,idxDate
	ret=true
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
		objDb.cnDatabase.BeginTrans
			
		For idxDate=cdate(dFromNew) to cdate(dToNew)
			if weekday(idxDate)<>1 AND weekday(idxDate)<>7 then
				strSql="INSERT INTO ATC_Holiday(Holiday, smonth, sday, syear, ratio) VALUES " & _
						"('" & replace(strHolidayName,"'","''") & "'," & month(idxDate) & "," & day(idxDate) & "," & year(idxDate) & "," & intRatio & ")"
'Response.Write strSql						
				ret = objDb.runActionQuery(strSql)
				if not ret then Exit for
			end if
		Next
			
		if ret=false then				
			objDb.cnDatabase.RollbackTrans
			gErrMessage = objDb.strMessage
		else
			objDb.cnDatabase.CommitTrans
				
		end if
		  
	else
		gErrMessage=objDb.strMessage
	end if
End Function
'*****************************************
'UpdateHoliday
'*****************************************
Function UpdateHoliday
	dim strConnect,objDb,ret,idxDate
	ret=true
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	
	If objDb.dbConnect(strConnect) then
		objDb.cnDatabase.BeginTrans
		strQuery = "DELETE ATC_Holiday WHERE CONVERT(Datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay)) BETWEEN '" & dFromOld & "' AND '" & dToOld & "'" 
	  
		ret = objDb.runActionQuery(strQuery)
		if ret=true then
			For idxDate=cdate(dFromNew) to cdate(dToNew)
				if weekday(idxDate)<>1 AND weekday(idxDate)<>7 then
					strSql="INSERT INTO ATC_Holiday(Holiday, smonth, sday, syear, ratio) VALUES " & _
							"('" & replace(strHolidayName,"'","''") & "'," & month(idxDate) & "," & day(idxDate) & "," & year(idxDate) & "," & intRatio & ")"
					ret = objDb.runActionQuery(strSql)
					if not ret then Exit for
				end if
			Next
		end if
		
		if ret=false then				
			objDb.cnDatabase.RollbackTrans
			gErrMessage = objDb.strMessage
		else
			objDb.cnDatabase.CommitTrans
			gErrMessage="Update successfully."
		end if
		  
	else
		gErrMessage=objDb.strMessage
	end if
End Function
'*****************************************
'DeleteHoliday
'*****************************************
Function DeleteHoliday
	dim strConnect,objDb,ret,idxDate,strQuery
	ret=true
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  objDb.cnDatabase.BeginTrans
	  strQuery = "DELETE ATC_Holiday WHERE CONVERT(Datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay)) BETWEEN '" & dFromOld & "' AND '" & dToOld & "'" 
	  ret = objDb.runActionQuery(strQuery)
	  if ret=false then
	  	objDb.cnDatabase.RollbackTrans
		gErrMessage = objDb.strMessage
	  else
		objDb.cnDatabase.CommitTrans
		gErrMessage = "Deleted successfully."
		
		dFromNew=date()
		dToNew=date()
		strHolidayName=""
		intRatio=""
	  end if
	  objDb.dbdisConnect
	end if
	set objDb = nothing
End Function

'==================================================================
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage,gErrMessage
	Dim dFromOld,dToOld,dFromNew,dToNew
	Dim arrlstFrom(2),arrlstTo(2),arrlongmon
	Dim strHolidayName,intRatio,rsHoliday,strAct,strStatus
	
	arrlongmon  = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	gMessage=""

Call freeListpro
Call freeProInfo
Call freeAssignment
Call freeAssignRight
Call freeShort
Call freeSinglepro
Call freeSumpro
Call freelistEmp

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
	if strChoseMenu="" then strChoseMenu = "A"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Get data from holiday detail
'--------------------------------------------------
	strAct=Request.QueryString("act")
	
	dFromOld=Request.Form("txtDFrom")
	dToOld=Request.Form("txtDTo")
	
	if strAct="EDIT" Then
		dFromNew=dFromOld
		dToNew=dToOld
	Elseif strAct="" then
		dFromNew=date()
		dToNew=date()
	else
		dFromNew=Request.Form("lstmonthF") & "/" & Request.Form("lstdayF") & "/" & Request.Form("lstyearF")
		dToNew=Request.Form("lstmonthT") & "/" & Request.Form("lstdayT") & "/" & Request.Form("lstyearT")			
	End If
	
	if dFromNew="//" OR dToNew="//" Then 
		dFromNew=dFromOld
		dToNew=dToOld
	End if
	
	strHolidayName=Request.Form("txtName")
	intRatio=Request.Form("txtratio")
	strStatus=strAct
'--------------------------------------------------
' Perform saving data to atc_holiday
'--------------------------------------------------
	If strAct="SAVE" then
		If CheckHolidayData() Then
			strStatus=Request.Form("txtStatus")		
			if strStatus="" then
				call AddNewHoliday
			else
				call UpdateHoliday
			end if
		end if
	Elseif strAct="DEL" then
		call DeleteHoliday
	End If

	arrlstFrom(0) = selectmonth("lstmonthF",Month(cdate(dFromNew)), -1)
	arrlstFrom(1) = selectday("lstdayF", Day(CDate(dFromNew)), -1)
	arrlstFrom(2) = selectyear("lstyearF", Year(CDate(dFromNew)),year(now()), year(now()) +1, -1)

	arrlstTo(0) = selectmonth("lstmonthT",Month(cdate(dToNew)), -1)
	arrlstTo(1) = selectday("lstdayT", Day(CDate(dToNew)), -1)
	arrlstTo(2) = selectyear("lstyearT", Year(CDate(dToNew)),year(now()), year(now())+1, -1)

'--------------------------------------------------
' Get data from atc_holiday
'--------------------------------------------------

strSQL="SELECT MIN(CONVERT(datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay))) as Dfrom, " & _
			"MAX(CONVERT(datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay))) as Dto, Holiday, Ratio, " &_
			"fgEdit=CASE WHEN MIN(CONVERT(datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay)))<Getdate() OR MAX(CONVERT(datetime,str(sYear) + '-' + str(sMonth) + '-' + str(sDay)))<Getdate() THEN 0 ELSE 1 END " & _
		"FROM ATC_Holiday WHERE sYear>=year(getdate()) GROUP BY Holiday, Ratio ORDER BY Dfrom DESC"

Call GetRecordset(strSQL,rsHoliday)

if gMessage="" then 
	strholiday = OutBody1(rsHoliday)	
	if strAct="EDIT" then	
		rsHoliday.MoveFirst
		rsHoliday.Filter="DFrom='" & dFromOld & "' AND DTo='" & dToOld & "'"
	
		if rsHoliday.RecordCount>0 then 
			intRatio=rsHoliday("Ratio")
			strHolidayName=rsHoliday("Holiday")
			if rsHoliday("fgEdit")=0 then
				arrlstFrom(1)="<b>" & day(dFromOld) & "-" & arrlongmon(month(dFromOld)-1) & "-" & year(dFromOld) & "</b>"
				arrlstFrom(0)=""
				arrlstFrom(2)=""
				
				arrlstTo(1)="<b>" & day(dToOld) & "-" & arrlongmon(month(dToOld)-1) & "-" & year(dToOld) & "</b>"
				arrlstTo(0)=""
				arrlstTo(2)=""		
			End If
		end if
	end if
	rsHoliday.close
end if
set rsHoliday=nothing

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

function holidaydetail(dFrom,dTo) {
	document.frmwh.txtDFrom.value = dFrom;
	document.frmwh.txtDTo.value = dTo;
	
	document.frmwh.action = "workinghours.asp?act=EDIT";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function Add() {
	document.frmwh.txtDFrom.value = "";
	document.frmwh.txtDTo.value = "";
	document.frmwh.txtname.value = "";
	document.frmwh.txtratio.value = "";
	
	document.frmwh.action = "workinghours.asp";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function Delete(){
	document.frmwh.action = "workinghours.asp?act=DEL";
	document.frmwh.target = "_self";
	document.frmwh.submit();
}

function CheckData(){
	
	var blnCheckDay="<%=(arrlstTo(2)<>"")%>"
	var dToday="<%=day(Date()+1) & "/" & month(Date()) & "/" & Year(Date())%>"

	if (document.frmwh.txtname.value == "") {
		alert("Please enter value for this field.");
		document.frmwh.txtname.focus();
		return false;
	}
	if (document.frmwh.txtratio.value == "") {
		alert("Please enter value for this field.");
		document.frmwh.txtratio.focus();
		return false;
	}
	else
		if (isNaN(document.frmwh.txtratio.value)==true) {
			alert("Please enter a number.");
			document.frmwh.txtratio.focus();
			return false;
		}
		else if (document.frmwh.txtratio.value<=0) {
			alert("The ratio value must be greater than 0.");
			document.frmwh.txtratio.focus();
			return false;			
		}
		
	
	if (blnCheckDay=="True"){
		var dateFrom=document.frmwh.lstdayF.value + "/" + document.frmwh.lstmonthF.value + "/" + document.frmwh.lstyearF.value;
		var dateTo=document.frmwh.lstdayT.value + "/" + document.frmwh.lstmonthT.value + "/" + document.frmwh.lstyearT.value;

		if (isdate(dateFrom)==false){
			alert("The first date (" + dateFrom + ") is invalid.");
			document.frmwh.lstdayF.focus();
			return false;
		}
		if (isdate(dateTo)==false){
			alert("The last date (" + dateTo + ") is invalid.");
			document.frmwh.lstdayT.focus();
			return false;
		}
	
		if (comparedate(dateFrom,dateTo)==false){
			alert("The startdate must be less than the finishdate.")
			document.frmwh.lstdayF.focus();
			return false;
		}
		if (comparedate(dToday,dateFrom)==false || comparedate(dToday,dateTo)==false){
			alert("The startdate and enddate must be greater today.")
			document.frmwh.lstdayF.focus();
			return false;
		}
	}
	return true;
}

function Save(){
	if (CheckData()==true){
		document.frmwh.action = "workinghours.asp?act=SAVE";
		document.frmwh.target = "_self";
		document.frmwh.submit();
	}
}

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
                  <td class="title" height="50" align="center">Holidays</td>
                </tr>
<%if fgUpdate then%>              
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top">
					<table width="55%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#003399">
                      <tr> 
                        <td > <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" >
                            <tr> 
                              <td bgcolor="#C0CAE6" >
                              
								<table width="100%" border="0" cellspacing="5" cellpadding="0">
                                  <tr> 
                                    <td valign="middle" class="blue-normal" width="30%">&nbsp;&nbsp;Name</td>
                                    <td valign="middle" width="70%" class="blue-normal"> 
                                      <input type="text" name="txtname" maxlength="100" class="blue-normal" size="20" style='width:95%' value="<%=showlabel(strHolidayName)%>">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;&nbsp;Overtime Ratio</td>
                                    <td valign="middle" class="blue-normal"> 
                                      <input type="text" name="txtratio" class="blue-normal" maxlength="5" size="5" style='width:50%' value='<%=intRatio%>'> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;&nbsp;From</td>
                                    <td valign="middle" class="blue-normal"> 
<%
	Response.Write arrlstFrom(1)
	Response.Write arrlstFrom(0)
	Response.Write arrlstFrom(2)
%>                                      
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;&nbsp;To</td>
                                    <td valign="middle" class="blue-normal"> 
<%
	Response.Write arrlstTo(1)
	Response.Write arrlstTo(0)
	Response.Write arrlstTo(2)
%>                                        
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;</td>
                                    <td valign="middle" class="blue-normal"><table border="0" cellspacing="5" cellpadding="0" align="right" height="20" name="aa">
                                        <tr> 
                                          <td width="70" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center"> 
                                            <a href="javascript:Save();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Save</a></td>
<%'If in Edit mode and this holiday still be in the future
if strAct="EDIT" AND arrlstFrom(2)<>"" then%>                                            
                                          <td width="70" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
											<a href="javascript:Delete();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Delete</a></td>
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
<%if fgUpdate then%><a href="javascript:Add();" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add New</a>
<%end if%></td>
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
                              <td class="blue" bgcolor="8CA0D1" align="center" width="18%">From</td>
                              <td class="blue" align="center" width="18%">To</td>
                              <td class="blue" align="center" width="46%">Holiday</td>
                              <td class="blue" align="center" width="18%">Overtime Ratio </td>
                            </tr>
<%Response.Write strHoliday%>
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
<input type="hidden" name="txtDFrom" value=<%=Request.Form("txtDFrom")%>>
<input type="hidden" name="txtDTo" value=<%=Request.Form("txtDTo")%>>
<input type="hidden" name="txtStatus" value=<%=strStatus%>>
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>