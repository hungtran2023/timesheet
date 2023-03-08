<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDatabase, strError, intPageSize, fgRight 'view all or Not
	dim arrSortType
	arrSortType=array("ASC","DESC")
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

Function Outbody(ByRef rsSrc,  ByVal intPage,ByVal intPageSize)

	dim dblExRate,strStatus,strNextStep
	'Method 3 : using 'Array' Parameter
	Dim arrStatus, arrNext
	arrStatus = Array("New","In-progress","Rejected","Done")
	strOut = ""
	
	strYes="<img src='../../images/yes.gif' width='15' height='17'>"
	strNo="<img src='../../images/no.gif' width='14' height='17'>"
	
	If Not rsSrc.EOF Then
			
		rsSrc.AbsolutePage = intPage
		intStart = rsSrc.AbsolutePosition
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsData.PageSize - 1)
		End if
	
	    strOut = ""
	    For i = intStart to intFinish
			strColor = "#FFF2F2"
			If i Mod 2 = 0 Then	strColor = "#E7EBF5"
		
			strRequeter=rsSrc("Requester")
			if rsSrc("Note")<>"" then strRequeter=strRequeter & "<br>Note: " &  rsSrc("Note")
			
			strAuthorizer1=rsSrc("Authoriser1") 				
			if (rsSrc("isAuthoriser1Approved")) then  
				strAuthorizer1 =strAuthorizer1 & " " & strYes
			elseif Cint(rsSrc("Status"))=2 then
				strAuthorizer1 =strAuthorizer1 & " " & strNo
			end if				
			if rsSrc("Authoriser1Note")<>"" then strAuthorizer1=strAuthorizer1 & "<br>Note: " &  rsSrc("Authoriser1Note")
			
			strAuthorizer2=rsSrc("Authoriser2")
			if strAuthorizer2<>"" then
				if (rsSrc("isAuthoriser2Approved")) then  
					strAuthorizer2 =strAuthorizer2 & " " & strYes
				elseif rsSrc("isAuthoriser1Approved") and Cint(rsSrc("Status"))=2 then
					strAuthorizer2 =strAuthorizer2 & " " & strNo
				end if				
			end if				
			if rsSrc("Authoriser2Note")<>"" then strAuthorizer2=strAuthorizer2 & "<br>Note: " &  rsSrc("Authoriser2Note")
			
			strHR=""
			
			if rsSrc("isAuthorisedByHr") then 
				strHR="Need HR Authorise"
				
				if Cint(rsSrc("Status"))=3 and rsSrc("isHrApproved") then strHR = "Approve " & " " & strYes
				if Cint(rsSrc("Status"))=2 then
					if (rsSrc("isAuthoriser1Approved")) AND (rsSrc("Authoriser2_ID")=0 OR (rsSrc("Authoriser2_ID")<>0 AND rsSrc("isAuthoriser2Approved"))) then
						strHR="Rejected" & " " & strNo						
					end if
				end if
				if Cint(rsSrc("Status"))=1 then
					if (rsSrc("isAuthoriser1Approved")) AND (rsSrc("Authoriser2_ID")=0 OR (rsSrc("Authoriser2_ID")<>0 AND rsSrc("isAuthoriser2Approved"))) then
						strHR="In-progress"
					end if
				end if
				
				if rsSrc("HrNote") <>"" then strHR=strHR & "<br>Note: " &  rsSrc("HrNote")
				 
				
			end if				
						
			strStatus=arrStatus(rsSrc("Status"))
		
			if Cint(rsSrc("Status"))=2 then
				if  not rsSrc("isAuthoriser1Approved") then
					strStatus = strStatus & " by " & rsSrc("Authoriser1")
				elseif rsSrc("Authoriser2_Id")<>0  AND not rsSrc("isAuthoriser2Approved") then
					strStatus = strStatus & " by " & rsSrc("Authoriser2")
				elseif rsSrc("isAuthorisedByHr")  AND not rsSrc("isHrApproved") then
					strStatus = strStatus & " by HR"
				end if
				strNextStep="Done"
			else
				if not rsSrc("isAuthoriser1Approved") then
					 strNextStep=rsSrc("Authoriser1") 						 
				elseif rsSrc("Authoriser2_Id")<>0  AND not rsSrc("isAuthoriser2Approved") then
					 strNextStep=rsSrc("Authoriser2")
				elseif rsSrc("isAuthorisedByHr")  AND not rsSrc("isHrApproved") then
					strNextStep="HR Authorise"
				else
					 strNextStep="Done"
				end if
			end if
		

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
					 "<td class='blue-normal'>" & strRequeter & "</td>" &_
					 "<td class='blue-normal'>" & day(rsSrc("DateSubmitted")) & "/" & month(rsSrc("DateSubmitted")) & "/" & year(rsSrc("DateSubmitted")) & " " & timevalue(rsSrc("DateSubmitted"))& "</td>" &_
					 "<td class='blue-normal'>" & day(rsSrc("DateFrom")) & "/" & month(rsSrc("DateFrom")) & "/" & year(rsSrc("DateFrom")) & " " & timevalue(rsSrc("DateFrom")) &"</td>" &_
					 "<td class='blue-normal'>"& day(rsSrc("DateTo")) & "/" & month(rsSrc("DateTo")) & "/" & year(rsSrc("DateTo")) & " " & timevalue(rsSrc("DateTo")) &"</td>" &_			         
					 "<td class='blue-normal'>"& rsSrc("EventName") & "</td>" &_
					 "<td class='blue-normal'>" & strAuthorizer1  & "</td>" &_
					 "<td class='blue-normal'>" & strAuthorizer2  & "</td>" &_
					"<td class='blue-normal'>" & strHR  & "</td>" &_
					 "<td class='blue-normal'>" & strStatus  & "</td>" &_
					 "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
		
	End If
	Outbody = strOut
End Function

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	strDepartment = Request.Form("lbdepartment")
	fgSort = Request.Form("S")
	
	intCurPage = trim(Request.Form("P"))
	If intCurPage = "" Then
		intCurPage = 1
	End If		
	strName = Request.Form("name")
	intDepart = Request.Form("depart")
	
'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
	
'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

	If Not isEmpty(session("Preferences")) Then
		arrPre = session("Preferences")
		If arrPre(1, 0)>0 Then intPageSize = arrPre(1, 0) Else intPageSize = PageSizeDefault
		Set arrPre = Nothing
	Else
		intPageSize = PageSizeDefault
	End If

'--------------------------------------------------
' Check ACCESS right
'--------------------------------------------------

	strTemp = Request.ServerVariables("URL") 
	While Instr(strTemp, "/")<>0
		strTemp = Mid(strTemp, Instr(strTemp, "/") + 1, Len(strTemp))
	Wend
	
	strFilename = strTemp
	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = strTemp Then
				fgRight=True
				Exit For
			End If
		Next
		
		Set varGetRight = Nothing		
	End If	
	If fgRight = False Then		
		Response.Redirect("../../welcome.asp")
	End If

'--------------------------------------------------
' Check VIEWALL right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "view all" Then
				fgRight = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Initialize department array
'--------------------------------------------------
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	If isEmpty(session("varDepartment")) = False Then
		varDepartment = session("varDepartment")
		intNum = Ubound(varDepartment,2)
	Else
		If objDatabase.dbConnect(strConnect) Then			
			strSQL = "SELECT * FROM ATC_Department WHERE fgActivate=1 ORDER BY Department"

			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					varDepartment = objDatabase.rsElement.GetRows
					intNum = Ubound(varDepartment,2)					
					session("varDepartment") = varDepartment
					objDatabase.closeRec
				End If
			Else
				Response.Write objDatabase.strMessage
			End If
		Else
			Response.Write objDatabase.strMessage		
		End If
	End If	

'--------------------------------------------------
' End Of initializing department array
'--------------------------------------------------

'--------------------------------------------------
' Analyse query and prepare staff list
'--------------------------------------------------

	strAct = Request.QueryString("act")
	If strAct = "" Then
		strAct = Request.Form("txtstatus")
	End If

						 
	strSQL="SELECT b.FirstName + ' ' + b.LastName AS Requester, a.Note, c.FirstName + ' ' + c.LastName AS Authoriser1, Authoriser1Note,  " & _
				 "d.FirstName + ' ' + d.LastName AS Authoriser2,Authoriser2Note , a.Id, e.EventName,  " & _
				 "a.StaffId, a.Authoriser1_Id, ISNULL(a.Authoriser2_Id, 0) AS Authoriser2_Id, a.DateFrom, a.DateTo,  " & _
				 "a.isAuthorisedByHr, a.isAuthoriser1Approved, ISNULL(a.isAuthoriser2Approved, 0) AS isAuthoriser2Approved,  " & _
				 "ISNULL(a.isHrApproved, 0) AS isHrApproved, a.Status, a.HrNote,DateSubmitted " & _
			"FROM ATC_AbsenceRequests AS a  " & _
			"INNER JOIN ATC_Events e ON a.Type=e.EventID " & _
			"INNER JOIN ATC_PersonalInfo AS b ON a.StaffId = b.PersonID  " & _
			"INNER JOIN ATC_PersonalInfo AS c ON a.Authoriser1_Id = c.PersonID  " & _
			"LEFT OUTER JOIN ATC_PersonalInfo AS d ON a.Authoriser2_Id = d.PersonID "
	
	strSQL="SELECT b.FirstName + ' ' + b.LastName AS Requester, a.DateSubmitted, a.DateFrom, a.DateTo, e.EventName, c.FirstName + ' ' + c.LastName AS Authoriser1,   " & _
                         "d.FirstName + ' ' + d.LastName AS Authoriser2,ISNULL(a.isHrApproved, 0) AS isHrApproved, a.Status, a.Authoriser2Note, a.Authoriser1Note, a.Id, a.StaffId, a.Authoriser1_Id, ISNULL(a.Authoriser2_Id, 0) AS Authoriser2_Id, a.Note, a.isAuthorisedByHr,   " & _
                         "a.isAuthoriser1Approved, ISNULL(a.isAuthoriser2Approved, 0) AS isAuthoriser2Approved, a.HrNote  " & _
			"FROM            ATC_AbsenceRequests AS a INNER JOIN  " & _
                        "ATC_Events AS e ON a.Type = e.EventID INNER JOIN  " & _
                        "ATC_PersonalInfo AS b ON a.StaffId = b.PersonID INNER JOIN  " & _
                        "ATC_PersonalInfo AS c ON a.Authoriser1_Id = c.PersonID LEFT OUTER JOIN  " & _
                        "ATC_PersonalInfo AS d ON a.Authoriser2_Id = d.PersonID"
	response.write strSQL					
	strStatus=Request.Form("lstStatus")
	strSearchName=Request.Form("txtsearch")
	strSearch=""
	if trim(strSearchName)<>"" then
		strPrefix=Request.Form("lstType")
		strSearch= " " & strPrefix & ".FirstName + ' ' + " & strPrefix & ".LastName  like '" & strSearchName & "%'"
	end if
	
	if trim(strStatus)<>"" then
		if strSearch<>"" then strSearch=strSearch & " AND "
		strSearch=strSearch & " Status =" & strStatus
	end if
	
	IF strSearch<>"" then strSearch=" WHERE " & strSearch
	
	strSQL=strSQL & strSearch 
	
'--------------------------------------------------
'Sort
'--------------------------------------------------
    intSortColum = request.Form("txtsortcol")
    if intSortColum="" then intSortColum="a.id"

    intSortType= request.Form("txtsorttype")
    if intSortType="" then intSortType=2
	
	'rsData.movefirst
      
    'rsData.Sort=rsData(cint(intSortColum)).Name& " " & arrSortType(cint(intSortType)-1)  
	strSQL= strSQL & " ORDER BY "  & intSortColum & " " & arrSortType(cint(intSortType)-1)
	
	'response.write strSQL
	'response.end

	Call GetRecordset(strSQL,rsData)
	
'--------------------------------------------------
'Start Paging
'--------------------------------------------------

' Set the PageSize, CacheSize and populate the intPageCount

	rsData.PageSize=intPageSize

' The Cachesize property sets the number of records that will be cached locally in memory	
	rsData.CacheSize=rsData.PageSize	
	intPageCount=rsData.PageCount
	intRecordCount=rsData.RecordCount

	' Checking to make sure that we are not before the start or beyond end of the recordset
' If we are beyond the end, set the current page equal the last page of the recordset.
' If we are before the start, set the current page equal the start of the recordset
	
	intPage=Request.QueryString("Navi")

	if intPage="" then intPage=1
	
	if cint(intPage)>Cint(intPageCount) then intPage=intPageCount
	if cint(intPage)<=0 then intPage=1


	strLast=Outbody(rsData,intPage,intPageSize)
	
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
	If strChoseMenu = "" Then strChoseMenu = "AE"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta charset="UTF-8">
<title><%=webname%></title>

<link rel="stylesheet" type="text/css" href="../../jQuery/jquery-ui.css"/>
<link rel="stylesheet" href="../../timesheet.css"/>
<link href="../../jQuery/tablestyle.css" rel="stylesheet" type="text/css" />

   
<style type="text/css">
.ui-widget
    {
        font-family: Arial, Verdana;
        font-size: 10pt;
    }
.submit ul
{
      padding:0;
      margin:0;
}

.submit ul li
{       
    display:inline;
}

.linkbutton
{    
    margin-left:2px;
    padding: 2px 10px 4px 10px;
    width:75px;
    height:17px;
    background-color:#8CA0D1;
    text-align:center;
    font-weight: bold;
    text-decoration:none;
}

.submit a:hover
{
    background-color:#7791D1;
    color:white;
}
ol 
{
	margin:0px;
	padding:10px;
	width:236px;
	list-style:none; 
	background-color:#C0CAE6; 
}

.highlight { background-color: yellow }
        
</style>


</head>

<body>
<form name="navi" id="navi" method="post" action="listofemployee.asp" class="submit"> 
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			Response.Write(arrTmp(0))
			%>
          <tr > 
            <td style="padding:20px 0 0 20px;" align="center"> 
            <%if gMessage<>"" then%>
               <div style="font-weight:bold; height:20px; background-color:#E7EBF5;" class="red"><%=gMessage%></div>
            <%end if%>
            
				
				<ul>          
                    <li style="padding-left:120px">
                        <input type="text" id="txtsearch" name="txtsearch" class="blue-normal" size="15" style="width:300px" value="<%'=Showvalue(varSearch)%>"/>
                    </li>
					<li >
                        <select name="lstType" height="26px" width="70px" class="blue-normal">
							<option value="b">Requestor</option>
							<option value="c">Authoriser 1</option>
							<option value="d">Authoriser 2</option>
						</select>
                     
                    </li>
                    <li class="linkbutton"><a href="#" class="b" id="lnkSearch">Search</a></li>
                </ul>
				<ul>
					<li>
						<div class="title" style="padding:20px;">List of booking request</div>
					</li>
				</ul>
                <ul style="padding-bottom:20px">
          
                    <li style="padding-left:120px">Status</li>
					<li >
						<select name="lstStatus" class="blue-normal" id="lstStatus">
							<option value=""></option>
							<option value="0">New</option>
							<option value="1">In-Progress</option>
							<option value="3">Authorised</option>
							<option value="2">Rejected</option>
						</select>
						
					</li>
					
                </ul>     

                <input type="hidden" name="txtsortcol" value="<%=intSortColum %>"/>
                <input type="hidden" name="txtsorttype" value="<%=intSortType %>"/>
			
            </td>
          </tr>
          <tr> 
            <td height="100%" style="vertical-align:top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:600px" >
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table id="tblList" class="tablesorter">
                            <thead>
                                <tr>
									<th scope="col" id="Requester" width="15%">Requester</th>
									<th scope="col" id="DateSubmitted" width="7%">Submitted<br> Date</th>
									<th scope="col" id="DateFrom" width="7%">From</th>
									<th scope="col" id="DateTo" width="7%">To</th>
									<th scope="col" id="EventName" width="10%">Type Of Absence</th>
									<th scope="col" id="Authoriser1" width="15%">Authorizer1</th>
									<th scope="col" id="Authoriser2" width="15%">Authorizer2</th>
									<th scope="col" id="isAuthorisedByHr" width="14%">HR Authorizer </th>
									<th scope="col" id="Status" width="10%">Status</th>
                                </tr>
                            </thead>
                            <tbody>
                            <%
	                            Response.Write(strLast)
                            %>     
                            </tbody>                       
                          </table>
						  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal">
                                &nbsp;&nbsp;*Click on each column header to sort the list by alphabetical order.</td>
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
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
</form>			            
<script type="text/javascript" language="javascript" src="../../library/menu.js"></script>
<script type="text/javascript" src="../../library/library.js"></script>

<script type="text/javascript" src="../../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../../jQuery/jquery-ui.min.js"></script>
<script type="text/javascript">
    
$(document).ready(function() {
       
	//Searching
	$("#lnkSearch").click(function() {
	
		search();		
	});
    
	$( "#lstStatus" ).change(function() {
		search();
	});
	
	//For sort
	  $("#tblList thead th").click(function(){
			sort($(this).attr("id"));
			
	  })
})

function search() {
	
	document.navi.action = "annual_list_request.asp"
	document.navi.target = "_self";
	document.navi.submit();
}

function next() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage < numpage) {
	
		curpage=<%=intPage+1%>
		document.navi.action = "annual_list_request.asp?navi=" + curpage;
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function prev() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage > 1) {
		curpage=<%=intPage-1%>
		document.navi.action = "annual_list_request.asp?navi=" + curpage;
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function go() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	var intpage = document.navi.txtpage.value;
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.navi.action = "annual_list_request.asp?navi=" + intpage;
		document.navi.target = "_self";
		document.navi.submit();		
	}
}

function sort(col) {
    
    var curSortCol;
    var curSortType;

    curSortCol=document.navi.txtsortcol.value;    
    curSortType=parseInt(document.navi.txtsorttype.value);
    
    document.navi.txtsortcol.value=col;
    
    if (col!=curSortCol)
    {
        document.navi.txtsorttype.value=1;
    }
    else
    {
        document.navi.txtsorttype.value=1;
        if (curSortType==1)
            document.navi.txtsorttype.value=2;
    }
        
	document.navi.action = "annual_list_request.asp";
	document.navi.target = "_self";
	document.navi.submit();
}

</script>

</body>
</html>