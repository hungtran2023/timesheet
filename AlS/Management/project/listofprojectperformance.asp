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
function Outbody(ByRef rsSrc, ByVal intPage,ByVal PageSize)
	dim intStart,intFinish
	strOut = ""

	if not rsSrc.EOF then
		
		rsSrc.AbsolutePage = intPage
		intStart = rsSrc.AbsolutePosition
		
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsData.PageSize - 1)
		End if
	
	    strOut = ""
	    For i = intStart to intFinish
			strClass=""
			if i mod 2 = 0 then	strClass = "class='odd'"
			
			strOut = strOut & "<tr " & strClass & " idValue=" & rsSrc("ProjectID") & " >" &_
			         "<td>" & Showlabel(rsSrc("ProjectID")) & "</td>" &_
			         "<td>" & Showlabel(rsSrc("ProjectName")) & "</td>" &_
			         "<td>" & formatnumber(rsSrc("CSOHours"),2) & "</td>" &_
			         "<td>" & formatnumber(rsSrc("ActualHours"),2) & "</td>" &_
			         "<td>" & formatnumber(rsSrc("CSOPayment"),2) & "</td>" &_
			         "<td>" & formatnumber(rsSrc("Sales"),2) & "</td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function
'****************************************
' Function: SearchPhrase
' Description: 
'			  
' Return value: search phase base on filter condition
' Author: 
' Date: 
' Note:
'****************************************
function SearchPhrase()
    dim  strSearch
    
    varSearchType = Request.Form("lbSeachType")
	if varSearchType="" then varSearchType=3
    
    strSearchName=Request.Form("txtSearch")
    strSearch=""
		
	if cint(varSearchType)=1 then
		strSearch= strSearch & " ProjectName Like '%" & strSearchName & "%'"
	elseif cint(varSearchType)=2 then
		strSearch= strSearch & " b.ProjectID Like '%" & strSearchName & "%'"
	elseif cint(varSearchType)=3 then
		strSearch= strSearch & " (b.ProjectID Like '%" & strSearchName & "%' OR ProjectName Like '%" & strSearchName & "%')"
	end if

    if strSearch<>"" then strSearch= "WHERE " & strSearch
    SearchPhrase=strSearch
    
end function
'------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage, PageSize, fgUpdate, fgRight
	dim arrSortType, varSearchType,strSearchName
	arrSortType=array("ASC","DESC")

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
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if

'-----------------------------------
'Check VIEWALL right
'-----------------------------------
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = "View all projects" then
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
'-------------------------------
' Get Fullname and Job Title
'-------------------------------
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	
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
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

    gMessage = ""

    if Request.QueryString("fgMenu") <> "" then
	    fgExecute = false
    else
	    fgExecute = true
	    if Request.TotalBytes=0 or Request.QueryString("outside")<>"" then
	    end if
    end if
    
    strQuery="SELECT A.ProjectID, ProjectName, SUM(Hours+OTHours) as ActualHours, SUM(InvoiceValue) as Sales, SUM(CSOHours) as CSOHOurs, SUM(CSOPayment) as CSOPayment " & _
              "FROM rp_ProjectPerformanceByPeriod A INNER JOIN ATC_Projects B ON A.ProjectID=B.ProjectID " & _
                SearchPhrase() &" GROUP BY A.ProjectID,B.ProjectName ORDER BY A.ProjectID" 
  
    		
'    if Not fgRight then strQuery= strQuery & " AND PersonID IN (SELECT StaffID FROM UserByReportTo(" & session("USERID") & "))"		  
    
'--------------------------------------------------
'For searching
'--------------------------------------------------    

    Call GetRecordset(strQuery ,rsData)		
    
'--------------------------------------------------
'Start Paging
'--------------------------------------------------

' Set the PageSize, CacheSize and populate the intPageCount

	rsData.PageSize=PageSize
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

'--------------------------------------------------
'Sort
'--------------------------------------------------
    intSortColum = request.Form("txtsortcol")
    if intSortColum="" then intSortColum=1

    intSortType= request.Form("txtsorttype")
    if intSortType="" then intSortType=1
    
    'rsData.Sort=rsData(cint(intSortColum)).Name& " " & arrSortType(cint(intSortType)-1)  

'--------------------------------------------------
'Generate data
'--------------------------------------------------
    strLast=Outbody(rsData,intPage,PageSize)

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
<title><%=webname%></title>

<link rel="stylesheet" href="../../timesheet.css"/>
<link rel="stylesheet" type="text/css" href="../../jQuery/jquery-ui.css"/>
<link href="../../jQuery/atlasJquery.css" rel="stylesheet" type="text/css" />
<link href="../../jQuery/tablestyle.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" language="javascript" src="../../library/menu.js"></script>
<script type="text/javascript" src="../../library/library.js"></script>

<script type="text/javascript" src="../../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../../jQuery/jquery-ui.min.js"></script>

    
<style type="text/css">

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

<script type="text/javascript">

    var arrHeaderClass=["headerSortUp","headerSortDown"];

$(document).ready(function() {
        
        $("#<%=intSortColum%>").addClass(arrHeaderClass[<%=intsorttype-1%>]);
        
        
        //Searching
        $("#lnkSearch").click(function() {
            search();
              });
        
          //For edit user
          $("#tblList tbody tr").click(function(){
                getdetail($(this).attr("idValue"));
          })
          
          //For sort
          $("#tblList thead th").click(function(){
                sort($(this).attr("id"));
                
          })
    })

function next() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage < numpage) {
	
		curpage=<%=intPage+1%>
		document.navi.action = "listofprojectperformance.asp?navi=" + curpage;
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function prev() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage > 1) {
		curpage=<%=intPage-1%>
		document.navi.action = "listofprojectperformance.asp?navi=" + curpage;
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
		document.navi.action = "listofprojectperformance.asp?navi=" + intpage;
		document.navi.target = "_self";
		document.navi.submit();		
	}
}

function sort(col) {
    
    var curSortCol;
    var curSortType;
    
    curSortCol=parseInt(document.navi.txtsortcol.value);    
    curSortType=parseInt(document.navi.txtsorttype.value);
    
    document.navi.txtsortcol.value=col;
    
    if (parseInt(col)!=curSortCol)
    {
        document.navi.txtsorttype.value=1;
    }
    else
    {
        document.navi.txtsorttype.value=1;
        if (curSortType==1)
            document.navi.txtsorttype.value=2;
    }
        
	document.navi.action = "listofprojectperformance.asp";
	document.navi.target = "_self";
	document.navi.submit();
}

function search() {

	document.navi.action = "listofprojectperformance.asp"
	document.navi.target = "_self";
	document.navi.submit();
}

function onLoad(){
    if (document.layers) {
        _loadmenu();
    }
}

function getdetail(varid){
	document.navi.txthidden.value = varid;
	
	document.navi.action = "Pro_PerformmanceDetails.asp?act=EDIT";
	document.navi.target = "_self";
	document.navi.submit();
}



</script>
</head>

<body>
<form name="navi" id="navi" method="post" class="submit"> 
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			Response.Write(arrTmp(0))
			%>
          <tr> 
            <td style="padding:20px 0 0 20px;"> 
            <%if gMessage<>"" then%>
               <div style="font-weight:bold; height:20px; background-color:#E7EBF5;" class="red"><%=gMessage%></div>
            <%end if%>
            
                <ul>
                     <li style="padding-left:120px">
                        <input type="text" id="txtsearch" name="txtsearch" class="blue-normal" size="15" style="width:250px" value="<%=strSearchName%>"/>
                     </li>
                     <li >
                     <select name="lbSeachType" class='blue-normal' >
							<option value="1"<%if cint(varSearchType)=1 then %>selected<%end if%>>Project Name</option>
							<option value="2"<%if cint(varSearchType)=2 then %>selected<%end if%>>APK</option>
							<option value="3"<%if cint(varSearchType)=3 then %>selected<%end if%>>Project Name &amp; APK</option>
						</select></li>
                    <li class="linkbutton"><a href="#" class="b" id="lnkSearch">Search</a></li>
                    
                </ul>
                        
                <input type="hidden" name="txthidden" value=""/>
                <input type="hidden" name="txtpreviouspage" value="<%=strFilename%>"/>  
                <input type="hidden" name="txtsortcol" value="<%=intSortColum %>"/>
                <input type="hidden" name="txtsorttype" value="<%=intSortType %>"/>
            
            
            <div class="title" style="padding:10px; text-align:center;">Project Performance</div>
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
                                    <th scope="col" id="1" width="14%">APK</th>
                                    <th scope="col" id="2" width="30%">Project Name</th>
                                    <th scope="col" id="3" width="14%">CSO Hours</th>
                                    <th scope="col" id="4" width="14%">Actual Hours</th>
                                    <th scope="col" id="5" width="14%">Payment Schedule</th>
                                    <th scope="col" id="6" width="14%">Invoiced</th>
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

<script type="text/javascript">
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
</script>
</form>
</body>
</html>