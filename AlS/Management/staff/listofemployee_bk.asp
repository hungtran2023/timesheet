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
			
			strOut = strOut & "<tr " & strClass & " idValue=" & rsSrc("PersonID") & " >" &_
			         "<td>" & Showlabel(rsSrc("Fullname")) & "</td>" &_
			         "<td>" & Showlabel(rsSrc("JobTitle")) & "</td>" &_
			         "<td>" & Showlabel(rsSrc("LeaderName")) & "</td>" &_
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
    dim strSearchName,strSearchJobtitle,strSearchDepartment,strSearchReportTo
    
    strSearchName=Request.Form("txtSearch")
    if Request.Form("txtStaffID")<>"" AND Request.Form("txtStaffID")<>"-1" then 
        strSearchName="='" & strSearchName & "'"
    else
        strSearchName="like '%" & strSearchName & "%'"
    end if
    
    strSearchJobtitle=Request.Form("txtJobtitle")
    if Request.Form("txtJobtitleID")<>"" AND Request.Form("txtJobtitleID")<>"-1" then 
        strSearchJobtitle="='" & strSearchJobtitle & "'"
    else
        strSearchJobtitle="like '%" & strSearchJobtitle & "%'"
    end if
    
    strSearchDepartment=Request.Form("txtDepartment")   
    
    if Request.Form("txtDepartmentID")<>"" AND Request.Form("txtDepartmentID")<>"-1" then 
        strSearchDepartment="='" & strSearchDepartment & "'"
    else
        strSearchDepartment="like '%" & strSearchDepartment & "%'"
    end if
    
    strSearchReportTo=Request.Form("txtReportTo")
    
    strSearch=" WHERE Jobtitle " & strSearchJobtitle &_
                    " AND Department " & strSearchDepartment & _
                    " AND Fullname " & strSearchName 
                    
    if strSearchReportTo <>"" then strSearch=strSearch & " AND (FirstNameLeader + ' ' + LastnameLeader) like '%" & strSearchReportTo & "%'"

    SearchPhrase=strSearch
    
end function
'------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage, PageSize, fgUpdate, fgRight
	dim arrSortType
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
	if strChoseMenu = "" then strChoseMenu = "AB"
	
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
    
    strQuery="SELECT PersonID,Fullname,JobTitle, (FirstNameLeader + ' ' + LastnameLeader) as LeaderName FROM HR_Employee" & SearchPhrase()
    		
    if Not fgRight then strQuery= strQuery & " AND PersonID IN (SELECT StaffID FROM UserByReportTo(" & session("USERID") & "))"		  
	
'response.write strQuery
    
'--------------------------------------------------
'For searching
'--------------------------------------------------    
'response.write strQuery

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
      
    rsData.Sort=rsData(cint(intSortColum)).Name& " " & arrSortType(cint(intSortType)-1)  

'--------------------------------------------------
'Generate data
'--------------------------------------------------
    strLast=Outbody(rsData,intPage,PageSize)
'--------------------------------------------------
'Get JSON for autocomplete
'--------------------------------------------------
 
        strOut3=getArrJSON("SELECT * FROM ATC_JobTitle WHERE fgActivate=1 ORDER BY JobTitle")		
		
		strOut4 = getArrJSON("SELECT * FROM ATC_Department WHERE fgActivate=1 ORDER BY Department")
			
		strQuery = "SELECT DISTINCT a.DirectLeaderID, e.Firstname + ' ' + ISNULL(e.LastName, '') + ' ' + ISNULL(e.MiddleName, '') as Fullname  " & _
                "FROM ATC_Employees a LEFT JOIN ATC_PersonalInfo e ON a.DirectLeaderID = e.PersonID " & _
                "WHERE a.Leavedate IS NULL AND a.DirectLeaderID IS NOT NULL AND " & _
	                "a.DirectLeaderID NOT IN (SELECT PersonID FROM ATC_PersonalInfo  " & _
                "WHERE e.fgDelete = 1) ORDER BY Fullname"
				
'response.write strQuery				
                
        strOut5 =getArrJSON(strQuery) 

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
<meta http-equiv="X-UA-Compatible" content="IE=edge;chrome=1" />

<head>
<title><%=webname%></title>

<link rel="stylesheet" type="text/css" href="../../jQuery/jquery-ui.css"/>
<link rel="stylesheet" href="../../timesheet.css"/>
<link href="../../jQuery/tablestyle.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" language="javascript" src="../../library/menu.js"></script>
<script type="text/javascript" src="../../library/library.js"></script>

<script type="text/javascript" src="../../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../../jQuery/jquery-ui.min.js"></script>

    
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

<script type="text/javascript">
    var arrJobtitle=[<%=strOut3%>];
    var arrDepart=[<%=strOut4%>];
    var arrReport=[<%=strOut5%>];
    var arrHeaderClass=["headerSortUp","headerSortDown"];

$(document).ready(function() {
        
        $("#<%=intSortColum%>").addClass(arrHeaderClass[<%=intsorttype-1%>]);
        
        //Autocomplete for search fulname
        $( "#txtsearch" ).autocomplete({
            source: "../../inc/script.asp",
            minLength: 2,
            change:function(event, ui) {
                if (ui.item==null)
                {
                    $('#txtstaffID').val(-1); 
                 }
            },
            select: function(event, ui) { 
                 $('#txtstaffID').val(ui.item.id);
            }
         })
         //Option
         $("#filter").toggle();
         
         $("#lnkOption").click(function() {
            if (!($("#filter").is(":hidden"))) {
                CleanFilterForm(); }         
                var position = $("#txtsearch").position();   
            $("#filter").css("left",position.left-430);
            $("#filter").toggle();
        });
        //Searching
        $("#lnkSearch").click(function() {
            search();
              });
        //Autocomplete for filter
        $('#txtReportTo').autocomplete({
            source: arrReport,
            change:function(event, ui) {
                if (ui.item==null)
                    $('#txtReportToID').val(-1); 
            },
            select: function(event, ui) { 
                 $('#txtReportToID').val(ui.item.id); 
                }
            });
        $('#txtDepartment').autocomplete({
            source: arrDepart,
            change:function(event, ui) {
                if (ui.item==null)
                    $('#txtDepartmentID').val(-1); 
            },
            select: function(event, ui) { 
                 $('#txtDepartmentID').val(ui.item.id); 
                }
            });
        $('#txtJobtitle').autocomplete({
            source: arrJobtitle,
            change:function(event, ui) {
                if (ui.item==null)
                    $('#txtJobtitleID').val(-1); 
            },
            select: function(event, ui) { 
                 $('#txtJobtitleID').val(ui.item.id); 
                }
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
		document.navi.action = "listofemployee.asp?navi=" + curpage;
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function prev() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage > 1) {
		curpage=<%=intPage-1%>
		document.navi.action = "listofemployee.asp?navi=" + curpage;
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
		document.navi.action = "listofemployee.asp?navi=" + intpage;
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
        
	document.navi.action = "listofemployee.asp";
	document.navi.target = "_self";
	document.navi.submit();
}

function search() {

	document.navi.action = "listofemployee.asp"
	document.navi.target = "_self";
	document.navi.submit();
}

function onLoad(){
    if (document.layers) {
        _loadmenu();
    }
}

function getdetail(varid){
	document.navi.txtuserid.value = varid;
	
	document.navi.action = "employeeProfile.asp";
	document.navi.target = "_self";
	document.navi.submit();
}

function addnew(){
    document.navi.txtuserid.value = -1;
	document.navi.action = "employeeProfile.asp";
	document.navi.target = "_self";
	document.navi.submit();
}

function CleanFilterForm(){
	$('#txtDepartment').val("");
	$('#txtJobtitle').val("");
	$('#txtReportTo').val("");
	
	$('#txtDepartmentID').val("-1");
	$('#txtJobtitleID').val("-1");
	$('#txtReportToID').val("-1");
}
</script>
</head>

<body>
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
            <form name="navi" id="navi" method="post" action="listofemployee.asp" class="submit"> 
                <ul>
<%if fgUpdate then%>
                    <li><a class="blue" href="javascript:addnew();"  onmouseover="self.status='Add a new employee'; return true;" onmouseout="self.status=''">Add New</a></li>
<%end if %>            
                    <li style="padding-left:120px">
                        <input type="text" id="txtsearch" name="txtsearch" class="blue-normal" size="15" style="width:250px" value="<%=Showvalue(varSearch)%>"/>
                        <input type="hidden" id="txtstaffID" name="txtstaffID" value="<%=intreport%>" />
                     </li>
                    <li class="linkbutton"><a href="#" class="b" id="lnkSearch">Search</a></li>
                    <li class="linkbutton"><a href="#" class="b" id="lnkOption">Option >></a></li>
                </ul>
                <div id="filter" style="position:relative;">
                    <ol>
                       <li><label>Job Title
                            <input type="text" id="txtJobtitle" name="txtJobtitle" style="width:230px" value="<%=trim(strJobTitle)%>" class="blue-normal" /></label>
                            <input type="hidden" id="txtJobtitleID" name="txtJobtitleID" value="<%=intjobtitle%>" />
                            </li>
                        <li><label for="txtDepartment">Department
                            <input type="text" id="txtDepartment" name="txtDepartment" style="width:230px" value="<%=trim(strDepartment)%>" class="blue-normal"/></label>
                            <input type="hidden" id="txtDepartmentID" name="txtDepartmentID" value="<%=intdepartment%>" />
                            </li>
                        <li><label for="txtReport">Reports To
                            <input type="text" id="txtReportTo" name="txtReportTo" style="width:230px" value="<%=trim(strReportTo)%>" class="blue-normal"/></label>
                            <input type="hidden" id="txtReportToID" name="txtReportToID" value="<%=intreport%>" />
                        </li>
                    </ol>
                </div>                         
                <input type="hidden" name="txtuserid" value=""/>
                <input type="hidden" name="txtpreviouspage" value="<%=strFilename%>"/>  
                <input type="hidden" name="txtsortcol" value="<%=intSortColum %>"/>
                <input type="hidden" name="txtsorttype" value="<%=intSortType %>"/>
            </form>
            
            <div class="title" style="padding:10px; text-align:center;">List of Employees</div>
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
                                    <th scope="col" id="1" width="33%">Full Name</th>
                                    <th scope="col" id="2" width="33%">Job Title</th>
                                    <th scope="col" id="3" width="34%">Report To</th>
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
    var hotkey = 13
    if (document.layers)
        document.captureEvents(Event.KEYPRESS)
    function backhome(e) {
        if (document.layers) {
            if (e.which == hotkey)
                search();
        }
        else if (document.all) {
            if (event.keyCode == hotkey) {
                event.keyCode = 0;
                search();
            }
        }
    }
    document.onkeypress = backhome
</script>
</body>
</html>