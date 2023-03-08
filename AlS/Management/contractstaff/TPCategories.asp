<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim strPCCode,strSupplier,	dateBuying,strNote,strCPU,strRAM,strSysMem,strVideo,intStatus
	dim intPCID,intDetailPCID
	Dim objEmployee, objDatabase, strError,rsData
	Dim arrlstFrom(2),arrlongmon,arrCategories,intCategoryType
	
	arrCategories =Array("Contract Level", "Overtime rule")
	
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
			
			strActivate="<img src='../../images/yes.gif'>"
			if not rsSrc("fgActivate") then strActivate=""
			
			
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & i+1 & "</td>"
			strOut=strOut & "<td valign='top' class='blue'>" & _
						"<a href='javascript:CategoryDetail(" & rsSrc("CategoryID") & ");' " &_
						"class='c'>" & rsSrc("CategoryName") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("CategoryNote") & "</td>"
			strOut=strOut & "<td valign='top' align='center' class='blue-normal'>" & strActivate & "</td>"
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
function ExecuteSQL(strSql,prefix)

	dim strConnect,ret,strMessage
	dim objDb	

	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
		
	If objDb.dbConnect(strConnect) then
			
		ret = objDb.runActionQuery(strSql)
				
		if ret=false then				
			strMessage = objDb.strMessage
		else
			strMessage=prefix & " successfully."
		end if
			  
	else
		strMessage=objDb.strMessage
	end if
	
	ExecuteSQL=strMessage
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
	intCategoryType=Request.Form("txtCategoryType")
	if intCategoryType="" then intCategoryType=0	
	strTile=arrCategories(intCategoryType)
	
		
	intPCID = Request.Form("txtID")
	fgDel=Request.Form("fgstatus")

	if Request.QueryString("act") = "save" then
		intCategoryID=Request.Form("txtCategoryID")
		
		if fgDel<>"D" then
			strName=Request.Form("txtName")
			strNote=Request.Form("txtNote")
			fgActivate=Request.Form("optActivate")
			if fgActivate="" then fgActivate=0
		
            Select Case intCategoryType
		    Case 0
		        intLevelRate=Request.Form("txtLevelRate")
		    case 1
		        intNormalRate =Request.Form("txtNormal")
		        if intNormalRate="" then intNormalRate=1
		        intNightRate=Request.Form("txtNight")
		        if intNightRate="" then intNightRate=1
		        intWeekendRate=Request.Form("txtWeekend")
		        if intWeekendRate="" then intWeekendRate=1
	        end select		
		
			if cint(intCategoryID)=-1 then
				Select Case intCategoryType
					Case 0
					  strSql="INSERT INTO ATC_TPLevel (LevelName,Note,fgActivate,Rate) VALUES('" & _
								Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & "," & intLevelRate & ")"
					case 1
					  strSql= "INSERT INTO ATC_TPOverTime (OTName , Note , fgActivate, NormalRate , NightRate , WeekendRate ) VALUES('" & _
								Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate &_
								"," & intNormalRate & "," & intNightRate & "," & intWeekendRate  & ")"
				end select
				strPrefix="Add new "
			else
				Select Case intCategoryType
					Case 0
					  strSql="UPDATE ATC_TPLevel SET " & _
								"LevelName = '" & Replace(strName,"'","''") & "'" & _
								",Note = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
								",Rate =" & intLevelRate & _
							" WHERE LevelID=" & intCategoryID
					case 1				  
					  strSql="UPDATE ATC_TPOverTime SET " & _
								"OTName = '" & Replace(strName,"'","''") & "'" & _
								",Note = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
								",NormalRate = " & intNormalRate & _
								",NightRate = " & intNightRate & _
								",WeekendRate = " & intWeekendRate & _
								" WHERE OvertimeID= " & intCategoryID
				end select
				strPrefix="Update "
			end if
		else
			Select Case intCategoryType
				Case 0
				  strSql="DELETE ATC_TPLevel WHERE LevelID=" & intCategoryID
				case 1				  
				  strSql="DELETE ATC_TPOverTime WHERE OvertimeID= " & intCategoryID
			end select
			strPrefix="Delete "
			fgDel=""
		end if
		strError=ExecuteSQL(strSql,strPrefix)
	
	End If
	
'--------------------------------------------------
' 
'--------------------------------------------------
'arrCategories =Array("Computer Status", "Computer Types", "Software Categories", "Licence Types")	
	Select Case intCategoryType
		Case 0
		  strSql="SELECT LevelID as CategoryID,LevelName as CategoryName, Note as CategoryNote,fgActivate as fgActivate, rate  FROM ATC_TPLevel"
		case 1
		    strSql="SELECT  OvertimeID as CategoryID , OTName as CategoryName, Note as CategoryNote,fgActivate as fgActivate , NormalRate , NightRate , WeekendRate FROM  ATC_TPOverTime "
	end select

	Call GetRecordset(strSql,rsSrc)
	
	strLast=OutBody(rsSrc)	
	
	If Request.QueryString("act") = "EDIT" Then			
		if rsSrc.RecordCount>0 then
		
			intCategoryID=Request.Form("txtCategoryID")
			rsSrc.Filter = "CategoryID=" & intCategoryID
			
			strName=rsSrc("CategoryName")
			strNote=rsSrc("CategoryNote")
			fgActivate=rsSrc("fgActivate")
			
			Select Case intCategoryType
		    Case 0
		        intLevelRate=rsSrc("Rate")
		    case 1
		        intNormalRate =rsSrc("NormalRate")
		        intNightRate=rsSrc("NightRate")
		        intWeekendRate=rsSrc("WeekendRate")
	        end select
			
		end if			
	else
		intCategoryID=-1
		strName=""
		strNote=""
		fgActivate=false
		Select Case intCategoryType
	    Case 0
	        intLevelRate=""
	    case 1
	        intNormalRate =""
	        intNightRate=""
	        intWeekendRate=""
        end select
	end if
	
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
	If strChoseMenu = "" Then strChoseMenu = "AG"
	
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

<link rel="stylesheet" type="text/css" href="../../library/DropDownMenu/anylinkcssmenu.css" />
<script type="text/javascript" src="../../library/DropDownMenu/anylinkcssmenu.js"></script>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--

anylinkcssmenu.init("anchorclass")

function savedata()
{
	if (checkdata())
	{
		window.document.frmreport.action = "TPCategories.asp?act=save"			
		window.document.frmreport.submit();
	}
}

function Category(type)
{
	window.document.frmreport.txtCategoryType.value = type
	window.document.frmreport.action = "TPCategories.asp"			
	window.document.frmreport.submit();
}

function CategoryDetail(id)
{
	window.document.frmreport.txtCategoryID.value = id
	window.document.frmreport.action = "TPCategories.asp?act=EDIT"			
	window.document.frmreport.submit();
}

function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "TPCategories.asp?act=save"
	window.document.frmreport.submit();
}

function checkdata()
{
	if (window.document.frmreport.txtName.value=="")
	{
		alert("Please enter <%=strTile%> description.");
		document.frmreport.txtName.focus();
		return false	
	}	
	<%select case intCategoryType
    case 0%>
        
        if (window.document.frmreport.txtLevelRate.value=="")
	    {
		    alert("Please enter rate value.");
		    document.frmreport.txtName.focus();
		    return false	
	    }	
	    if (isNaN(document.frmreport.txtLevelRate.value)==true) {
			alert("The Rate value must be a number.");
			document.frmreport.txtLevelRate.focus();
			return false;
		}
		else if (document.frmreport.txtLevelRate.value<0) {
			alert("The Rate value must be greater than 0.");
			document.frmreport.txtLevelRate.focus();
			return false;			
		}
    <%  case 1%>
        if (isNaN(document.frmreport.txtNormal.value)||isNaN(document.frmreport.txtWeekend.value)||isNaN(document.frmreport.txtNight.value)) {
			alert("The Rate value must be a number.");
			document.frmreport.txtNormal.focus();
			return false;
		}
		else if ((document.frmreport.txtNormal.value<0)||(document.frmreport.txtWeekend.value<0)||(document.frmreport.txtNight.value<0)) {
			alert("The Rate value must be greater than 0.");
			document.frmreport.txtNormal.focus();
			return false;			
		}
        
    <%end select%>
	return true	
}

function Add() {
	document.frmreport.txtName.value = "";
	document.frmreport.txtNote.value = "";
	document.frmreport.optActivate.checked=false
	document.frmreport.txtCategoryID.value = -1
<%select case intCategoryType
    case 0%>
	    document.frmreport.txtLevelRate.value="";
<%  case 1%>
	    document.frmreport.txtNormal.value="";
	    document.frmreport.txtWeekend.value="";
	    document.frmreport.txtNight.value="";
<%end select%>
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
                  <td class="blue" height="10" align="left" width="23%"> </td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;</td>
                  <td class="blue" height="30" align="right" width="77%">
					<table width="120" border="0" cellspacing="2" cellpadding="0" align="right" height="20" name="aa">
                      <tr> 
                       
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue">
							<a href="javascript:Category(0)" class="anchorclass" rel="submenu1">Select Category</a>							
                        </td>						                     
                      </tr>                    
                    </table>
                  </td>
                </tr>                
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2"><%=strTile%></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign=top>
            
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="2">
                                 
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="20%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="15%">Description *</td>
                              <td valign="middle" width="45%" class="blue">
								<input type="text" name="txtName" maxlength="20" class="blue-normal" style="width:95%;" value="<%=strName%>"></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>          

<%
			Select Case intCategoryType
		    Case 0 %>
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="20%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="15%">Level Rate *</td>
                              <td valign="middle" width="45%" class="blue">
								<input type="text" name="txtLevelRate" maxlength="20" class="blue-normal" style="width:95%;" value="<%=intLevelRate%>"></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>   
<%          case 1 %>
                             <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >OT Rate</td>
                              <td valign="middle" class="blue">
                                <table width="95%" style="background-color:'#E7EBF5'; margin-left:0">
                                    <tr>
                                        <td align="center" width="33%" class="blue-small">Normal</td>
                                        <td align="center" width="34%" class="blue-small">Weekend</td>
                                        <td align="center" width="33%" class="blue-small">Night</td>
                                    </tr>
                                    <tr>
                                        <td align="left"><input type="text" name="txtNormal" maxlength="20" class="blue-normal" style="width:95%;" value="<%=intNormalRate%>"/></td>
                                        <td align="center"><input type="text" name="txtWeekend" maxlength="20" class="blue-normal" style="width:95%;" value="<%=intWeekendRate%>"/></td>
                                        <td align="right"><input type="text" name="txtNight" maxlength="20" class="blue-normal" style="width:95%;" value="<%=intNightRate%>"/></td>
                                    </tr>
                                </table>
                              </td>
                                <td valign="top" class="blue-normal" align="center">  &nbsp;   </td>
                              </tr>	      
<%          end select %>                                                                       
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note</td>
                              <td valign="middle" class="blue">
                              
                              <TEXTAREA NAME="txtNote" ROWS=2 style="width:95%" class="blue-normal"><%=strNote%></TEXTAREA>

							</td>
                             <td valign="top" class="blue-normal" align="center">&nbsp;</td>                                
                            </tr>  
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal"></td>
                              <td valign="middle" class="blue">
								<input type="checkbox" name="optActivate" value="1" <%if fgActivate then%>checked<%end if%>>Activate</td>
                             <td valign="top" class="blue-normal" align="center">&nbsp;</td>                                
                            </tr> 

                          </table>
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
							<tr> 
							  <td class="blue" height="20" align="left">&nbsp;&nbsp;
								<a href="javascript:Add();" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add New</a></td>
							  </tr>
                          </table>

<%if strLast<>"" then %>
						  <table width="100%" border="0" cellspacing="1" cellpadding="5">
						  
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="10%">No.</td>
                              <td class="blue" align="center" width="35%">Name</td>  
                              <td class="blue" align="center" width="40%">Note</td>  
                              <td class="blue" align="center" width="15%">Activate</td>  
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
<input type="hidden" name="txtCategoryType" value="<%=intCategoryType%>">
<input type="hidden" name="txtCategoryID" value="<%=intCategoryID%>">

<div id="submenu1" class="anylinkcss">
<ul>
<%for i=0 to UBound(arrCategories)%>
<li><a href="javascript:Category(<%=i%>)"><%=arrCategories(i)%></a></li>
<%next%>
</ul>
</div>

</form>

</body>
</html>
