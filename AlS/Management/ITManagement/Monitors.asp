<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim intBrand, intSizeID
	Dim objEmployee, objDatabase, strError,rsData

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
'SELECT [GroupMonitorsID],[BrandID],[SizeID],[Description],[Qty] FROM [dbo].[ATC_GroupMonitors]		
			
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & i+1 & "</td>"
			strOut=strOut & "<td valign='top' class='blue'>" & _
						"<a href='javascript:CategoryDetail(" & rsSrc("GroupMonitorsID") & ");' " &_
						"class='c'>" & rsSrc("BrandName") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("SizeDetail") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("Qty") & "</td>"
			strOut=strOut & "<td valign='top' align='center' class='blue-normal'>" & rsSrc("Description") & "</td>"
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

'***************************************************************
'
'***************************************************************

function GetListBox(rsSrc,strIDField, strNameField, intIDValue)
	dim strOut
	
	strOut=""
	
	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
									
			strSelect=""
			if cint(rsSrc(strIDField)) =cint(intIDValue) then strSelect="selected"
			
			
			strOut=strOut & "<option value='" & rsSrc(strIDField) & "' " & strselect & " >" & rsSrc(strNameField)  & "</option>"
			rsSrc.MoveNext
		loop
		
	end if

	GetListBox=strOut
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

    fgDel=Request.Form("fgstatus")

	if Request.QueryString("act") = "save" then
		intID=Request.Form("txtID")
		
		if fgDel<>"D" then
		
			intBrandID=Request.Form("lbBrand")
			intSizeID=Request.Form("lbSize")
			dblQty=Request.Form("txtQty")
			strDescription=Request.Form("txtNote")
		
			if cint(intID)=-1 then
			      strSql="INSERT INTO ATC_GroupMonitors (BrandID ,SizeID ,[Description] ,[Qty]) VALUES( " &_
			                 intBrandID & "," & intSizeID & ",'" &  Replace(strDescription,"'","''") & "'," & dblQty & ")"
				strPrefix="Add new "
			else
                strSql="UPDATE [dbo].[ATC_GroupMonitors] " &_
                   "SET [BrandID] = " & intBrandID  &_
                      ",[SizeID] = " & intSizeID  &_
                      ",[Description] ='" & strDescription & "' " &_
                      ",[Qty] =  " & dblQty  &_
                 "WHERE GroupMonitorsID=" & intID
				strPrefix="Update "
			end if
		else
           strSQl="DELETE FROM [dbo].[ATC_GroupMonitors] WHERE GroupMonitorsID=" & intID
			strPrefix="Delete "
			fgDel=""
		end if
		strError=ExecuteSQL(strSql,strPrefix)		
'response.write strSql		
	else

		intID=-1
		intBrandID=0
		intSizeID=0
		strDescription=""
		dblQty=0
		
	End If
	
'--------------------------------------------------
' 
'--------------------------------------------------
	
	strSql="SELECT a.*, b.[BrandName], c.[SizeDetail] FROM [ATC_GroupMonitors] a " & _
	        "INNER JOIN [ATC_MonitorBrands] b ON a.[BrandID]=b.[BrandID] " & _
	        "INNER JOIN [dbo].[ATC_MonitorSize] c ON a.[SizeID]=c.[SizeID]"

	Call GetRecordset(strSql,rsSrc)
	
	strLast=OutBody(rsSrc)	
	
	If Request.QueryString("act") = "EDIT" Then		
		
		if rsSrc.RecordCount>0 then
		
		    intID=Request.Form("txtID")
			rsSrc.Filter = "GroupMonitorsID=" & intID
		
			intBrandID=rsSrc("BrandID")
			intSizeID=rsSrc("SizeID")
			strDescription=rsSrc("Description")
			dblQty=rsSrc("Qty")
			
		end if			
	end if


    strSql="SELECT * FROM ATC_MonitorBrands WHERE fgActivate=1"
	Call GetRecordset(strSql,rsSrc)
	
	strBrandList=GetListBox(rsSrc,"BrandID","BrandName", intBrandID)

    strSql="SELECT * FROM ATC_MonitorSize WHERE fgActivate=1"
	Call GetRecordset(strSql,rsSrc)
	
	strSizeList=GetListBox(rsSrc,"SizeID","SizeDetail", intSizeID)

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
		window.document.frmreport.action = "Monitors.asp?act=save"			
		window.document.frmreport.submit();
	}
}

function CategoryDetail(id)
{
	window.document.frmreport.txtID.value = id
	window.document.frmreport.action = "Monitors.asp?act=EDIT"			
	window.document.frmreport.submit();
}

function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "Monitors.asp?act=save"
	window.document.frmreport.submit();
}

function checkdata()
{
	if (window.document.frmreport.txtQty.value=="")
	{
		alert("Please enter a quantity for this kind of monitor.");
		document.frmreport.txtQty.focus();
		return false	
	}	
	else if (isNaN(window.document.frmreport.txtQty.value)==true) {
	    alert("Please enter a number.");
		document.frmreport.txtQty.focus();
		return false;
	}
	else if (window.document.frmreport.txtQty.value<0) {
		alert("This number must be greater than 0.");
		document.frmreport.txtQty.focus();
		return false;			
	}

	return true	
}

function Add() {
	document.frmreport.txtQty.value = "";
	document.frmreport.txtNote.value = "";
	document.frmreport.optActivate.checked=false
	document.frmreport.txtID.value=-1
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
                  <td class="title" height="50" align="center" colspan="2">Monitors</td>
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
                              <td valign="middle" class="blue-normal" width="15%">Brand *</td>
                              <td valign="middle" width="45%" class="blue">
								<select name='lbBrand' size='1' class='blue-normal' style="width:95%">
<%if strBrandList ="" then %>
                                    <option value='0'>&nbsp;</option>
<%else %>
								<%=strBrandList%>
<%End if%>								
								</select></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>       
                            
                            
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="20%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="15%">Size *</td>
                              <td valign="middle" width="45%" class="blue">
								<select name='lbSize' size='1' class='blue-normal' style="width:95%">
<%if strsizeList ="" then %>
                                    <option value='0'>&nbsp;</option>
<%else %>
								    <%=strsizeList%>
<%End if%>									
                                </select></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>       
                                              
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="20%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="15%">Quantity *</td>
                              <td valign="middle" width="45%" class="blue"><input type="text" name="txtQty" maxlength="20" class="blue-normal" style="width:95%;" value="<%=dblQty%>">
								</td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>                                                 
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Remark</td>
                              <td valign="middle" class="blue">
                              
                              <TEXTAREA NAME="txtNote" ROWS=5 style="width:95%" class="blue-normal"><%=strDescription%></TEXTAREA>

							</td>
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
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="8%">No.</td>
                              <td class="blue" align="center" width="17%">Brand</td>  
                              <td class="blue" align="center" width="10%">Size</td>
                              <td class="blue" align="center" width="10%">Qty</td> 
                              <td class="blue" align="center" width="55%">Remark</td>  
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
<input type="hidden" name="txtID" value="<%=intID%>">


</form>

</body>
</html>
