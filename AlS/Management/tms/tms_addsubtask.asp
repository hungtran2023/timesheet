<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

'--------------------------------------------------
' Declare variables
'--------------------------------------------------

	Dim intUserID, objEmployee, objDb, intMonth, intYear, intParentID, intAssignmentID
	Public strColor, strPID, strPName

Sub CopyDataPrivate(ByRef rsSrc, ByRef rsDes)
	
	rsSrc.MoveFirst
	
	Do While not rsSrc.EOF
	    rsDes.AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID", "fgLeaf", "AssignmentID"), _
	 					Array(rsSrc(0), rsSrc(1), rsSrc(2), rsSrc(3), rsSrc(4), rsSrc(5), rsSrc(6), rsSrc(7))
		rsSrc.MoveNext
	Loop
	
End Sub

Sub Extract(ByRef rsSrc, ByRef rsDes)
	If rsDes.RecordCount > 0 Then
	  Set rsDes = Server.CreateObject("ADODB.Recordset")
	  Call SetAtt(rsDes)
	End If
	
	rsSrc.MoveFirst
	Do While not rsSrc.EOF
	    rsDes.AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID", "fgLeaf", "AssignmentID", "Bookm"), _
	 					Array(rsSrc(0), rsSrc(1), rsSrc(2), rsSrc(3), rsSrc(4), rsSrc(5), rsSrc(6), rsSrc(7), rsSrc.Bookmark)
		rsSrc.MoveNext
	Loop
End Sub

Sub SetAttRsPrivate(ByRef rsSrc)
	rsSrc.CursorLocation = adUseClient												' Set the Cursor Location to Client
'--------------------------------------------------
' Append some Fields to the Fields Collection
'--------------------------------------------------

	rsSrc.Fields.Append "pID", advarChar, 20
	rsSrc.Fields.Append "pName", adVarChar, 150
	rsSrc.Fields.Append "sID", adInteger
	rsSrc.Fields.Append "sName", adVarChar, 150
	rsSrc.Fields.Append "sParentID", adInteger,,adFldIsNullable
	rsSrc.Fields.Append "chainID", adVarChar, 100,adFldIsNullable
	rsSrc.Fields.Append "fgLeaf", adtinyInt
	rsSrc.Fields.Append "AssignmentID", adInteger
	rsSrc.CursorType = adOpenStatic
	rsSrc.Open
End Sub

Sub SetAtt(ByRef rsSrc)
	rsSrc.CursorLocation = adUseClient												' Set the Cursor Location to Client

'--------------------------------------------------
' Append some Fields to the Fields Collection
'--------------------------------------------------

	rsSrc.Fields.Append "pID", advarChar, 20
	rsSrc.Fields.Append "pName", adVarChar, 150
	rsSrc.Fields.Append "sID", adInteger
	rsSrc.Fields.Append "sName", adVarChar, 150
	rsSrc.Fields.Append "sParentID", adInteger,,adFldIsNullable
	rsSrc.Fields.Append "chainID", adVarChar, 100,adFldIsNullable
	rsSrc.Fields.Append "fgLeaf", adtinyInt
	rsSrc.Fields.Append "AssignmentID", adInteger
	rsSrc.Fields.Append "Bookm", adInteger
	rsSrc.CursorType = adOpenStatic
	rsSrc.Open
End Sub

Function AppendTree (ByVal strpID, ByVal strpName, ByVal strsName, ByVal fgLeaf, ByVal intLevel, ByVal AssignmentID)
	Dim strTmp, i
	strTmp = ""	
	strOut = strsName
	
	If intLevel > 0 Then		
		For i = 1 to intLevel
			strTmp = strTmp & "<IMG alt='' border='0' height='18' src='../../images/t_dot.gif' width='36'>"
		Next
		strTmp = strTmp & "<IMG alt='' border='0' src='../../images/dot1.gif'>"
		strTmp = strTmp & "<IMG alt='' border='0' height='10' width='12' src='../../images/nosign.gif'>"
	Else
		strOut = strpID & " - " & strpName
	End If
	If fgLeaf = "0" Then
		If Request.QueryString("act") = "U" Then 
			strTmp = "<tr bgcolor='" & strColor &"'><td>" &  strTmp & strOut & "</td><td width='5%'><input type='checkbox' name='chkget' value='" & AssignmentID & "' onClick='CountChecked(this)'></td></tr>"
		Else
			strTmp = "<tr bgcolor='" & strColor &"'><td>" &  strTmp & strOut & "</td><td width='5%'><input type='checkbox' name='chkget' value='" & AssignmentID & "'></td></tr>"
		End If	
	Else
		strTmp = "<tr bgcolor='" & strColor &"'><td colspan='2'>" & strTmp & strOut & "</td></tr>"
	End If

	AppendTree = strTmp
	
End Function

Sub FetchChild(ByRef rsGet, ByRef strTree, ByVal intLevel)
	Dim strName, intContinue
	
	Do Until rsGet.EOF
	  strTree = strTree & AppendTree(rsGet("pID"), rsGet("pName"), rsGet("sName"), rsGet("fgLeaf"), intLevel, rsGet("Bookm"))

	  objData.Filter = "sParentID = " & rsGet("sID")

	  intContinue = 0
	  If objData.RecordCount > 0 then
		intContinue = 1
		Call Extract(objData, arrRs(intLevel+1))
		arrRs(intLevel+1).MoveFirst
	  End If
	  objData.Filter = ""
	  If intContinue = 1 Then
		FetchChild arrRs(intLevel+1), strTree, intLevel+1
	  End if
	  rsGet.MoveNext
	Loop
	
End Sub

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	
	If checkSession(session("USERID")) = False Then%>
<script LANGUAGE="javascript">
<!--
	opener.document.location = "../../message.htm";
	window.close();
	//-->
</script>
	
<%	
	End If
	intUserID = session("USERID")
	intStaffID = Request.QueryString("s")
	
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	
	intMonth	= Request.Querystring("m")
	intYear		= Request.Querystring("y")
	strType		= Request.QueryString("sel")
	strAction	= Request.QueryString("act")
	
'--------------------------------------------------
' Check subtask to write timesheet
'--------------------------------------------------

If strType = "1" Then
	Dim strError, varError, intCount

	If strAction = "U" Then
		intOldAssignmentID = Request.Cookies("assignID")

		Set objData = session("objData")

		varTmp = Request.Form("chkget")
		objData.Bookmark = int(varTmp)

		intNewAssignmentID = objData("AssignmentID")

		strError = tmsUpdateSubtask(intOldAssignmentID, intNewAssignmentID, intMonth, intYear, intStaffID)

		If strError = "" Then
%>
<script LANGUAGE="javascript">
<!--
	opener.document.frmtms.txthidden.value = "<%=intStaffID%>";
	opener.document.frmtms.action = "timesheet.asp?act=vmya"
	opener.document.frmtms.submit()
	window.close();
//-->
</script>
<%		
'		Else
'			Response.Write strError		
		End If
	Else
		
		intCount = -1
		Redim varError(intCount)
		
		Set objData = session("objData")
		For ii = 1 To Request.Form("chkget").Count
		
			varTmp = Request.Form("chkget")(ii)
			objData.Bookmark = int(varTmp)

			strPID			= Trim(objData("pId"))
			strPName		= Trim(objData("pName"))
			strSubTask		= objData("sName")
			intParentID		= objData("sParentID")
			intAssignmentID = objData("AssignmentID")
			
			Call tmsAddsubtask(strPID, strPName, intParentID, strSubTask, intAssignmentID, intMonth, intYear)
		Next
%>
<script LANGUAGE="javascript">
<!--
	opener.document.frmtms.txthidden.value = "<%=intStaffID%>";
	opener.document.frmtms.action = "timesheet.asp?act=ast"
	opener.document.frmtms.submit()
	window.close();
//-->
</script>
<%
	End If		

	objData.Close
	Set objData = Nothing
	Set session("objData") = Nothing

End If
'--------------------------------------

'--------------------------------------------------	
' Making custom record set, get all employee's Sub Task and Parent Task 
'--------------------------------------------------

Public objData
set objData = Server.CreateObject("ADODB.Recordset")								' Create the ADO Object
objData.CursorLocation = adUseClient												' Set the Cursor Location to Client

'--------------------------------------------------
' Append some Fields to the Fields Collection
'--------------------------------------------------

objData.Fields.Append "pID", advarChar, 20
objData.Fields.Append "pName", adVarChar, 120
objData.Fields.Append "sID", adInteger
objData.Fields.Append "sName", adVarChar, 100
objData.Fields.Append "sParentID", adInteger,,adFldIsNullable
objData.Fields.Append "chainID", adVarChar, 100,adFldIsNullable
objData.Fields.Append "fgLeaf", adtinyInt
objData.Fields.Append "AssignmentID", adInteger
objData.CursorType = adOpenStatic
objData.Open
	
strAncestor = ""
strConnect = Application("g_strConnect")
	  
Set objDb = New clsDatabase
blnErr = 0
If objDb.dbConnect(strConnect) then
	If Request.QueryString("search") = "" Then
	
		strQuery = "SELECT b.projectID, c.ProjectName, b.SubTaskID, b.SubTaskName, ISNULL(b.taskID, 0), ISNULL(b.ChainID, ''), 0 AS fgLeaf, a.AssignmentID FROM " &_
					"((SELECT * FROM ATC_Assignments WHERE StaffID = " & intStaffID & " AND fgDelete=0) a " &_
					"INNER JOIN ATC_Tasks b ON b.SubTaskID = a.SubTaskID) " &_
					"INNER JOIN ATC_Projects c ON c.ProjectID = b.ProjectID " &_
					"WHERE c.fgActivate = 1 AND c.fgDelete=0"

	ElseIf Request.QueryString("search") <> "" Then
	
		strSearch = replace(Request.Form("txtsearch"),"'","''")

		strQuery = "SELECT b.ProjectID, c.ProjectName, b.SubTaskID, b.SubTaskName, ISNULL(b.taskID, 0), ISNULL(b.ChainID, ''), 0 AS fgLeaf, a.AssignmentID FROM " &_
					"((SELECT * FROM ATC_Assignments WHERE StaffID = " & intStaffID & " AND fgDelete=0) a " &_
					"INNER JOIN ATC_Tasks b ON b.SubTaskID = a.SubTaskID) " &_
					"INNER JOIN ATC_Projects c ON c.ProjectID = b.ProjectID " &_
					"WHERE b.ProjectID LIKE '%" & strSearch & "%' AND c.fgActivate = 1 AND c.fgDelete=0"
	End If
			
	strError = ""
	If objDb.runQuery(strQuery) Then
		If Not objDb.noRecord Then
			
'--------------------------------------------------
' Copy data
'--------------------------------------------------
'On error resume next

			objDb.MoveFirst
			Do While not objDb.rsElement.EOF

	
				
				objData.AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID", "fgLeaf", "AssignmentID"), _
							Array(objDb.rsElement(0), objDb.rsElement(1), objDb.rsElement(2), _
							objDb.rsElement(3), objDb.rsElement(4), objDb.rsElement(5), objDb.rsElement(6), objDb.rsElement(7))
				strAncestor = strAncestor & objDb.rsElement(5)
				
			  objDb.MoveNext
			Loop
				
'--------------------------------------------------
' End of copy data
'--------------------------------------------------

			If strAncestor <> "" Then
				strAncestor = Mid(strAncestor, 1, Len(strAncestor)-1)
				blnErr = 0
				strQuery = "SELECT a.ProjectID, b.ProjectName, a.SubTaskID, a.SubTaskName, ISNULL(a.taskID, 0), ISNULL(a.ChainID, ''), 1 AS fgLeaf, 0 AS AssignmentID FROM " &_
							"ATC_Tasks a INNER JOIN ATC_Projects b ON b.ProjectID = a.ProjectID " &_
 							"WHERE a.SubTaskID IN (" & strAncestor & ") ORDER BY a.SubTaskID" 
	 	
				If objDb.runQuery(strQuery) Then
					'response.write 		strQuery 

					Call CopyDataPrivate(objDb.rsElement, objData)
				Else
					strError = objDb.strMessage
				End if
			End If
		Else
			strError = "No assigned"
		End if
	Else
		strError = objDb.strMessage
	End if
Else
	strError = objDb.strMessage
End if

If strError <> "" Then
	objDb.dbDisConnect
	Set objDb = Nothing
Else
	
'--------------------------------------------------
' Begin analyse
'--------------------------------------------------
	Dim arrRs(4)

	For i = 0 To 4
		Set arrRs(i) = Server.CreateObject("ADODB.Recordset")
		Call SetAtt(arrRs(i))
	Next

	objData.Sort = "pID"
	objData.Filter = "sParentID = 0"
	Set objRoot = Server.CreateObject("ADODB.Recordset")								' Create the ADO Object
	Call SetAtt(objRoot)
	objData.MoveFirst
	Do While not objData.EOF
	    objRoot.AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID", "fgLeaf", "AssignmentID", "Bookm"), _
	 					Array(objData(0), objData(1), objData(2), objData(3), objData(4), objData(5), objData(6), objData(7), objData.Bookmark)
		objData.MoveNext
	Loop
	
	objData.Filter = ""
	k = 0
	strTree = ""
	strLast = "<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & chr(13) & _
			  "  <tr><td bgcolor='#617DC0'>" &_
			  "<table width='100%' border='0' cellspacing='1' cellpadding='0'>" & chr(13)
	objRoot.MoveFirst
	
'--------------------------------------------------
' Loop for every project
'--------------------------------------------------

	Do Until objRoot.EOF
		k = k + 1
		arrRs(0).AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID", "fgLeaf", "AssignmentID", "Bookm"), _
						Array(objRoot(0), objRoot(1), objRoot(2), objRoot(3), objRoot(4), objRoot(5), objRoot(6), objRoot(7),  objRoot(8))
		If k Mod 2 = 1 Then
			strColor = "#E7EBF5"
		Else
			strColor = "#FFF2F2"
		End If

		FetchChild arrRs(0), strTree, 0
		
'--------------------------------------------------
' Reset all of recordset
'--------------------------------------------------

		For i = 0 To 4
		  arrRs(i).Close
		  Set arrRs(i) = Nothing
		  Set arrRs(i) = Server.CreateObject("ADODB.Recordset")
		  Call SetAtt(arrRs(i))
		Next
			
		strTmp = "<tr bgcolor='#DDDDDD'><td><table width='100%' border='0' cellspacing='1' cellpadding='0' valign='top' class='blue'>" & strTree & "</table></td></tr>"
		strLast = strLast & strTmp & chr(13)
		strTree = ""
	  	objRoot.MoveNext
	Loop

	strLast = strLast & "</table></td></tr></table>"

	'--------------------------------------------------
	' Free variable
	'--------------------------------------------------

	objRoot.Close
	For i = 0 To 4
		arrRs(i).Close
		Set arrRs(i) = Nothing
	Next
	set session("objData") = objData
	
	Set objRoot = Nothing
	
	objDb.CloseRec
	objDb.dbDisConnect
	Set objDb = Nothing    
End if

%>	

<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<title>Atlas Industries - Timesheet - Add subtask</title>

<link rel="stylesheet" href="../../timesheet.css">

<script language="javascript" src="../../library/library.js"></script>
<script LANGUAGE="JavaScript">
<!--
var url;

// Use the maxChecked variable to set the maximum number that can be checked
var maxChecked = 1
var totalChecked = 0

function p_search()
{
	if (isempty(window.document.frmtms.txtsearch.value))
	{
		alert("Enter value for searching.");
		window.document.frmtms.txtsearch.focus();
	}
	else
	{	
		url = "tms_addsubtask.asp?search=Yes&m=" + "<%=intMonth%>" + "&y=" + "<%=intYear%>" + "&act=" + "<%=strAction%>" + "&s=" + "<%=intStaffID%>" + "&sel=" + "<%=strType%>";

		window.document.frmtms.action = url;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	} 
}

function submitform()
{
	url = "tms_addsubtask.asp?sel=1&m=" + "<%=intMonth%>" + "&y=" + "<%=intYear%>" + "&act=" + "<%=strAction%>" + "&s=" + "<%=intStaffID%>";

	window.document.frmtms.action = url;
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
} 

function setchecked(val) 
{
	with (document.frmtms) 
	{
		len = elements.length;
		for(var ii=0; ii<len; ii++) 
     		if (elements[ii].name == "chkget") 
				elements[ii].checked = val
	}
}

function CountChecked(field) 
{
	if (field.checked)
	    totalChecked += 1
    else
        totalChecked -= 1

    if (totalChecked > maxChecked) 
    {
//        alert ("You can't check more than one box.")
		  field.checked = false
          totalChecked = maxChecked
    }
}

function showall()
{
	url = "tms_addsubtask.asp?m=" + "<%=intMonth%>" + "&y=" + "<%=intYear%>" + "&act=" + "<%=strAction%>" + "&s=" + "<%=intStaffID%>";

	window.document.frmtms.action = url;
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function ResetCount()
{
	totalChecked = 0
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmtms" method="post">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
<%If strError <> "" Then%>  
    <tr bgcolor="#E7EBF5">
	  <td class="red">&nbsp;<b><%=strError%></b></td>
	</tr>  
<%End If%>	
    <tr> 
      <td height="80"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%If strError = "" Then%>          
          <tr> 
            <td class="blue" width="10%" valign="middle">&nbsp; </td>
            <td class="blue-normal" align="right" width="32%" valign="middle"> 
              Search For ProjectID&nbsp; </td>
            <td class="blue" align="right" width="26%" valign="middle"> 
              <input type="text" name="txtsearch" class="blue-normal" <%If strSearch <> "" Then%> value="<%=strSearch%>" <%End If%>>
            </td>
            <td class="blue" align="right" width="32%" valign="middle"> 
              <table width="150" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
                <tr> 
                  <td width="75" bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
                    <div align="center"> 
                      <p class="blue"><a href="javascript:p_search();" class="b">Search</a> 
                    </div>
                  </td>
                  <td width="75" bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
                    <a href="javascript:showall();" class="b">Show all</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
<%End If%>          
          <tr align="center"> </tr>
          <tr> 
            <td class="title" height="50" align="center" colspan="4">Select Project</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td valign="top">
<%
'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------
	Response.Write(strLast)
%>
<%If strError = "" Then%>          
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
<%
	If Request.QueryString("act") = "U" Then
%>          
            <td class="blue-normal" align="left" height="20" width="70%">&nbsp;*Choose 
                   one checkbox, then click &quot;select&quot; to update sub-task.
            </td>
<%	Else%>
            <td class="blue-normal" align="left" height="20" width="70%">&nbsp;*Choose 
                   the checkbox, then click &quot;select&quot; to select sub-task.
            </td>
<%	End If%>            
            <td bgcolor="#FFFFFF" height="20" class="blue" align="right">
              <a href="javascript:setchecked(1);">Check All</a>&nbsp;&nbsp;&nbsp;<a href="javascript:setchecked(0);">Clear All</a>&nbsp;
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" height="20" class="blue-normal" align="center" colspan="2"> 
              <table width="120" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
                <tr> 
                  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
                    <div align="center" class="blue"><a href="javascript:submitform();" class="b">Select</a></div>
                  </td>
                  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
                    <a href="javascript:void(0);" class="b" onClick="window.close()">Close</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
<%End If%>        
      </td>
    </tr>  		
  </table>
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">

</form>
</body>
</html>