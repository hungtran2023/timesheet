<SCRIPT language="VBScript" RUNAT="SERVER">
'****************************************
' function: showlabel_1
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function showlabel_1(ByVal text)
	if not isNull(text) then
	    replacedString = Replace(text, "<", "&lt;")
	  	showlabel_1 = replacedString
	end if
End function
'****************************************
' function: showlabel_2
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function showlabel_2(ByVal text)
	if not isNull(text) then
		replacedString = Replace(text, ">", "&gt;")
  		showlabel_2 = replacedString
  	end if
End function
'****************************************
' function: showlabel_3
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************

Function showlabel_3(ByVal text)
	if not isNull(text) then
		replacedString = Replace(text, chr(34), "&quot;")
  		showlabel_3 = replacedString
  	end if
End function
'****************************************
' function: showlabel
' Description: Show data to table
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************

Function showlabel(ByVal text)
	if not isNull(text) then
		if(trim(text)="") then
		  showlabel = ""
		else
			text1 = showlabel_1(text)
			text2 = showlabel_2(text1)
			text3 = showlabel_3(text2)
			showlabel = text3
		end if
	end if
End function
'****************************************
' function: showvalue
' Description: Show data to text box
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function showvalue(ByVal text) 
	if not isNull(text) then
		if trim(text)=""then
		    showvalue = ""
		else
			text1 = showlabel_3(text)
		  	showvalue = text1
		end if
	end if
End function
'****************************************
' function: recCount
' Description: Count number of records
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function recCount(ByRef rsSrc)
dim arrTmp, cnt
	arrTmp = rsSrc.GetRows
	cnt = UBound(arrTmp, 2) + 1
	rsSrc.MoveFirst
	recCount = cnt
end function
'****************************************
' function: pageCount
' Description: Count the number of pages (depends on number of lines on one page)
' Parameters: recordset or array
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function pageCount(ByRef rsSrc, ByVal psize)
 dim arrTmp, cnt, numpage
	if not isArray(rsSrc) then
		rsSrc.MoveFirst
		arrTmp = rsSrc.GetRows
	else
		arrTmp = rsSrc
	end if
	cnt = UBound(arrTmp, 2) + 1
	numpage = Int(cnt/psize)
	if cnt mod psize <> 0 then
		numpage = numpage + 1
	end if
	pageCount = numpage
end function
'****************************************
' Function: selectday
' Description: make a list to select a day
' Parameters: name of list
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectday(ByVal vname, ByVal vselected, ByVal fgNull)
	strlst = "<select name='" & vname & "' size='1' height='26px' width='40px' " &_
			"style='width:40px;height=24px; background-color: #ffffff; border-style:1px; border-color: #A0AEA4' " &_
			"class='blue-normal'"
	
	if fgNull<>-1 then strlst=strlst & " onClick='CheckMode(this)'"
	strlst=strlst & ">"
	strTmp = ""
	For i = 1 to 31
		strTmp1 = CStr(i)
		if len(strTmp1) = 1 then strTmp1 = "0" & strTmp1
		if i = vselected then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" & strTmp1 & "' " & strSel & ">" & strTmp1 & "</option>"
	Next
	if fgNull=1 then 
		if vselected = 0 then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value=''" & strSel & ">--</option>"
	end if	
	strlst = strlst & strTmp & "</select>"
	selectday = strlst
end function
'****************************************
' Function: selectmonth
' Description: make a list to select a month
' Parameters: name of listbox, item selected, have null value or not
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectmonth(ByVal vname, Byval vselected, ByVal fgNull)
	arrlongmon  = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	strlst = "<select name='" & vname & "' size='1' height='26px' width='50px' " &_
			"style='width:50px;height=24px; background-color: #ffffff; border-style:1px; border-color: #A0AEA4' " &_
			"class='blue-normal'"
	if fgNull<>-1 then strlst=strlst & " onClick='CheckMode(this)'"
	strlst=strlst & ">"
	strTmp = ""
	For i = 1 to 12
		strTmp1 = CStr(i)
		if len(strTmp1) = 1 then strTmp1 = "0" & strTmp1 end if
		if i = vselected then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" & strTmp1 & "' " & strSel & ">" & arrlongmon(i-1) & "</option>"
	Next
	if fgNull=1 then 
		if vselected = 0 then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value=''" & strSel & ">--</option>"
	end if
	strlst = strlst & strTmp & "</select>"
	set arrlongmon = nothing
	selectmonth = strlst
end function
'****************************************
' Function: selectyear
' Description: make a list to select a year
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectyear(ByVal vname, ByVal vselected, Byval startyear, Byval endyear, ByVal fgNull)
	strlst = "<select name='" & vname & "' size='1' height='26px' width='50px' " &_
			"style='width:50px;height=24px; background-color: #ffffff; border-style:1px; border-color: #A0AEA4' " &_
			"class='blue-normal'"
	if fgNull<>-1 then strlst=strlst & " onClick='CheckMode(this)'"
	strlst=strlst & ">"
	strTmp = ""
	For i = startyear to endyear
		if i = vselected then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" & CStr(i) & "' " & strSel & ">" & CStr(i) & "</option>"
	Next

	if fgNull=1 then 
		if vselected = 0 then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value=''" & strSel & ">--</option>"
	end if
	strlst = strlst & strTmp & "</select>"
	selectyear = strlst
end function
'****************************************
' function: SetAttRs
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub SetAttRs(ByRef rsSrc)
	rsSrc.CursorLocation = adUseClient     ' Set the Cursor Location to Client

  ' Append some Fields to the Fields Collection
	rsSrc.Fields.Append "sID", adInteger
	rsSrc.Fields.Append "sName", adVarChar, 100
	rsSrc.Fields.Append "sParentID", adInteger,,adFldIsNullable
	rsSrc.Fields.Append "ChainID", adVarChar, 100,adFldIsNullable
	rsSrc.Fields.Append "Owner", adInteger
	rsSrc.Fields.Append "RightOn", adInteger
	rsSrc.Fields.Append "Leaf", adInteger
	rsSrc.CursorType = adOpenStatic
	rsSrc.Open
End Sub
'****************************************
' function: CopyData
' Description:
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub CopyData(ByRef rsSrc, ByRef rsDes)
	if rsDes.RecordCount>0 then
	  set rsDes = Server.CreateObject("ADODB.Recordset")
	  Call SetAttRs(rsDes)
	end if
	rsSrc.MoveFirst
	Do While not rsSrc.EOF
	    rsDes.AddNew Array("sID", "sName", "sParentID", "chainID", "Owner", "RightOn", "Leaf"), _
					Array(rsSrc(0), rsSrc(1), rsSrc(2), rsSrc(3), rsSrc(4), rsSrc(5), rsSrc(6))
		rsSrc.MoveNext
	Loop
End Sub
'****************************************
' function: AppendList
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function AppendList (ByVal strsName, ByVal intLevel, ByVal blnShow, ByVal intValue, ByVal strChain)
Dim strTmp, i, strColor
	strTmp = ""
	If intLevel > 0 Then		
		For i = 1 to intLevel
			strTmp = strTmp & "&nbsp;&nbsp;"
		Next
	End If
	if blnShow = true then
		AppendList = "<option value='" & intValue & "@" & strChain & "'>" & strTmp & "* " & showlabel(strsName) & "</option>"
	else
		AppendList = "<option value=''>" & strTmp & showlabel(strsName) & "</option>"
	end if
End Function
'****************************************
' function: FetchList
' Description: recursive function
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub FetchList(ByRef rsGet, ByRef strTree, ByVal intLevel, ByRef rsAll,  ByVal fgKind, ByVal fgRighonfParent)
Dim strName, intContinue
	Do Until rsGet.EOF	
		blnOwner = false
		select case fgKind
		case 1 'add subtask, if having right on parent (or own parent) then having all right on their child
			If rsGet("Owner")<>0 or rsGet("RightOn")<>0 or fgRighonfParent then 
				blnOwner = true
			end if
			blnAncestor = (rsGet("Owner")<>0) or fgRighonfParent
		case 2 'assignment
			If rsGet("Leaf") = 0 and (rsGet("RightOn")<>0 or rsGet("Owner")<>0 or fgRighonfParent) then 
				blnOwner = true
			end if
			blnAncestor = blnOwner or (rsGet("Owner")<>0) or fgRighonfParent
		case 3' assign right
			If rsGet("Owner")<>0 or fgRighonfParent then 
				blnOwner = true
			end if
			blnAncestor = blnOwner or (rsGet("Owner")<>0) or (rsGet("RightOn")<>0) or fgRighonfParent
		End select		
		
	    strTree = strTree & AppendList(rsGet("sName"), intLevel, blnOwner, rsGet("sID"), rsGet("ChainID"))
		rsAll.Filter = "sParentID = " & rsGet("sID")
		intContinue = 0
		If rsAll.RecordCount > 0 then
		  intContinue = 1
		  Call CopyData(rsAll, arrRs(intLevel + 1))
		  arrRs(intLevel + 1).MoveFirst
		End If
		rsAll.Filter = ""
		If intContinue = 1 and (((intLevel + 1) < 4 and fgKind = 1) or (fgKind <> 1)) Then 'for add subtask, only show 4 levels
		  FetchList arrRs(intLevel + 1), strTree, intLevel + 1, rsAll, fgKind, blnAncestor
		End If
		rsGet.MoveNext
	Loop
End Sub
'****************************************
' function: makeList
' Description: making list of tasks
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function makeList(ByRef rsSrc, ByVal fgKind)
Dim strRes
		'-- Create the ADO Objects
		For i = 0 to 4
		  set arrRs(i) = Server.CreateObject("ADODB.Recordset")
		  Call SetAttRs(arrRs(i))
		Next
		rsSrc.Filter = "sParentID = 0"	
		arrRs(0).AddNew Array("sID", "sName", "sParentID", "ChainID", "Owner", "RightOn", "Leaf"),_
		  			 Array(rsSrc(0), rsSrc(1), rsSrc(2), rsSrc(3), rsSrc(4), rsSrc(5), rsSrc(6))
		'make string list of task
		rsSrc.Filter = ""
		arrRs(0).MoveFirst
		rsSrc.MoveFirst
		strRes=""
		
		blnRightOn = false
		if fgKind <> 3 then
			If arrRs(0)("RightOn")<>0 or arrRs(0)("Owner")<>0 then 'add subtask or assignment, 
			'if having right on parent(or own parent) then having all of rights on their child
				blnRightOn = true
			End if
		else
			If arrRs(0)("Owner")<>0 then 'right
				blnRightOn = true
			End if
		end if
		
		FetchList arrRs(0), strRes, 0, rsSrc, fgKind, blnRightOn
		For i = 0 to 4
		  arrRs(i).Close
		  Set arrRs(i) = Nothing
		Next
		makeList = "<select name='lsttask' class='blue-normal' style='HEIGHT: 22px; WIDTH: 180px'>" & strRes & "</select>"
End Function
'**************************************************
' Function: Linecount
' Description: count the number of lines in strSrc
' Parameters: 
' Return value: integer
' Author: 
' Date: 29/08/2001
' Note:
'**************************************************
function Linecount(ByVal strSrc)
	if trim(strSrc) = "" then
		Linecount = 0
	else
		intTmp = 1
		intNum = 0		
		While intTmp<>0
			intTmp = Instr(intTmp+4, strSrc, "<BR>")
			if intTmp<>0 then intNum = intNum + 1
		Wend
		Linecount = intNum
	end if
End function
'**************************************************
' Sub: MakeHTTP
' Description: http://newtimesheet/
' Parameters: userID of login user
' Return value: array
' Author: 
' Date: 28/06/2001
' Note:
'**************************************************
Sub MakeHTTP
	'create http path				
	Dim strTmp, strHTTP
	strTmp = Request.ServerVariables("URL")
	strTmp = Mid(strTmp , 1, Instr(2, strTmp, "/")-1)
	strHTTP = "http://" & Request.ServerVariables("SERVER_NAME") & strTmp & "/"
	Session("strHTTP") = strHTTP
End Sub


'**************************************************
' Sub: getarrMenu
' Description: make a array that contains the menus for login user
' Parameters: userID of login user
' Return value: array
' Author: 
' Date: 28/06/2001
' Note:
'**************************************************
Function getarrMenu(Byval userID)

  '-----------------------------------------
  '-- Declare all variables used
  '-----------------------------------------

  Dim objConn                 '-- The ADO Connection to the Database
  Dim objRs1             '-- The ADO Parent Recordset (Publishers)
  Dim objRs2         '-- The ADO child Recordset (Titles)
 
  Dim strShape                '-- The SHAPE Syntax
  Dim strConn                 '-- Connection String to the Database
  Dim intCnt
  
  '-- Create the ADO Objects
  set objConn = Server.CreateObject("ADODB.Connection")
  set objRs1 = Server.CreateObject("ADODB.Recordset")
  set objRs2 = Server.CreateObject("ADODB.Recordset")
  
  strConn = Application("g_strConnectShape")
  
  '-- Define the Shape Provider
  objConn.Provider = "MSDataShape"

  '-- Open the Connection
  objConn.Open strConn
  
  strShape = ""  
  
  strShape = strShape & "SHAPE {SELECT DISTINCT d.Description, d.Form, isnull(e.Description, d.Description) GroupName, d.GroupID, ISNULL(e.GroupID, 0) varLevel, " &_
					"e.Form query, d.LevelOrder, e.LevelOrder ParentOrder FROM ((((select * from ATC_UserGroup where UserID=" & userID & ") a " &_
					"INNER JOIN ATC_Group b ON a.GroupID = b.GroupID) " &_
					"INNER JOIN ATC_Permissions c ON b.GroupID = c.GroupID) " &_
					"LEFT JOIN ATC_Functions d ON c.FunctionID = d.FunctionID) " &_
					"LEFT JOIN ATC_Functions e ON d.GroupID = e.FunctionID " &_
					"WHERE d.Form like '%/%' or d.Form like '%asp' order by d.GroupID, d.LevelOrder ASC} AS rs " &_
					"COMPUTE rs, ANY(rs.GroupName) Group By GroupID"

  '-- Set the Parent Recordset Connection to the Active Connection
  objRs1.ActiveConnection = objConn

  '-- Open the Data Shape
  objRs1.Open strShape

  Dim arrMenu()
  intCnt = -1  

  '-- Begin with the parent
  Do Until objRs1.EOF
    intCnt = intCnt + 1

	ReDim Preserve arrMenu(4, intCnt)
  '-- Get out the Group
    arrMenu(0, intCnt) = 1 ' flag
    arrMenu(1, intCnt) = showlabel(objRs1("Group"))
        '-- Go to the child
    Set objRs2 = objRs1("rs").Value
    arrMenu(2, intCnt) = objRs2("query")
    arrMenu(3, intCnt) = objRs2("varLevel") '0: level 1, <>0: level 2
    arrMenu(4, intCnt) = objRs2("ParentOrder")
	

      '-- Loop through the Titles    
    Do Until objRs2.EOF    
	  intCnt = intCnt + 1

	  ReDim Preserve arrMenu(4, intCnt)
  '-- Get out the Child
      arrMenu(0, intCnt) = 0 ' flag
      arrMenu(1, intCnt) = showlabel(objRs2("Description"))
      arrMenu(2, intCnt) = objRs2("Form")
      arrMenu(3, intCnt) = objRs2("varLevel")
      arrMenu(4, intCnt) = objRs2("LevelOrder")
      
      '-- Move to the next Title
      objRs2.MoveNext
    Loop   
            
    '-- Move to the next Publisher
    objRs1.MoveNext
  Loop

'-- Clean up and Destory used objects

  on error resume next
  objRs1.Close
  objRs2.Close

  objConn.Close

  set objRs1 = nothing
  set objRs2 = nothing
  set objConn = nothing

'------append default menu (tools)

  ReDim Preserve arrMenu(4, intCnt + 3)
  intCnt = intCnt + 1
  arrMenu(0, intCnt) = 1 ' flag
  arrMenu(1, intCnt) = "Tools" 'objRs2("Description")
  arrMenu(2, intCnt) = "" 'objRs2("Form")
  arrMenu(3, intCnt) = 0 'objRs2("varLevel")
  arrMenu(4, intCnt) = 100
  
  intCnt = intCnt + 1
  arrMenu(0, intCnt) = 0 ' flag
  arrMenu(1, intCnt) = "Preferences" 'objRs2("Description")
  arrMenu(2, intCnt) = "tools/preferences.asp" 'objRs2("Form")
  arrMenu(3, intCnt) = 0 'objRs2("varLevel")
  arrMenu(4, intCnt) = 1
  
  intCnt = intCnt + 1
  arrMenu(0, intCnt) = 0 ' flag
  arrMenu(1, intCnt) = "Change password" 'objRs2("Description")
  arrMenu(2, intCnt) = "tools/changepassword.asp" 'objRs2("Form")
  arrMenu(3, intCnt) = 0 'objRs2("varLevel")
  arrMenu(4, intCnt) = 2
  
'----------
  if Err.number>0 then
	'Response.Write "ERROR"
	Err.Clear
  end if

  If IsEmpty(arrMenu) then
	getarrMenu = ""
  else
  	getarrMenu = arrMenu
  end if
End function
'**************************************************
' Sub: MakepartLevel1
' Description: make html string fora group, is called by getMenuTMS
' Parameters: arrMenu: contain all of menu for login user
'		strGroupName : name of group need to make
'		strValue : URL of current page (maybe have query string)
'		strChose : real URL of current page (in database, without query string)
'       strSelectedID : ID of just selected menu (who has children)
'		strGroupID : id of group
' Return value: string (html string)
' Author: 
' Date: 28/06/2001
' Note:
'**************************************************
Function MakepartLevel1(Byref arrMenu, Byval strValue, Byval strChose, byval strGroupName, Byval strSelectedID, Byval strGroupID)
	
	intManage = -1
	strRes = ""
	intLen = Ubound(arrMenu, 2)
'Seek to record
	For i = 0 to intLen
	'Response.Write strValue
		If arrMenu(0, i) = 1 and arrMenu(1, i) = strGroupName then 'is a group name
			intManage = i + 1
			if Instr(strValue, "?")>0 then
				'testing for choose_menu
				if Instr(strValue, "choose_menu=")>0 then
					idx1 = Instr(strValue, "choose_menu=") + len("choose_menu=")
					if Instr(idx1, strValue, "&") > 0 then 
						idx2 = Instr(idx1, strValue, "&")
					else
						idx2 = len(strValue) + 1
					end if
					strHref= mid(strValue, 1, idx1-1) & strGroupID & mid(strValue, idx2, len(strValue))
				else
					strHref = strValue & "&choose_menu=" & strGroupID
				end if
				'testing for fgMenu
				if Instr(strValue, "fgMenu=")=0 then
					strHref = strHref & "&fgMenu=1"
				end if
			else
				strHref = strValue & "?choose_menu=" & strGroupID & "&fgMenu=1"
			end if
			Exit for
		end if
	next
'Response.Write strGroupName & "--" & intManage & "<br>"		
	if intManage > -1 then ' if exist
		if strSelectedID = strGroupID or Mid(strSelectedID,1,1) = strGroupID then
		  'append the items belongs to strGroupName
		  strRes = strRes & "<tr><td class='blue' bgcolor='#FFFFFF' colspan='3'>" & arrMenu(1, intManage - 1) & "</td></tr>" & CHR(13)
		  Do Until arrMenu(0, intManage) <> 0
			if strChose = arrMenu(2, intManage) then
				strLink = "<font color='#CA0000'>" & arrMenu(1, intManage) & "</font>"
			else
				strLink = arrMenu(1, intManage)
			end if
'Response.Write strGroupName & "--" & strLink & "-->" & (strLink<>"empty") & "<br>"			
			'append
			if LCase(strLink)<>"empty" then
				strRes = strRes & "<tr><td class='blue' bgcolor='#FFFFFF' width='6%'>&nbsp;</td>" &_
						"<td class='blue-normal' bgcolor='#FFFFFF' width='94%' colspan='2'><img src='images/dot.gif' width='5' height='5'>&nbsp;" &_
						"<a href='" & session("strHTTP") & arrMenu(2, intManage) & "' class='c' OnMouseOver = "&chr(34)&_
						"self.status=&quot;" & arrMenu(1, intManage) & "&quot; ;" &_
						" return true;"&chr(34)&" OnMouseOut = 'self.status = &quot;&quot;'>" &_
						strLink & "</a></td></tr>" & CHR(13)
			end if
			intManage = intManage + 1
			if intManage > intLen then
				Exit Do
			end if			
		  Loop
		else
		  'only append group name
		  strRes = strRes & "<tr><td class='blue' bgcolor='#FFFFFF' colspan='3'>" &_
					"<a href='javascript:selfsubmit("&chr(34)& strHref & chr(34)& ");' OnMouseOver = "&chr(34)&"self.status=&quot;" &_
					 arrMenu(1, intManage - 1) & "&quot; ; return true;"&chr(34)&" OnMouseOut = 'self.status = &quot;&quot;'>" &_
					 arrMenu(1, intManage - 1) & "</a></td></tr>" & CHR(13)
		end if
	end if
	MakepartLevel1 = strRes
End Function

'**************************************************
' Sub: MakepartLevel2
' Description: similar to MakepartLevel1, make html string for a group, is called by getMenuTMS
' Parameters: arrMenu: contain all of menu for login user
'		strGroupName : name of group need to make
'		strValue : URL of current page (maybe have query string)
'		strChose : real URL of current page (in database, without query string)
'       strSelectedID : ID of just selected menu (who has children)
'		strGroupID : id of group
' Return value: string (html string)
' Author: 
' Date: 28/06/2001
' Note:
'**************************************************
Function MakepartLevel2(Byref arrMenu, Byval strValue, Byval strChose, byval strGroupName, Byval strSelectedID, Byval strGroupID)
	
	intManage = -1
	strRes = ""
	intLen = Ubound(arrMenu, 2)
	For i = 0 to intLen
		If arrMenu(0, i) = 1 and arrMenu(1, i) = strGroupName then 'is a group name
			intManage = i + 1
			if Instr(strValue, "?")>0 then
				'testing for choose_menu
				if Instr(strValue, "choose_menu=")>0 then
					idx1 = Instr(strValue, "choose_menu=") + len("choose_menu=")
					if Instr(idx1, strValue, "&") > 0 then 
						idx2 = Instr(idx1, strValue, "&")
					else
						idx2 = len(strValue) + 1
					end if
					strHref= mid(strValue, 1, idx1-1) & strGroupID & mid(strValue, idx2, len(strValue))
				else
					strHref = strValue & "&choose_menu=" & strGroupID
				end if
				'testing for fgMenu
				if Instr(strValue, "fgMenu=")=0 then
					strHref = strHref & "&fgMenu=1"
				end if 
			else
				strHref = strValue & "?choose_menu=" & strGroupID & "&fgMenu=1"
			end if
			Exit for
		end if
	next
	
	if intManage > -1 then ' if exist
		if strSelectedID = strGroupID or Mid(strSelectedID,1,1) = strGroupID then
		  'append
		  strRes = strRes & "<tr><td class='blue' bgcolor='#FFFFFF' width='6%'>&nbsp;</td>" &_
						"<td class='blue' bgcolor='#FFFFFF' width='94%' colspan='2'>" & arrMenu(1, intManage - 1) & "</td></tr>" & CHR(13)
						
		  do Until arrMenu(0, intManage) <> 0
			if strChose = arrMenu(2, intManage) then
				strLink = "<font color='#CA0000'>" & arrMenu(1, intManage) & "</font>"
			else
				strLink = arrMenu(1, intManage)
			end if
			'append
			strRes = strRes & "<tr><td class='blue' bgcolor='#FFFFFF' width='6%'>&nbsp;</td><td class='blue' bgcolor='#FFFFFF' width='6%'>&nbsp;</td>" &_
					"<td class='blue-normal' bgcolor='#FFFFFF' width='88%'><img src='images/dot.gif' width='5' height='5'>&nbsp;" &_
					"<a href='" & session("strHTTP") & arrMenu(2, intManage) & "' class='c' OnMouseOver = "&chr(34)&"self.status=&quot;" &_
					arrMenu(1, intManage) & "&quot; ; return true;"&chr(34)&" OnMouseOut = 'self.status = &quot;&quot;'>" &_
					strLink & "</a></td></tr>" & CHR(13)
			intManage = intManage + 1
			if intManage > intLen then
				Exit Do
			end if	
		  Loop
		else
		  'append
		  strRes = strRes & "<tr><td class='blue' bgcolor='#FFFFFF' width='6%'>&nbsp;</td><td class='blue' bgcolor='#FFFFFF' width='94%' colspan='2'>" &_
					"<a href='javascript:selfsubmit("& chr(34) & strHref & chr(34) & ");' OnMouseOver = "&chr(34)&"self.status=&quot;" & arrMenu(1, intManage - 1) &_
					"&quot; ; return true;"&chr(34)&" OnMouseOut = 'self.status = &quot;&quot;'>" &_
					arrMenu(1, intManage - 1) & "</a></td></tr>" & CHR(13)
		end if
	end if
	MakepartLevel2 = strRes
End Function


'**************************************************
' Sub: getMenuTMS
' Description: make html menu string for login user
' Parameters: arrMenu: contain all of menu for login user
'		strValue : URL of current page including query string if exists
'       intChoseMenu : group is just selected (A: management console, B: report, AA: financial, "": no selected
'       strChose : bare URL of current page without name of site and query string
' Return value: string (html string)
' Author: 
' Date: 27/06/2001
' Note:
'**************************************************

Function getMenuTMS(arrMenu, ByVal strValue, ByVal strChoseMenu, ByVal strChose, ByVal strFullName, ByVal strImage)

'Response.Write InStr(1,strFullName,"Managers Group") & "thao"
'timesheet

Dim intDailyTimesheet

intDailyTimesheet=-1

intDailyTimesheet=session("GroupManager")

	if strChose = "tms/timesheet.asp" then
		strLink = "<font color='#CA0000'>" & "Complete Timesheet" & "</font>"
	else
		strLink = "Complete Timesheet"
	end if

strResult = "<table width='100%' border='0' cellspacing='0' cellpadding='4'>"
strResult = strResult & "<tr><td colspan='3'>&nbsp;</td></tr>"

'Dim txt,pos
'txt="This is a beautiful day!"
'pos=InStr(txt,"his")
'document.write(pos)

if intDailyTimesheet=-1 then

strResult = strResult & "<tr><td class='blue' bgcolor='#FFFFFF' colspan='3'>"
strResult = strResult & "<a href='" & session("strHTTP") & "tms/timesheet.asp' OnMouseOver = 'self.status=&quot;Fill In Timesheet&quot; ;"
strResult = strResult & " return true' OnMouseOut = 'self.status = &quot;&quot;'>" & strLink & "</a></td></tr>"

' Assigned Project	
	if strChose = "assignedproject.asp" then
		strLink = "<font color='#CA0000'>" & "Assigned Projects" & "</font>"
	else
		strLink = "Assigned Projects"
	end if
' append
	strResult = strResult & "<tr><td class='blue' bgcolor='#FFFFFF' colspan='3'>" &_
			"<a href='" & session("strHTTP") & "assignedproject.asp' OnMouseOver = 'self.status=&quot;" &_
			"Assigned Projects&quot; ; return true' OnMouseOut = 'self.status = &quot;&quot;'>" & strLink & "</a></td></tr>" & CHR(13)
End if
	if isArray(arrMenu) then  ' Special rights
		'Show the items belongs to management console but not is a group name
		strRet = MakepartLevel1(arrMenu, strValue, strChose, "Management Console", strChoseMenu, "A")
		if strRet <> "" then
			strResult = strResult & strRet
		end if

		if Mid(strChoseMenu, 1, 1) = "A" then 'because Financial belong to Management Console ; strChoseMenu = "AA" 
			'Show the items belongs to Timesheet
			strRet = MakepartLevel2(arrMenu, strValue, strChose, "Timesheets", strChoseMenu, "AD")
			if strRet <> "" then
				strResult = strResult & strRet
			end if
			'Show the items belongs to Employee
			strRet = MakepartLevel2(arrMenu, strValue, strChose, "Employees", strChoseMenu, "AC")
			if strRet <> "" then
				strResult = strResult & strRet
			end if
			'Show the items belongs to Project
			strRet = MakepartLevel2(arrMenu, strValue, strChose, "Projects", strChoseMenu, "AB")
			if strRet <> "" then
				strResult = strResult & strRet
			end if			
			
			'Show the items belongs to Financial
			strRet = MakepartLevel2(arrMenu, strValue, strChose, "Financial", strChoseMenu, "AA")
			if strRet <> "" then
				strResult = strResult & strRet
			end if
			
			'Show the items belongs to Annual Leave
			strRet = MakepartLevel2(arrMenu, strValue, strChose, "Annual Leave", strChoseMenu, "AE")
			if strRet <> "" then
				strResult = strResult & strRet
			end if
			
			'Show the items belongs to ITManagement
			strRet = MakepartLevel2(arrMenu, strValue, strChose, "IT Asset Management", strChoseMenu, "AF")
			if strRet <> "" then
				strResult = strResult & strRet
			end if
		end if
		
		strRet = MakepartLevel1(arrMenu, strValue, strChose, "Reporting", strChoseMenu, "B")
		if strRet <> "" then
			strResult = strResult & strRet
		end if
		strRet = MakepartLevel1(arrMenu, strValue, strChose, "Tools", strChoseMenu, "C")
		if strRet <> "" then
			strResult = strResult & strRet
		end if		
	end if
	
'log off

	if strChose = "logoff.asp" then
		strLink = "<font color='#CA0000'>Log out</font>"
	else
		strLink = "Log out"
	end if
' append
	strResult = strResult & "<tr><td class='blue' bgcolor='#FFFFFF' colspan='3'>" &_
			"<a href='" & session("strHTTP") & "logout.asp' OnMouseOver = 'self.status=&quot;" &_
			"Log Out&quot; ; return true' OnMouseOut = 'self.status = &quot;&quot;'>" & strLink &_
			"</a><SPAN class='red'>&nbsp;" & session("USERNAME") & "</span></td></tr></table>" & CHR(13)

'--------------------------------------------------
' Fix path to images
'--------------------------------------------------

	strResult = Replace(strResult, "images/", strImage & "images/")
	getMenuTMS = strResult
End function
'**************************************************
' Sub: getarrRight
' Description: make a array containing the pages that current user have right on.
' Parameters: userID of login user
' Return value: array(2-dimension: 1: name, 2:updateable)
' Author: 
' Date: 28/06/2001
' Note:
'**************************************************
Function getarrRight(Byval userID)
Dim objDb, strConnect
dim strTemp

strConnect = Application("g_strConnect")
Set objDb = New clsDatabase
If objDb.dbConnect(strConnect) then

	strQuery = "SELECT d.Description, d.Form, c.updateable " &_
				"FROM ((((select * from ATC_UserGroup where UserID=" & userID & ") a " &_
				"INNER JOIN ATC_Group b ON a.GroupID = b.GroupID) " &_
				"INNER JOIN ATC_Permissions c ON b.GroupID = c.GroupID) " &_
				"LEFT JOIN ATC_Functions d ON c.FunctionID = d.FunctionID) " &_
				"WHERE d.Form like '%/%' or d.Form like '%asp' or d.Form like '%?' ORDER BY d.Description"

	strTemp=""
	
	ret = objDb.runQuery(strQuery)
	if ret then
		if not objDb.noRecord then
			Dim arrRight()
			intCnt = -1
			Do Until objDb.rsElement.EOF
				
				'If we have more than 1 function for this user
				'then the updatable is high priority
				if strTemp<>objDb.rsElement("Description") then
					intCnt = intCnt + 1
					ReDim Preserve arrRight(1, intCnt)
					tmp = objDb.rsElement("Form")
						
					if Instr(tmp, ".asp")=0 and Instr(tmp, ".htm")=0 then 'special
						arrRight(0, intCnt) = objDb.rsElement("Description")
					else
						while Instr(tmp, "/")<>0
							tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
						Wend
						arrRight(0, intCnt) = tmp
					end if
					if objDb.rsElement("updateable") = true then 
						arrRight(1, intCnt) = 1
					else
						arrRight(1, intCnt) = 0
					end if				
				else
					if objDb.rsElement("updateable") = true then arrRight(1, intCnt) = 1
				end if
				
				strTemp = objDb.rsElement("Description")
				
				objDb.rsElement.MoveNext
			Loop
			objDb.CloseRec
		end if
	else
		strError = objDb.strMessage
	end if
	objDb.dbdisConnect
else
	strError = objDb.strMessage
end if
set objDb = nothing

If IsEmpty(arrRight) then
	getarrRight = ""
else
  	getarrRight = arrRight
end if
End function

'**************************************************
' Sub: getarrPreference
' Description: make a array containing all of preferences that current user haved.
' Parameters: userID of login user
' Return value: array(2-dimension, 1: field, 2: record)
' Author: 
' Date: 6/08/2001
' Note:
'**************************************************
Function getarrPreference(Byval userID)
Dim objDb, strConnect
 
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
		strQuery = "SELECT isnull(FavoriteURL,'') FavoriteURL, isnull(NumofRows, 0 ) NumofRows, isnull(ProCriteria, '') ProCriteria, " &_
					"isnull(EmpCriteria, '') EmpCriteria FROM ATC_Preferences WHERE StaffID = " & userID
		ret = objDb.runQuery(strQuery)
		if ret then
			if not objDb.noRecord then
				arrPreference = objDb.rsElement.getRows
				objDb.CloseRec
			end if
		else
			strError = objDb.strMessage
		end if
		objDb.dbdisConnect
	else
		strError = objDb.strMessage
	end if
	set objDb = nothing

	If IsEmpty(arrPreference) then
		getarrPreference = ""
	else
	  	getarrPreference = arrPreference
	end if
End function
'****************************************
' function: freeListpro
' Description: free session variables of this page
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub freeListpro
	if not isEmpty(session("READY")) then session("READY") = empty
	if not isEmpty(session("rsPro")) then
		set rsPar = session("rsPro")
		session("rsPro") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("rsResult")) then
		set rsTmp = session("rsResult")
		session("rsResult") = empty
		rsTmp.Close
		set rsTmp = nothing
	end if
	if not isEmpty(session("search")) then session("search") = empty
'	if not isEmpty(session("filter")) then session("filter") = empty
'	if not isEmpty(session("CurPage")) then session("CurPage") = empty
	if not isEmpty(session("NumPage")) then session("NumPage") = empty
	if not isEmpty(session("fgShow")) then session("fgShow") = empty '0:all, 1:result returned
'	if not isEmpty(session("arrSort")) then session("arrSort") = empty
End sub
'****************************************
' function: freeAssignment
' Description: free session variables of Assignment page
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub freeAssignment
	if not isEmpty(session("READYASSIGN")) then session("READYASSIGN") = empty
	if not isEmpty(session("READYPARTICIPANT")) then session("READYPARTICIPANT") = empty
	if not isEmpty(session("rsParticipant")) then
		set rsPar = session("rsParticipant")
		session("rsParticipant") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("rsTaskCache")) then
		set rsTmp = session("rsTaskCache")
		session("rsTaskCache") = empty
		rsTmp.Close
		set rsTmp = nothing
	end if
	if not isEmpty(session("arrBookmark")) then session("arrBookmark") = empty
	if not isEmpty(session("CurPageass")) then session("CurPageass") = empty
	if not isEmpty(session("NumPageass")) then session("NumPageass") = empty
End sub
'****************************************
' function: freeAssignRight
' Description: free session variables of Assignright page
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub freeAssignRight
	if not isEmpty(session("READYRIGHT")) then session("READYRIGHT") = empty
	if not isEmpty(session("READYJUNIOR")) then session("READYJUNIOR") = empty
	if not isEmpty(session("rsJunior")) then
		set rsPar = session("rsJunior")
		session("rsJunior") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("rsTaskCache")) then
		set rsTmp = session("rsTaskCache")
		session("rsTaskCache") = empty
		rsTmp.Close
		set rsTmp = nothing
	end if
	if not isEmpty(session("arrPageright")) then session("arrPageright") = empty
	if not isEmpty(session("CurPageright")) then session("CurPageright") = empty
	if not isEmpty(session("NumPageright")) then session("NumPageright") = empty
End sub
'****************************************
' function: freeproInfo
' Description: free session variables of Assignright page
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub freeproInfo
	session("READYPRO")	= empty
	if not isEmpty(session("rsTaskCache")) then
		set rsTask = session("rsTaskCache")
		rsTask.Close
		set rsTask = nothing
		session("rsTaskCache") = empty
	end if
	if not isEmpty(session("typeofproject")) then session("typeofproject") = empty
	if not isEmpty(session("selected")) then session("selected") = empty
End Sub
'****************************************
' Function: freeListemp
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freeListemp
	if not isEmpty(session("READYPER")) then session("READYPER") = empty
	if not isEmpty(session("rsPerson")) then
		set rsPar = session("rsPerson")
		session("rsPerson") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("rsResultle")) then
		set rsTmp = session("rsResultle")
		session("rsResultle") = empty
		rsTmp.Close
		set rsTmp = nothing
	end if
	if not isEmpty(session("search")) then session("search") = empty
'	if not isEmpty(session("filteremp")) then session("filteremp") = empty
'	if not isEmpty(session("CurPagele")) then session("CurPagele") = empty
	if not isEmpty(session("NumPagele")) then session("NumPagele") = empty
	if not isEmpty(session("fgShowle")) then session("fgShowle") = empty '0: all, 1: result of search
	if not isEmpty(session("lstShortlist")) then session("lstShortlist") = empty
'	if not isEmpty(session("arrSort")) then session("arrSort") = empty
	if not isEmpty(session("arrBookmark")) then session("arrBookmark") = empty
End sub
'****************************************
' Function: freeshort
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freeShort
	if not isEmpty(session("rsShortcache")) then
		set rsPar = session("rsShortcache")
		session("rsShortcache") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("CurPageShort")) then session("CurPageShort") = empty
	if not isEmpty(session("NumPageShort")) then session("NumPageShort") = empty
End sub
'****************************************
' Function: freeSinglepro
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freeSinglepro
	if not isEmpty(session("READYSINGLEPRO")) then session("READYSINGLEPRO") = empty
	if not isEmpty(session("READYSINGLEPRO-L")) then session("READYSINGLEPRO-L") = empty
	if not isEmpty(session("arrSinglePro")) then
		arrTmp = session("arrSinglePro")
		session("arrSinglePro") = empty
		set arrTmp = nothing
	end if
	if not isEmpty(session("rsProSINGLEPRO-L")) then
		set rsPar = session("rsProSINGLEPRO-L")
		session("rsProSINGLEPRO-L") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("rsResultSINGLEPRO-L")) then
		set rsTmp = session("rsResultSINGLEPRO-L")
		session("rsResultSINGLEPRO-L") = empty
		rsTmp.Close
		set rsTmp = nothing
	end if
'	if not isEmpty(session("filter")) then session("filter") = empty
	if not isEmpty(session("fgShowSINGLEPRO-L")) then session("fgShowSINGLEPRO-L") = empty '0:all, 1:result returned
	if not isEmpty(session("CurPageSinglePro")) then session("CurPageSinglePro") = empty
	if not isEmpty(session("NumPageSinglePro")) then session("NumPageSinglePro") = empty
End sub
'****************************************
' Function: freeSumpro
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freeSumpro
	if not isEmpty(session("READYSUMPRO")) then session("READYSUMPRO") = empty
	if not isEmpty(session("arrSumPro")) then
		arrTmp = session("arrSumPro")
		session("arrSumPro") = empty
		set arrTmp = nothing
	end if
	if not isEmpty(session("arryearValid")) then session("arryearValid") = empty
	if not isEmpty(session("arrDepartment")) then session("arrDepartment") = empty
	if not isEmpty(session("CurPageSumPro")) then session("CurPageSumPro") = empty
	if not isEmpty(session("NumPageSumPro")) then session("NumPageSumPro") = empty
End sub
'****************************************
' Function: freeAdmininput
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freeAdmininput
	if not isEmpty(session("READYIN")) then session("READYIN") = empty
	if not isEmpty(session("rsItem")) then
		set rsPar = session("rsItem")
		session("rsItem") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("CurPagein")) then session("CurPagein") = empty
	if not isEmpty(session("NumPagein")) then session("NumPagein") = empty
End sub
'****************************************
' Function: freeRole
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freeRole
	'roles information
	if not isEmpty(session("arrFunccache")) then session("arrFunccache") = empty
	if not isEmpty(session("CurPageFunc")) then session("CurPageFunc") = empty
	if not isEmpty(session("NumPageFunc")) then session("NumPageFunc") = empty		
End sub
'****************************************
' Function: freelistRole
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freelistRole
	'listofrole
	if not isEmpty(session("rsRolecache")) then
		set rsPar = session("rsRolecache")
		session("rsRolecache") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("READYROLE")) then session("READYROLE") = empty
	if not isEmpty(session("CurPageRole")) then session("CurPageRole") = empty
	if not isEmpty(session("NumPageRole")) then session("NumPageRole") = empty
End sub
'****************************************
' Function: freeRoleAss
' Description: 
' Parameters: - 
'			  
' Return value: 
' Author: 
' Date: 
' Note: only use temporary
'****************************************
Sub freeRoleAss
	if not isEmpty(session("rsEmpcache")) then
		set rsPar = session("rsEmpcache")
		session("rsEmpcache") = empty
		rsPar.Close
		set rsPar = nothing
	end if
	if not isEmpty(session("READYROLEASS")) then session("READYROLEASS") = empty
	if not isEmpty(session("CurPageRoleAss")) then session("CurPageRoleAss") = empty
	if not isEmpty(session("NumPageRoleAss")) then session("NumPageRoleAss") = empty
End sub
</SCRIPT>