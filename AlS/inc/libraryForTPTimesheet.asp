<SCRIPT language="VBScript" RUNAT="SERVER">

'**************************************************
' Function: tmsInitialForTP
' Description: 
'			  -	Initialize array that stores timesheet data. 
'			  -	varTimesheet array stores data which has concern with projects & subtask. Its structure is:
'					* varTimesheet(0, 0, 0..n)				: Field ProjectID or SubTask Name
'					* varTimesheet(1..intDayNum, 0, 0..n)	: Office hour
'					* varTimesheet(1..intDayNum, 1, 0..n)	: Overtime hour
'					* varTimesheet(1..intDayNum, 2, 0..n)	: Overtime hour for Normal Rate
'					* varTimesheet(1..intDayNum, 3, 0..n)	: Overtime hour for Night Rate
'					* varTimesheet(1..intDayNum, 4, 0..n)	: Notes
'					* varTimesheet(intDayCol-5, 0, 0..n)	: Working-Hour's SubTotal
'					* varTimesheet(intDayCol-4, 0, 0..n)	: Field AssignmentID
'					* varTimesheet(intDayCol-3, 0, 0..n)	: Field ProjectName Or SubTaskName
'					* varTimesheet(intDayCol-2, 0, 0..n)	: "P"-Project; "S"-SubTask; "N"-None 
'					* varTimesheet(intDayCol-1, 0, 0..n)	: 0 - Not remove; 1 - Remove
'					* varTimesheet(intDayCol, 0, 0..n)		: 0 - DeActivated; 1 - Acvtiated
'
'			  - varEvent array stores data which has concern with events(training, general/admin, etc...). Its structure is:
'					* varEvent(0, 0, 0..n)					: Field EventName
'					* varEvent(1..intDayNum, 0, 0..n)		: Office hour
'					* varEvent(1..intDayNum, 1, 0..n)		: Overtime hour
'					* varEvent(1..intDayNum, 2, 0..n)		: Overtime hour for Normal Rate
'					* varEvent(1..intDayNum, 3, 0..n)		: Overtime hour for Night Rate
'					* varEvent(1..intDayNum, 4, 0..n)		: Notes
'					* varEvent(intDayNum+1, 0, 0..n)		: Working-Hour's SubTotal
'					* varEvent(intDayNum+2, 0, 0..n)		: Field EventID
'
' Parameters: - intUserID: Integer
'			  - intMonth : Integer 
'			  - intYear	 : Integer
' Return value: Error message if there are any errors
' Author: Phan Thi Hong
' Date: 02/07/2001
' Note:
'**************************************************

Function tmsInitialForTP(ByVal intUserID, ByVal intMonth, ByVal intYear)

	Dim intDayNum, intDayCol, intDayCount, intDay, ii, kk, intTMS, intTMS1, intTMS2, intNewRow, intEvent1, intEvent2, intEvent 
	Dim strTableTMS, strFirstDay, dblHours
	Dim varTMS1, varTMS2, varTimesheet, varEvent1, varEvent2, varEvent, varTimeTotal, varTimeOffTotal, varHoliday

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	intTMS		= -1
	intTMS1		= -1
	intTMS2		= -1
	intNewRow	= -1
	intEvent	= -1
	intEvent1	= -1
	intEvent2	= -1
	dblHours	= 0
	intDayNum	= GetDay(intMonth,intYear)												' Numbers of days in a month
	intDayCol	= intDayNum + 6

	strTableTMS = selectTable(intYear)													' Select table to view timesheet
	strFirstDay = FirstOfMonth(intMonth,intYear)										' Get the first day in a month				
	intDayCount	= curDayNum(intDayNum,strFirstDay)										' Numbers of days since the first day in month to now

    ReDim varTimeTotal(intDayNum, 3)													' Save Project Total
    ReDim varTimeOffTotal(intDayNum)
    
    'ReDim varTimesheet(intDayCol,2,-1)
	'ReDim varEvent(intDayNum+2,2,-1)
	
	'Add two value for OTNormal and OTNight
	ReDim varTimesheet(intDayCol,4,-1)
	ReDim varEvent(intDayNum+2,4,-1)

    For ii = 1 To intDayNum
        varTimeTotal(ii, 0) = 0
        varTimeTotal(ii, 1) = 0
        varTimeTotal(ii, 2) = 0
        varTimeTotal(ii, 3) = 0
        varTimeOffTotal(ii) = 0
    Next

'--------------------------------------------------
' End of initializing variables
'--------------------------------------------------

	strConnect = Application("g_strConnect")											' Connection string 				
	Set objDatabase = New clsDatabase 

	If objDatabase.dbConnect(strConnect) Then
'--------------------------------------------------
' Check table timesheet if it exists or not	
'--------------------------------------------------

		strSQL = "SELECT ISNULL(OBJECT_ID('" & strTableTMS & "'),0) AS TableName"

		If (objDatabase.runQuery(strSQL)) Then

			If objDatabase.getColumn_by_name("TableName") = 0 Then

				
				if (intYear<2020 or (intYear=2020 AND intMonth<11)) then
					strSQL = "SELECT EventID, EventName FROM ATC_Events WHERE (EventID = 7 OR EventID=2) ORDER BY EventID"
				else
					strSQL = "SELECT EventID, EventName FROM ATC_Events WHERE (EventID = 7) ORDER BY EventID"
				end if
									
				If (objDatabase.runQuery(strSQL)) Then
					If objDatabase.noRecord = False Then
						varEvent2 = objDatabase.rsElement.GetRows
						intEvent2 = Ubound(varEvent2,2)
				
						objDatabase.closeRec											' Close recordset				
'--------------------------------------------------
' Add event into event array
'--------------------------------------------------				
						If intEvent2 >= 0 Then
							Redim Preserve varEvent(intDayNum+2,4,intEvent2)
							For ii = 0 To intEvent2
								varEvent(0,0,ii) = trim(varEvent2(1,ii))						' Field EventName
								For kk = 1 to intDayNum									
									varEvent(kk,0,ii) = 0										' Initialize office hour
									varEvent(kk,1,ii) = 0										' Initialize overtime hour
									varEvent(kk,2,ii) = 0										' Initialize OT Normal Rate
									varEvent(kk,3,ii) = 0										' Initialize OT Night Rate
									varEvent(kk,4,ii) = ""										' Initialize note
								Next
								varEvent(intDayNum+1,0,ii) = 0									' Working-Hour's SubTotal
								varEvent(intDayNum+2,0,ii) = varEvent2(0,ii)					' Field EventID
							Next
						End If
						
						intNewRow = Ubound(varEvent,3) + 1
						Redim Preserve varEvent(intDayNum+2,4,intNewRow)
						varEvent(0,0,intNewRow) = "Total hours"									' Field EventName 
						For kk = 1 to intDayNum									
							varEvent(kk,0,ii) = 0										' Initialize office hour
							varEvent(kk,1,ii) = 0										' Initialize overtime hour
							varEvent(kk,2,ii) = 0										' Initialize OT Normal Rate
							varEvent(kk,3,ii) = 0										' Initialize OT Night Rate
							varEvent(kk,4,ii) = ""										' Initialize note
						Next
						varEvent(intDayNum+1,0,intNewRow) = 0									' Working-Hour's SubTotal
						varEvent(intDayNum+2,0,intNewRow) = -1									' Field EventID

						intNewRow = Ubound(varEvent,3) + 1
						Redim Preserve varEvent(intDayNum+2,4,intNewRow)
						varEvent(0,0,intNewRow) = "Normal hours"		 						' Field EventName 
						For kk = 1 to intDayNum								
							varEvent(kk,0,ii) = 0										' Initialize office hour
							varEvent(kk,1,ii) = 0										' Initialize overtime hour
							varEvent(kk,2,ii) = 0										' Initialize OT Normal Rate
							varEvent(kk,3,ii) = 0										' Initialize OT Night Rate
							varEvent(kk,4,ii) = ""										' Initialize note
						Next
						varEvent(intDayNum+1,0,intNewRow) = 0									' Working-Hour's SubTotal
						varEvent(intDayNum+2,0,intNewRow) = -2									' Field EventID
						
						intNewRow = Ubound(varEvent,3) + 1
						Redim Preserve varEvent(intDayNum+2,4,intNewRow)
						varEvent(0,0,intNewRow) = "Overtime hours"		 						' Field EventName 
						For kk = 1 to intDayNum									
							varEvent(kk,0,ii) = 0										' Initialize office hour
							varEvent(kk,1,ii) = 0										' Initialize overtime hour
							varEvent(kk,2,ii) = 0										' Initialize OT Normal Rate
							varEvent(kk,3,ii) = 0										' Initialize OT Night Rate
							varEvent(kk,4,ii) = ""										' Initialize note
						Next
						varEvent(intDayNum+1,0,intNewRow) = 0									' Working-Hour's SubTotal
						varEvent(intDayNum+2,0,intNewRow) = -3									' Field EventID
						
						session("varEvent") = Empty
						session("varTimesheet") = Empty
						session("varEvent") = varEvent

						Set varEvent = Nothing
						Set varEvent2 = Nothing
					End If	
				Else
					strError = objDatabase.strMessage
					tmsInitialForTP = strError
					Exit Function 	
				End If

				strError = "No data for your request."
'				strError = ""
				tmsInitialForTP = strError
				Exit Function 	
			End If	
			objDatabase.closeRec()
		Else
			strError = objDatabase.strMessage
			tmsInitialForTP = strError
			Exit Function 	
		End If

'--------------------------------------------------
' Initialize holiday array
'--------------------------------------------------
		
		If isEmpty(session("varHoliday")) = False Then
			session("varHoliday") = Empty
		End If
	
		strSQL = "exec GetListHolidays " & intMonth & ", " & intYear & ", null, null, 0"

		If (objDatabase.runQuery(strSQL)) Then
			If objDatabase.noRecord = False Then
				varHoliday = objDatabase.rsElement.GetRows
				session("varHoliday") = varHoliday
				objDatabase.closeRec
			End If
		Else
			strError = objDatabase.strMessage
			tmsInitialForTP = "1. " & strError
			Exit Function 	
		End If
	
'--------------------------------------------------
' Initialize timesheet array
' Add more two fields: OTNormal, OTNight 
'--------------------------------------------------
	   
		strSQL = "SELECT a.AssignmentID, TDate, isnull(Hours,0) AS Hours, isnull(OverTime,0) AS Overtime, isnull(OTNormal,0) AS OTNormal,isnull(OTNight,0) AS OTNight, Note" & _
					" FROM " & strTableTMS & " a INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID" & _
					" WHERE a.StaffID=" & intUserID & " AND EventID=1 AND Month(TDate)=" & intMonth & " AND Year(TDate)=" & intYear & _
					" ORDER BY TDate"

		If (objDatabase.runQuery(strSQL)) Then
			If objDatabase.noRecord = False Then
				varTMS1 = objDatabase.rsElement.GetRows
				intTMS1 = Ubound(varTMS1,2)
				objDatabase.closeRec														' Close recordset
				
'--------------------------------------------------
' Add project, subtask into timesheet array
'--------------------------------------------------				

				strSQL = "SELECT a.AssignmentID, c.ProjectID, ProjectName, isnull(c.ParentName,'None') AS ParentName, c.ChildName, (sum(isnull(Hours,0))+sum(isnull(OverTime,0))) AS SubTotal, d.fgActivate" & _
							" FROM " & strTableTMS & " a INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID" & _ 
							" INNER JOIN (SELECT ChildInfo.ParentID AS SubTaskID, ChildInfo.ProjectID AS ProjectID, ParentInfo.ParentName AS ParentName, ChildInfo.ChildName AS ChildName" & _
							" FROM (SELECT SubTaskID AS ParentID, TaskID, SubTaskName AS ChildName, ProjectID FROM ATC_Tasks) AS ChildInfo" & _ 
							" LEFT JOIN (SELECT ATC_Tasks.SubTaskID AS ParentID, ATC_Tasks1.SubTaskName AS ParentName FROM ATC_Tasks INNER JOIN ATC_Tasks ATC_Tasks1 ON ATC_Tasks.TaskID = ATC_Tasks1.SubTaskID) AS ParentInfo ON ChildInfo.ParentID=ParentInfo.ParentID) AS c ON b.SubTaskID=c.SubTaskID" & _ 
							" INNER JOIN (SELECT ProjectID, ProjectName, convert(varchar,fgActivate) AS fgActivate FROM ATC_Projects) AS d ON c.ProjectID=d.ProjectID WHERE a.StaffID=" & intUserID & " AND EventID=1 AND Month(TDate)=" & intMonth & " AND Year(TDate)=" & intYear & _ 
							" GROUP BY a.AssignmentID, c.ProjectID, ProjectName, c.ParentName, c.ChildName, d.fgActivate ORDER BY c.ProjectID"
'Response.Write strsql
'Response.End

				If (objDatabase.runQuery(strSQL)) Then
					If objDatabase.noRecord = False Then
						varTMS2 = objDatabase.rsElement.GetRows
						intTMS2 = Ubound(varTMS2,2)
						objDatabase.closeRec												' Close recordset
						
						If intTMS2 >= 0 Then
							For ii = 0 To intTMS2
								If varTMS2(3,ii) = "None" Then								' Project no subtask
								
									intNewRow = Ubound(varTimesheet,3) + 1
									
									'Redim Preserve varTimesheet(intDayCol,2,intNewRow)
									Redim Preserve varTimesheet(intDayCol,4,intNewRow)
							
									varTimesheet(0,0,intNewRow) = trim(varTMS2(1,ii))		' Field ProjectID	
									
									For kk = 1 to intDayNum									
										varTimesheet(kk,0,intNewRow) = 0					' Initialize office hour
										varTimesheet(kk,1,intNewRow) = 0					' Initialize overtime hour
										varTimesheet(kk,2,intNewRow) = 0					' Initialize OT Normal Rate
										varTimesheet(kk,3,intNewRow) = 0					' Initialize OT Night
										varTimesheet(kk,4,intNewRow) = ""					' Initialize note
									Next
									
									varTimesheet(intDayCol-5,0,intNewRow) = varTMS2(5,ii)	' Working-Hour's SubTotal
									varTimesheet(intDayCol-4,0,intNewRow) = varTMS2(0,ii)	' Field AssignmentID
									varTimesheet(intDayCol-3,0,intNewRow) = varTMS2(2,ii)	' Field ProjectName
									varTimesheet(intDayCol-2,0,intNewRow) = "P"				' "P"-Project; "S"-SubTask; "N"-None 
									varTimesheet(intDayCol-1,0,intNewRow) = 0				' 0 - Not remove; 1 - Remove
									If varTMS2(6,ii) Then
										varTimesheet(intDayCol,0,intNewRow) = 1				' 0 - DeActivated; 1 - Activated
									Else
										varTimesheet(intDayCol,0,intNewRow) = 0				' 0 - DeActivated; 1 - Activated
									End If										
								Else														' Project has subtask
									
'--------------------------------------------------
' This row stores project name and no have time
'--------------------------------------------------								
									If Trim(varTMS2(1,ii)) <> trim(strPID) Then

										intNewRow = Ubound(varTimesheet,3) + 1
										'
										'Redim Preserve varTimesheet(intDayCol,2,intNewRow)
										
										Redim Preserve varTimesheet(intDayCol,4,intNewRow)

										varTimesheet(0,0,intNewRow) = trim(varTMS2(1,ii))		' Field ProjectID	
										For kk = 1 to intDayNum									
											varTimesheet(kk,0,intNewRow) = 0					' Initialize office hour
											varTimesheet(kk,1,intNewRow) = 0					' Initialize overtime hour
											varTimesheet(kk,2,intNewRow) = 0					' Initialize OT Normal hour
											varTimesheet(kk,3,intNewRow) = 0					' Initialize OT Night hour
											varTimesheet(kk,4,intNewRow) = ""					' Initialize note
										Next
										varTimesheet(intDayCol-5,0,intNewRow) = 0				' Working-Hour's SubTotal
										varTimesheet(intDayCol-4,0,intNewRow) = ""				' Field AssignmentID
										varTimesheet(intDayCol-3,0,intNewRow) = varTMS2(2,ii)	' Field ProjectName
										varTimesheet(intDayCol-2,0,intNewRow) = "N"				' "P"-Project; "S"-SubTask; "N"-None 
										varTimesheet(intDayCol-1,0,intNewRow) = 0				' 0 - Not remove; 1 - Remove
										varTimesheet(intDayCol,0,intNewRow)	  = 0				' 0 - Deactivated; 1 - Activated
									
									End If
									
'--------------------------------------------------
' This row stores subtask name
'--------------------------------------------------								

									intNewRow = Ubound(varTimesheet,3) + 1
									
									'Redim Preserve varTimesheet(intDayCol,2,intNewRow)
									Redim Preserve varTimesheet(intDayCol,4,intNewRow)
							
									varTimesheet(0,0,intNewRow) = trim(varTMS2(1,ii))			' Field ProjectID		
									For kk = 1 to intDayNum									
										varTimesheet(kk,0,intNewRow) = 0						' Initialize office hour
										varTimesheet(kk,1,intNewRow) = 0						' Initialize overtime hour
										varTimesheet(kk,2,intNewRow) = 0						' Initialize OT Normal hour
										varTimesheet(kk,3,intNewRow) = 0						' Initialize OT Night hour
										varTimesheet(kk,4,intNewRow) = ""						' Initialize note
									Next
									varTimesheet(intDayCol-5,0,intNewRow) = varTMS2(5,ii)		' Working-Hour's SubTotal
									varTimesheet(intDayCol-4,0,intNewRow) = varTMS2(0,ii)		' Field AssignmentID
'									varTimesheet(intDayCol-3,0,intNewRow) = trim(varTMS2(3,ii)) & " _ " & trim(varTMS2(4,ii))
									varTimesheet(intDayCol-3,0,intNewRow) = "&nbsp; " & trim(varTMS2(4,ii))
									varTimesheet(intDayCol-2,0,intNewRow) = "S"					' "P"-Project; "S"-SubTask; "N"-None 
									varTimesheet(intDayCol-1,0,intNewRow) = 0					' 0 - Not remove; 1 - Remove
									If varTMS2(6,ii) Then
										varTimesheet(intDayCol,0,intNewRow) = 1					' 0 - DeActivated; 1 - Activated
									Else
										varTimesheet(intDayCol,0,intNewRow) = 0					' 0 - DeActivated; 1 - Activated
									End If										

									strPID = trim(varTMS2(1,ii))
								End If		
							Next
						End If
					End If	
					
'--------------------------------------------------
' Restore time to timesheet array					
'--------------------------------------------------
					
					intTMS = Ubound(varTimesheet,3)
					
					If intTMS1 >= 0 Then
						For ii = 0 To intTMS1
							If CDbl(varTMS1(2,ii)) > 0 Or CDbl(varTMS1(3,ii)) > 0 Then
								intDay = Day(varTMS1(1,ii))				
								If intDay <= intDayCount Then
									For kk = 0 To intTMS
										If trim(varTMS1(0,ii)) = trim(varTimesheet(intDayCol-4,0,kk)) Then			
											varTimesheet(intDay,0,kk) = CDbl(varTMS1(2,ii))									' Office hour
											varTimeTotal(intDay,0)	  = CDbl(varTimeTotal(intDay,0)) + CDbl(varTMS1(2,ii))	
											varTimesheet(intDay,1,kk) = CDbl(varTMS1(3,ii))									' Overtime hour
											varTimeTotal(intDay,1)	  = CDbl(varTimeTotal(intDay,1)) + CDbl(varTMS1(3,ii))	
											varTimesheet(intDay,2,kk) = CDbl(varTMS1(4,ii))									' OT Normal
											varTimeTotal(intDay,2)	  = CDbl(varTimeTotal(intDay,2)) + CDbl(varTMS1(4,ii))	
											varTimesheet(intDay,3,kk) = CDbl(varTMS1(5,ii))									' OT Night
											varTimeTotal(intDay,3)	  = CDbl(varTimeTotal(intDay,3)) + CDbl(varTMS1(5,ii))	
											
											varTimesheet(intDay,4,kk) = varTMS1(6,ii)										' Notes
											Exit For
										End If							
									Next							
								End If	
							End If				
						Next	
					End If
				Else
					strError = objDatabase.strMessage
					tmsInitialForTP = "2. " & strError
					Exit Function 	
				End If
			End If
		Else
			strError = objDatabase.strMessage
			tmsInitialForTP = "3. " & strError
			Exit Function 	
		End If

'--------------------------------------------------
' Initialize event array
'--------------------------------------------------

		intNewRow = 0
' AND EventID <> 2		
		if (intYear<2020 or (intYear=2020 AND intMonth<11)) then
					strSQL = "SELECT EventID, EventName FROM ATC_Events WHERE (EventID = 7 OR EventID=2) ORDER BY EventID"
				else
					strSQL = "SELECT EventID, EventName FROM ATC_Events WHERE (EventID = 7) ORDER BY EventID"
				end if
		
		If (objDatabase.runQuery(strSQL)) Then
			If objDatabase.noRecord = False Then
				varEvent2 = objDatabase.rsElement.GetRows
				intEvent2 = Ubound(varEvent2,2)
				objDatabase.closeRec													' Close recordset
				
'--------------------------------------------------
' Add event into event array
'--------------------------------------------------				
				
				If intEvent2 >= 0 Then
					Redim Preserve varEvent(intDayNum+2,4,intEvent2)
					For ii = 0 To intEvent2
						varEvent(0,0,ii) = trim(varEvent2(1,ii))						' Field EventName
						For kk = 1 to intDayNum									
							varEvent(kk,0,ii) = 0										' Initialize office hour
							varEvent(kk,1,ii) = 0										' Initialize overtime hour
							varEvent(kk,2,ii) = 0										' Initialize OT Normal Rate
							varEvent(kk,3,ii) = 0										' Initialize OT Night Rate
							varEvent(kk,4,ii) = ""										' Initialize note
						Next
						varEvent(intDayNum+1,0,ii) = 0									' Working-Hour's SubTotal
						varEvent(intDayNum+2,0,ii) = varEvent2(0,ii)					' Field EventID
					Next
				End If

				strSQL = "SELECT a.EventID, TDate, isnull(Hours,0) AS Hours, isnull(OverTime,0) AS OverTime,isnull(OTNormal,0) AS OTNormal,isnull(OTNight,0) AS OTNight, Note" & _
							" FROM " & strTableTMS & " a INNER JOIN ATC_Events b ON a.EventID=b.EventID" & _
							" WHERE a.StaffID=" & intUserID & " AND AssignmentID=1 AND Month(TDate)=" & intMonth & " AND Year(TDate)=" & intYear & _
							" ORDER BY TDate"
				If (objDatabase.runQuery(strSQL)) Then
					If objDatabase.noRecord = False Then
						varEvent1 = objDatabase.rsElement.GetRows
						intEvent1 = Ubound(varEvent1,2)
						objDatabase.closeRec											 ' Close recordset
					End If

'--------------------------------------------------
' Restore time to event array					
'--------------------------------------------------

					intEvent = Ubound(varEvent,3)

					If intEvent1 >= 0 Then
						For ii = 0 To intEvent1
							If CDbl(varEvent1(2,ii)) > 0 Or CDbl(varEvent1(3,ii)) > 0 Then
								intDay = Day(varEvent1(1,ii))				
								For kk = 0 To intEvent
									If trim(varEvent1(0,ii)) = trim(varEvent(intDayNum+2,0,kk)) Then			
										varEvent(intDay,0,kk)		= CDbl(varEvent1(2,ii))										' Office hour
										
										varEvent(intDay,1,kk)		= CDbl(varEvent1(3,ii))										' Overtime hour
										varEvent(intDay,2,kk)		= CDbl(varEvent1(4,ii))										' OT Normal Rate
										varEvent(intDay,3,kk)		= CDbl(varEvent1(5,ii))										' OT Night Rate
										varEvent(intDay,4,kk)		= varEvent1(6,ii)											' Notes
										
										varEvent(intDayNum+1,0,kk)	= CDbl(varEvent(intDayNum+1,0,kk)) + CDbl(varEvent(intDay,0,kk)) + CDbl(varEvent(intDay,1,kk))' WorkingHour's SubTotal
						
										If CInt(varEvent1(0,ii)) = 2 Or CInt(varEvent1(0,ii)) = 3 Then
											varTimeTotal(intDay,0)	= CDbl(varTimeTotal(intDay,0)) + CDbl(varEvent1(2,ii))	
											varTimeTotal(intDay,1)	= CDbl(varTimeTotal(intDay,1)) + CDbl(varEvent1(3,ii))	
											varTimeTotal(intDay,2)	= CDbl(varTimeTotal(intDay,2)) + CDbl(varEvent1(4,ii))	
											varTimeTotal(intDay,3)	= CDbl(varTimeTotal(intDay,3)) + CDbl(varEvent1(5,ii))	
										Else
											varTimeOffTotal(intDay) = CDbl(varTimeOffTotal(intDay)) + CDbl(varEvent1(2,ii))	
										End If
											
										Exit For
									End If							
								Next							
							End If				
						Next	
					End If

'--------------------------------------------------
' Add Total hours row
'--------------------------------------------------
					
					intNewRow = Ubound(varEvent,3) + 1
					Redim Preserve varEvent(intDayNum+2,4,intNewRow)
					
					varEvent(0,0,intNewRow) = "Total hours"
					
					For kk = 1 To intDayNum
						varEvent(kk,0,intNewRow) = varTimeTotal(kk,0) + varTimeTotal(kk,1) + varTimeOffTotal(kk)
						dblHours = dblHours + varTimeTotal(kk,0) + varTimeTotal(kk,1) + varTimeOffTotal(kk) 
						varEvent(kk,1,intNewRow) = 0
						varEvent(kk,2,intNewRow) = ""
					Next					
					
					varEvent(intDayNum+1,0,intNewRow) = dblHours							' Sub Total
					varEvent(intDayNum+2,0,intNewRow) = -1													

'--------------------------------------------------
' Add Normal hours row
'--------------------------------------------------
					
					dblHours = 0
					intNewRow = Ubound(varEvent,3) + 1
					Redim Preserve varEvent(intDayNum+2,4,intNewRow)
					
					varEvent(0,0,intNewRow) = "Normal hours"
					
					For kk = 1 To intDayNum
						varEvent(kk,0,intNewRow) = varTimeTotal(kk,0)
						dblHours = dblHours + varTimeTotal(kk,0)
						varEvent(kk,1,intNewRow) = 0
						varEvent(kk,2,intNewRow) = ""
					Next					
					
					varEvent(intDayNum+1,0,intNewRow) = dblHours							' Sub Total													
 					varEvent(intDayNum+2,0,intNewRow) = -2													

'--------------------------------------------------
' Add Overtime hours row
'--------------------------------------------------
					
					dblHours = 0
					intNewRow = Ubound(varEvent,3) + 1
					Redim Preserve varEvent(intDayNum+2,4,intNewRow)
					
					varEvent(0,0,intNewRow) = "Overtime hours"
					
					For kk = 1 to intDayNum
						varEvent(kk,0,intNewRow) = varTimeTotal(kk,1)
						dblHours = dblHours + varTimeTotal(kk,1)
						varEvent(kk,1,intNewRow) = 0
						varEvent(kk,2,intNewRow) = varTimeTotal(kk,2)
						varEvent(kk,3,intNewRow) = varTimeTotal(kk,3)
						varEvent(kk,4,intNewRow) = ""
					Next					
					
					varEvent(intDayNum+1,0,intNewRow) = dblHours							' Sub Total													
					varEvent(intDayNum+2,0,intNewRow) = -3													
	
					session("varTimesheet") = varTimesheet
					session("varEvent")		= varEvent
	
'--------------------------------------------------
' Test data
'
'	For ii=0 To Ubound(varTimesheet,3)
'		Response.Write "<b>" & varTimesheet(0, 0, ii) & "</b>" & "  "
'		For kk=1 to intDayNum
'			Response.Write kk & "date " & varTimesheet(kk, 0, ii) & "Off" & "  "
'			Response.Write varTimesheet(kk, 1, ii) & "Over" & "  "
'			Response.Write varTimesheet(kk, 2, ii) & "Notes" & "  "
'		Next
'		Response.Write varTimesheet(intDayCol-5, 0, ii) & "SubTotal" & "  " 	
'		Response.Write varTimesheet(intDayCol-4, 0, ii) & "AssignmentID" & "  " 	
'		Response.Write varTimesheet(intDayCol-3, 0, ii) & "pName" & "  " 	
'		Response.Write varTimesheet(intDayCol-2, 0, ii) & "fg" & "  "
'		Response.Write varTimesheet(intDayCol-1, 0, ii) & "fgRemove" & "  "
'		Response.Write varTimesheet(intDayCol, 0, ii) & "fgActive" & "<br><br>" 	
'	Next
'	
'	For ii=0 To Ubound(varEvent,3)
'		Response.Write "<b>" & varEvent(0, 0, ii) & "</b>" & "  "
'		For kk=1 to intDayNum
'			Response.Write kk & "date " & varEvent(kk, 0, ii) & "Off" & "  "
'			Response.Write varEvent(kk, 1, ii) & "Over" & "  "
'			Response.Write varEvent(kk, 2, ii) & "Notes" & "  "
'		Next
'		Response.Write varEvent(intDayNum+1, 0, ii) & "SubTotal" & "  " 	
'		Response.Write varEvent(intDayNum+2, 0, ii) & "EventID" & "<br><br>" 	
'	Next
'
'	Response.End
'	
'--------------------------------------------------										
				Else
					strError = objDatabase.strMessage	
					tmsInitialForTP = "4. " & strError
					Exit Function 	
				End If
			End If	
		Else
			strError = objDatabase.strMessage
			tmsInitialForTP = "5. " & strError
			Exit Function 	
		End If
	Else
		strError = objDatabase.strMessage
		tmsInitialForTP = "6. " & strError
		Exit Function 	
	End If

'--------------------------------------------------
' Free variables
'--------------------------------------------------
	
	objDatabase.dbDisConnect()																' Disconnect to SQL database	
	
	If isarray(varTimesheet) Then
		erase varTimesheet
	End If
	If isarray(varTMS1)	Then
		erase varTMS1
	End If	
	If isarray(varTMS2) Then
		erase varTMS2
	End If
	If isarray(varEvent) Then
		erase varEvent
	End If	
	If isarray(varEvent1) Then
		erase varEvent1
	End If
	If isarray(varEvent2) Then
		erase varEvent2
	End If
				
End Function

'**************************************************
' Function: tmsWriteHourforTP
' Description: 
'			  -	Enters working hour for a user  
'
' Parameters: - intUserID	: Integer
'			  - intMonth	: Integer 
'			  - intYear		: Integer
'			  - Col			: Integer
'			  - Row			: Integer
'			  - strType		: String
'			  - txtOffHour	: TextBox
'			  - txtOTNormal	: TextBox
'			  - txtOTNight	: TextBox
'			  - txtNote		: TextBox
'			 		
' Return value: Error message if there are any errors
' Author: Phan Thi Hong
' Date: 26/07/2001
' Note:
'**************************************************

Function tmsWriteHourforTP(ByVal intUserID, ByVal Col, ByVal Row, ByVal strType, ByVal intMonth, ByVal intYear, txtOffHour, txtOTNormal,txtOTNight, txtNote)
	
	Dim dblNormalHour, dblOverHour, intDayNum, dblOldNormalHour, dblOldOverHour, dblOffHour, dblHourOff, dblHourTotal, dblOverRate 
	Dim dblOTNormal,dblOTNight
	Dim strDate, strNote, strConnect, strTableTMS, objDatabase, strError1, strError2

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	intRow				= -1
	eRow				= -1
	dblOldNormalHour	= 0
	dblOldOverHour		= 0
	dblNormalHour		= 0
	dblOverHour			= 0
	dblOverRate			= 0
	dblHourOff			= 0
	dblHourTotal		= 0
	
	dblOTNormal			= 0
	dblOTNight			= 0
	
	intDayNum			= GetDay(intMonth,intYear)							' Numbers of days in a month
	dblOffHour			= 8.5							' User's working hour

	strTableTMS			= selectTable(intYear)								' Select table to view timesheet

	varTimesheet		= session("varTimesheet")							' Array stores timesheet data
	varEvent			= session("varEvent")								' Array stores event data
	
	If isarray(varTimesheet) Then
		intRow	= Ubound(varTimesheet,3)
	End If
	
	If isarray(varEvent) Then
		eRow	= Ubound(varEvent,3)
	End If

	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

'--------------------------------------------------
' End Of initializing variables
'--------------------------------------------------

'--------------------------------------------------
' Check table timesheet if it exists or not	
'--------------------------------------------------

	If objDatabase.dbConnect(strConnect) Then
		strSQL = "SELECT ISNULL(OBJECT_ID('" & strTableTMS & "'),0) AS TableName"

		If (objDatabase.runQuery(strSQL)) Then
			If objDatabase.getColumn_by_name("TableName") = 0 Then
				strError = "No data for your request."
				tmsWriteHourforTP = strError
				Exit Function 	
			End If
			objDatabase.closeRec	
		Else
			strError = objDatabase.strMessage
			tmsWriteHourforTP = strError
			Exit Function 	
		End If
	Else
		strError = objDatabase.strMessage
		tmsWriteHourforTP = strError
		Exit Function 	
	End If
	
'--------------------------------------------------
' Check value of Normal Hour textbox
'--------------------------------------------------
			
	If replace(Request.Form(txtOffHour)," ","") <> "" Then
		
		dblNormalHour = CDbl(formatnumber(Request.Form(txtOffHour),2))
			
		If strType = "P" Then
			dblOldNormalHour = CDbl(varTimesheet(Col, 0, Row))
			varTimesheet(Col, 0, Row) = CDbl(formatnumber(Request.Form(txtOffHour),2))
		ElseIf strType = "E" Then
			dblOldNormalHour = CDbl(varEvent(Col, 0, Row))
			varEvent(Col, 0, Row) = CDbl(formatnumber(Request.Form(txtOffHour),2))
		End If	
			
	ElseIf Request.Form(txtOffHour) = "" Or Request.Form(txtOffHour) = " " Then
		
		dblNormalHour = "0"
			
		If strType = "P" Then
			dblOldNormalHour = CDbl(varTimesheet(Col, 0, Row))
			varTimesheet(Col, 0, Row) = 0
		ElseIf strType = "E" Then
			dblOldNormalHour = CDbl(varEvent(Col, 0, Row))
			varEvent(Col, 0, Row) = 0
		End If	
			
	End If	

'--------------------------------------------------
' End of checking value of Normal Hour textbox
'--------------------------------------------------

'--------------------------------------------------
' Check value of Overime Hour textbox
'--------------------------------------------------


	'OT for normal rate: before 9:00PM
	dblOTNormal=0	
	If replace(Request.Form(txtOTNormal)," ","") <> "" Then		
		dblOTNormal = CDbl(formatnumber(Request.Form(txtOTNormal),2))					
	End If
	'OT for night rate: after 9:00PM	
	dblOTNight=0
	If replace(Request.Form(txtOTNight)," ","") <> "" Then		
		dblOTNight = CDbl(formatnumber(Request.Form(txtOTNight),2))
	End If
	
	dblOverHour=dblOTNormal + dblOTNight
	
	If strType = "P" Then
		dblOldOverHour = Cdbl(varTimesheet(Col, 1, Row))
		varTimesheet(Col, 1, Row) = CDbl(formatnumber(dblOverHour,2))
		varTimesheet(Col, 2, Row) = CDbl(formatnumber(dblOTNormal,2))	
		varTimesheet(Col, 3, Row) = CDbl(formatnumber(dblOTNight,2))
	ElseIf strType = "E" Then
		dblOldOverHour = CDbl(varEvent(Col, 1, Row))
		varEvent(Col, 1, Row) = CDbl(formatnumber(dblOverHour,2))
		varEvent(Col, 2, Row) = CDbl(formatnumber(dblOTNormal,2))
		varEvent(Col, 3, Row) = CDbl(formatnumber(dblOTNight,2))
	End If		

'--------------------------------------------------
' End of checking value of Overime Hour textbox
'--------------------------------------------------
		
'--------------------------------------------------
' Check value of Note textbox
'--------------------------------------------------

	If Request.Form(txtNote) <> "" Then	
		strNote = "'" & replace(trim(Request.Form(txtNote)),"'","''") & "'"
		
		If strType = "P" Then
			varTimesheet(Col, 4, Row) = trim(Request.Form(txtNote))
		ElseIf strType = "E" Then
			varEvent(Col, 4, Row) = trim(Request.Form(txtNote))
		End If	

	Else
		strNote = "Null"

		If strType = "P" Then
			varTimesheet(Col, 4, Row) = ""
		ElseIf strType = "E" Then
			varEvent(Col, 4, Row) = ""
		End If	

	End If

'--------------------------------------------------
' End of checking value of Note textbox
'--------------------------------------------------

	strDate = CDate(CStr(intMonth) & "/" & Col & "/" & CStr(intYear))
	'If Weekday(strDate) = 1 Or isHoliday(Col) >= 0 Then
	'	dblOverRate = dblOverHour * 2
	'ElseIf Weekday(strDate) = 7 Then	
	'	dblOverRate = dblOverHour * 1.5
	'Else
	'	dblOverRate = dblOverHour
	'End If
		
'--------------------------------------------------
' Analyse and check data to insert/update/delete for projects   
'--------------------------------------------------

	For ii = 0 To intRow
		dblHourTotal = dblHourTotal + CDbl(varTimesheet(Col, 0, ii)) + CDbl(varTimesheet(Col, 1, ii))
		dblHourOff = dblHourOff + CDbl(varTimesheet(Col, 0, ii))
	Next

	if CCur(dblHourOff)>8 then
		strError1 = "Total of working hours for project: " & CStr(dblHourOff) & ".<br>Can't be over 8 hours per day"
	else
	
		For kk = 0 To eRow-3
			dblHourTotal = dblHourTotal + CDbl(varEvent(Col, 0, kk)) + CDbl(varEvent(Col, 1, kk))	
			dblHourOff = dblHourOff + CDbl(varEvent(Col, 0, kk))	
		Next

		If CCur(dblHourOff) > CCur(dblOffHour) Then
			strError1 = "Total of working office hours: " & CStr(dblHourOff) & ".<br>Can't be over " & CStr(dblOffHour) & " hours per day"
		End If
	
		If CCur(dblHourTotal) > 24 Then
			strError2 = "Total of working hours: " & CStr(dblHourTotal) & ".<br>Can't be over 24 hours per day"
		End If
	end if

	
	If strError1 <> "" Then
		strError = strError1
	End If
	If strError2 <> "" Then
		strError = strError & "@@" & strError2
	End If
	
	If strError <> "" Then			
		tmsWriteHourforTP = strError
		Exit Function
	End If
				
	If strType = "P" Then					' For Projects
		intAssignmentID = varTimesheet(intDayNum+2, 0, Row)
		
'********Thao insert this paragraph	******								
		'strSQL = "INSERT INTO trace(Userid, Staffid, tdate, tnow, TType, Ip,AssignmentID) VALUES(" & session("USERID") & "," & intuserid & ",'" & strdate & "','" & now() & "', 1, '" & Request.ServerVariables("REMOTE_ADDR") & "'," & intAssignmentID & ")"
		'tmp = objDatabase.runActionQuery(strSQL)
'******************************************		
		
	
		If objDatabase.dbConnect(strConnect) Then
			
			strSQL = "SELECT * FROM " & strTableTMS & " WHERE TDate='" & strDate & _
						"' AND StaffID=" & intUserID & " AND AssignmentID=" & intAssignmentID & _
						" AND EventID=1"

'--------------------------------------------------
' Check data if this record exists or not. 
' If it existed, it would be updated. 
' Or a new record would be inserted if it didn't exist. 
'--------------------------------------------------
				
			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then								' Update

					If dblNormalHour = "0" And dblOverHour = "0" Then
						
						strSQL = "DELETE FROM " & strTableTMS & " WHERE TDate='" & strDate & "' AND StaffID=" & intUserID & _
								 " AND AssignmentID=" & intAssignmentID & " AND EventID=1"

'--------------------------------------------------
' Update values of Total Column, Total Row, Total Column of Total Row,
' Normal hour Row, Overtime hour Row, Total Column of Normal hour Row,
' Total Column of OverTime hour Row.
'--------------------------------------------------
 						
						varTimesheet(intDayNum+1, 0, Row) = CDbl(varTimesheet(intDayNum+1, 0, Row)) - dblOldNormalHour - dblOldOverHour								' Total Column
						
						varEvent(Col, 0, eRow-2) = CDbl(varEvent(Col, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour												' Total Row
						varEvent(intDayNum+1, 0, eRow-2) = CDbl(varEvent(intDayNum+1, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour								' Total Column of Total Row
	
						varEvent(Col, 0, eRow-1) = CDbl(varEvent(Col, 0, eRow-1)) - dblOldNormalHour																' Normal hour Row	
						varEvent(Col, 0, eRow) = CDbl(varEvent(Col, 0, eRow)) - dblOldOverHour																		' OverTime hour Row	
						
						varEvent(intDayNum+1, 0, eRow-1) = CDbl(varEvent(intDayNum+1, 0, eRow-1)) - dblOldNormalHour												' Total Column of Normal hour Row	
						varEvent(intDayNum+1, 0, eRow)   = CDbl(varEvent(intDayNum+1, 0, eRow)) - dblOldOverHour													' Total Column of OverTime hour Row	

'--------------------------------------------------
' End Of Update
'--------------------------------------------------
						
					Else
						strSQL = "UPDATE " & strTableTMS & " SET Hours=" & dblNormalHour & ", OverTime=" & dblOverHour & ", OverRate=" & dblOverRate & _
								 ", OTNight=" & dblOTNight & ",OTNormal=" & dblOTNormal & ", Note=" & strNote & " WHERE TDate='" & strDate & "' AND StaffID=" & intUserID & _
								 " AND AssignmentID=" & intAssignmentID & " AND EventID=1"

'--------------------------------------------------
' Update values of Total Column, Total Row, Total Column of Total Row,
' Normal hour Row, Overtime hour Row, Total Column of Normal hour Row,
' Total Column of OverTime hour Row.				
'--------------------------------------------------

						varTimesheet(intDayNum+1, 0, Row) = CDbl(varTimesheet(intDayNum+1, 0, Row)) - dblOldNormalHour - dblOldOverHour + dblNormalHour + dblOverHour
						
						varEvent(Col, 0, eRow-2)	= CDbl(varEvent(Col, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour + dblNormalHour + dblOverHour					' Total Row
						varEvent(intDayNum+1, 0, eRow-2) = CDbl(varEvent(intDayNum+1, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour + dblNormalHour + dblOverHour		' Total Column of Total Row

						varEvent(Col, 0, eRow-1)	= CDbl(varEvent(Col, 0, eRow-1)) - dblOldNormalHour + dblNormalHour													' Normal hour Row	
						varEvent(Col, 0, eRow)	    = CDbl(varEvent(Col, 0, eRow)) - dblOldOverHour + dblOverHour														' OverTime hour Row	
	
						varEvent(intDayNum+1, 0, eRow-1) = CDbl(varEvent(intDayNum+1, 0, eRow-1)) - dblOldNormalHour + dblNormalHour									' Total Column of Normal hour Row	
						varEvent(intDayNum+1, 0, eRow)   = CDbl(varEvent(intDayNum+1, 0, eRow)) - dblOldOverHour + dblOverHour											' Total Column of OverTime hour Row	

'--------------------------------------------------
' End Of Update
'--------------------------------------------------
						
					End If		
											 
				Else																' Insert a new record	
					
					strSQL = "INSERT INTO " & strTableTMS & "(TDate, StaffID, AssignmentID, EventID, Hours, OverTime, OTNight,OTNormal,OverRate, Note) VALUES('" & _
								strDate & "', " & intUserID & ", " & intAssignmentID & ", 1, " & dblNormalHour & ", " & dblOverHour & ", " & dblOTNight & ", " & dblOTNormal & ", " & dblOverRate & _
								", " & strNote & ")"

'--------------------------------------------------
' Update values of Total Column, Total Row, Total Column of Total Row,
' Normal hour Row, Overtime hour Row, Total Column of Normal hour Row,
' Total Column of OverTime hour Row.				
'--------------------------------------------------

					varTimesheet(intDayNum+1, 0, Row) = CDbl(varTimesheet(intDayNum+1, 0, Row)) + dblNormalHour + dblOverHour

					varEvent(Col, 0, eRow-2) = CDbl(varEvent(Col, 0, eRow-2)) + dblNormalHour + dblOverHour																' Total Row
					varEvent(intDayNum+1, 0, eRow-2) = CDbl(varEvent(intDayNum+1, 0, eRow-2)) + dblNormalHour + dblOverHour												' Total Column of Total Row

					varEvent(Col, 0, eRow-1) = CDbl(varEvent(Col, 0, eRow-1)) + dblNormalHour																			' Normal hour Row	
					varEvent(Col, 0, eRow) = CDbl(varEvent(Col, 0, eRow)) + dblOverHour																					' OverTime hour Row	

					varEvent(intDayNum+1, 0, eRow-1) = CDbl(varEvent(intDayNum+1, 0, eRow-1)) + dblNormalHour															' Total Column of Normal hour Row	
					varEvent(intDayNum+1, 0, eRow)   = CDbl(varEvent(intDayNum+1, 0, eRow)) + dblOverHour																' Total Column of OverTime hour Row	

'--------------------------------------------------
' End Of Update
'--------------------------------------------------
					
				End If



				If objDatabase.runActionQuery(strSQL) = False Then
					strError = objDatabase.strMessage
					tmsWriteHourforTP = strError
					Exit Function
				End If
				
'********Chi insert this paragraph	******								
		strSQL = "INSERT INTO trace2004(Userid, Staffid, tdate, tnow, TType, Ip,AssignmentID,Act,OfficeHour,Overtime) " & _
					"VALUES(" & session("USERID") & "," & intuserid & ",'" & strdate & "','" & now() & "', 1, '" & Request.ServerVariables("REMOTE_ADDR") & "'," & intAssignmentID &_
							",'" & Left(strSQL,3) & "'," & dblNormalHour & "," & dblOverHour & ")"
		tmp = objDatabase.runActionQuery(strSQL)
'******************************************						
						
			Else
				strError = objDatabase.strMessage
				tmsWriteHourforTP = strError
				Exit Function
			End If	
				
		Else
			strError = objDatabase.strMessage
			tmsWriteHourforTP = strError
			Exit Function
		End If
			
		objDatabase.dbDisConnect()

'--------------------------------------------------
' Analyse and check data to insert/update/delete for events
'--------------------------------------------------
			
	ElseIf strType = "E" Then														' For Events

		intAssignmentID = varEvent(intDayNum+2, 0, Row)

		strConnect = Application("g_strConnect")									' Connection string 				
		Set objDatabase = New clsDatabase 

		If objDatabase.dbConnect(strConnect) Then

'********Thao insert this paragraph	******								
		'strSQL = "INSERT INTO trace(Userid, Staffid, tdate, tnow, TType, Ip) VALUES(" & session("USERID") & "," & intuserid & ",'" & strdate & "','" & now() & "', 2,'" & Request.ServerVariables("REMOTE_ADDR") & "')"
		'tmp = objDatabase.runActionQuery(strSQL)
'******************************************					

			strSQL = "SELECT * FROM " & strTableTMS & " WHERE TDate='" & strDate & _
						"' AND StaffID=" & intUserID & " AND EventID=" & intAssignmentID & _
						" AND AssignmentID=1"

'--------------------------------------------------
' Check data if this record exists or not. 
' If it existed, it would be updated. 
' Or a new record would be inserted if it didn't exist. 
'--------------------------------------------------
				
			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then								' Update

					If dblNormalHour = "0" And dblOverHour = "0" Then

						strSQL = "DELETE FROM " & strTableTMS & " WHERE TDate='" & strDate & "' AND StaffID=" & intUserID & _
								 " AND EventID=" & intAssignmentID & " AND AssignmentID=1"

'--------------------------------------------------
' Update values of Total Column, Total Row, Total Column of Total Row,
' Normal hour Row, Overtime hour Row, Total Column of Normal hour Row,
' Total Column of OverTime hour Row.				
'--------------------------------------------------
			
						varEvent(intDayNum+1, 0, Row) = (varEvent(intDayNum+1, 0, Row) - dblOldNormalHour - dblOldOverHour)

						varEvent(Col, 0, eRow-2) = CDbl(varEvent(Col, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour												' Total Row
						varEvent(intDayNum+1, 0, eRow-2) = CDbl(varEvent(intDayNum+1, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour								' Total Column of Total Row
		
						If trim(varEvent(0, 0, Row)) = "Personal Time" Or trim(varEvent(0, 0, Row)) = "General/Admin" Then
							varEvent(Col, 0, eRow-1) = CDbl(varEvent(Col, 0, eRow-1)) - dblOldNormalHour															' Normal hour Row	
							varEvent(Col, 0, eRow) = CDbl(varEvent(Col, 0, eRow)) - dblOldOverHour																	' OverTime hour Row	

							varEvent(intDayNum+1, 0, eRow-1) = CDbl(varEvent(intDayNum+1, 0, eRow-1)) - dblOldNormalHour											' Total Column of Normal hour Row	
							varEvent(intDayNum+1, 0, eRow)   = CDbl(varEvent(intDayNum+1, 0, eRow)) - dblOldOverHour												' Total Column of OverTime hour Row	
						End If					 
'--------------------------------------------------
' End Of Update
'--------------------------------------------------
						
					Else
						
						strSQL = "UPDATE " & strTableTMS & " SET Hours=" & dblNormalHour & ", OverTime=" & dblOverHour & ", OverRate=" & dblOverRate & _
								 ", OTNight=" & dblOTNight & ",OTNormal=" & dblOTNormal & ", Note=" & strNote & " WHERE TDate='" & strDate & "' AND StaffID=" & intUserID & _
								 " AND EventID=" & intAssignmentID & " AND AssignmentID=1"

'--------------------------------------------------
' Update values of Total Column, Total Row, Total Column of Total Row,
' Normal hour Row, Overtime hour Row, Total Column of Normal hour Row,
' Total Column of OverTime hour Row.				
'--------------------------------------------------

						varEvent(intDayNum+1, 0, Row) = ((varEvent(intDayNum+1, 0, Row) - dblOldNormalHour - dblOldOverHour) + dblNormalHour + dblOverHour)

						varEvent(Col, 0, eRow-2)	= CDbl(varEvent(Col, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour + dblNormalHour + dblOverHour				' Total Row
						varEvent(intDayNum+1, 0, eRow-2) = CDbl(varEvent(intDayNum+1, 0, eRow-2)) - dblOldNormalHour - dblOldOverHour + dblNormalHour + dblOverHour	' Total Column of Total Row
						
						If trim(varEvent(0, 0, Row)) = "Personal Time" Or trim(varEvent(0, 0, Row)) = "General/Admin" Then
							varEvent(Col, 0, eRow-1)	= CDbl(varEvent(Col, 0, eRow-1)) - dblOldNormalHour + dblNormalHour											' Normal hour Row	
							varEvent(Col, 0, eRow)	= CDbl(varEvent(Col, 0, eRow)) - dblOldOverHour + dblOverHour													' OverTime hour Row	

							varEvent(intDayNum+1, 0, eRow-1) = CDbl(varEvent(intDayNum+1, 0, eRow-1)) - dblOldNormalHour + dblNormalHour							' Total Column of Normal hour Row	
							varEvent(intDayNum+1, 0, eRow)   = CDbl(varEvent(intDayNum+1, 0, eRow)) - dblOldOverHour + dblOverHour									' Total Column of OverTime hour Row	
						End If		
						
'--------------------------------------------------
' End Of Update
'--------------------------------------------------						

					End If

				Else																' Insert a new record	
					
					strSQL = "INSERT INTO " & strTableTMS & "(TDate, StaffID, AssignmentID, EventID, Hours, OverTime,OTNight,OTNormal, OverRate, Note) VALUES('" & _
								strDate & "', " & CInt(intUserID) & ", 1, " & intAssignmentID & ", " & dblNormalHour & ", " & dblOverHour & ", " & dblOTNight & ", " & dblOTNormal & ", " & dblOverRate & _
								", " & strNote & ")"

'--------------------------------------------------
' Update values of Total Column, Total Row, Total Column of Total Row,
' Normal hour Row, Overtime hour Row, Total Column of Normal hour Row,
' Total Column of OverTime hour Row.				
'--------------------------------------------------

					varEvent(intDayNum+1, 0, Row) = CDbl(varEvent(intDayNum+1, 0, Row)) + dblNormalHour + dblOverHour

					varEvent(Col, 0, eRow-2) = CDbl(varEvent(Col, 0, eRow-2)) + dblNormalHour + dblOverHour															' Total Row
					varEvent(intDayNum+1, 0, eRow-2) = CDbl(varEvent(intDayNum+1, 0, eRow-2)) + dblNormalHour + dblOverHour											' Total Column of Total Row

					If trim(varEvent(0, 0, Row)) = "Personal Time" Or trim(varEvent(0, 0, Row)) = "General/Admin" Then
						varEvent(Col, 0, eRow-1) = CDbl(varEvent(Col, 0, eRow-1)) + dblNormalHour																	' Normal hour Row	
						varEvent(Col, 0, eRow) = CDbl(varEvent(Col, 0, eRow)) + dblOverHour																			' OverTime hour Row	

						varEvent(intDayNum+1, 0, eRow-1) = CDbl(varEvent(intDayNum+1, 0, eRow-1)) + dblNormalHour													' Total Column of Normal hour Row	
						varEvent(intDayNum+1, 0, eRow)   = CDbl(varEvent(intDayNum+1, 0, eRow)) + dblOverHour														' Total Column of OverTime hour Row	
					End If						
					
'--------------------------------------------------
' End Of Update
'--------------------------------------------------					
				End If
'Response.Write strsql
'Response.End
				If objDatabase.runActionQuery(strSQL) = False Then
					strError = objDatabase.strMessage
					tmsWriteHourforTP = strError
					Exit Function
				End If
'********Chi insert this paragraph	******								
		strSQL = "INSERT INTO trace2004(Userid, Staffid, tdate, tnow, TType, Ip,AssignmentID,Act,OfficeHour,Overtime) " & _
					"VALUES(" & session("USERID") & "," & intuserid & ",'" & strdate & "','" & now() & "', 2, '" & Request.ServerVariables("REMOTE_ADDR") & "'," & intAssignmentID &_
							",'" & Left(strSQL,3) & "'," & dblNormalHour & "," & dblOverHour & ")"
		tmp = objDatabase.runActionQuery(strSQL)
'******************************************				
			Else
				strError = objDatabase.strMessage
				tmsWriteHourforTP = strError
				Exit Function
			End If	
				
		Else
			strError = objDatabase.strMessage
			tmsWriteHourforTP = strError
			Exit Function
		End If
			
		objDatabase.dbDisConnect()
		
	End If

	
'--------------------------------------------------
' Free variables
'--------------------------------------------------

	session("varTimesheet") = varTimesheet
	session("varEvent") = varEvent

	Set varTimesheet = Nothing
	Set varEvent = Nothing

End Function

'**************************************************
' Function: tmsUpdateSubtaskforTP
' Description: 
'			  -	Remove a subtask from varTimesheet array
'
' Parameters: - intNewAssignmentID	: Integer
'			  - intNewAssignmentID	: Integer
'			  - intMonth			: Integer
'			  - intYear				: Integer	
'			 		
' Return value: Error message if there are any errors
' Author: Phan Thi Hong
' Date: 02/08/2001
' Note:
'**************************************************

Function tmsUpdateSubtaskforTP(ByVal intOldAssignmentID, ByVal intNewAssignmentID, ByVal intMonth, ByVal intYear, ByVal intUserID)
	Dim strTableTMS, objDatabase, strError, intDayNum, varTimesheet
	Dim ii, fgExist
	
	fgExist		= 0
	strTableTMS	= selectTable(intYear)								' Select table to view timesheet
	intDayNum	= GetDay(intMonth,intYear)

'--------------------------------------------------
' Check table timesheet if it exists or not	
'--------------------------------------------------

		strConnect = Application("g_strConnect")									' Connection string 				
		Set objDatabase = New clsDatabase 

		If objDatabase.dbConnect(strConnect) Then
			strSQL = "SELECT ISNULL(OBJECT_ID('" & strTableTMS & "'),0) AS TableName"

			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.getColumn_by_name("TableName") = 0 Then
					strError = "No data for your request."
					tmsUpdateSubtaskforTP = strError
					Exit Function 	
				End If
				objDatabase.closeRec	
			Else
				strError = objDatabase.strMessage
				tmsUpdateSubtaskforTP = strError
				Exit Function 	
			End If
		Else
			strError = objDatabase.strMessage
			tmsUpdateSubtaskforTP = strError
			Exit Function 	
		End If	

'--------------------------------------------------
' Check if this assignmentid exist or not
'--------------------------------------------------

	varTimesheet = session("varTimesheet")

	For ii = 0 To Ubound(varTimesheet,3)
		If varTimesheet(intDayNum+2,0,ii) = intNewAssignmentID And CInt(varTimesheet(intDayNum+5, 0, ii)) = 0 Then
			fgExist = 1
			Exit For
		End If	
	Next
	
	If fgExist = 1 Then
		tmsUpdateSubtaskforTP = "Data has already been inputted. Choose another one."
		Exit Function
	Else	

		If objDatabase.dbConnect(strConnect) Then

			strSQL = "UPDATE " & strTableTMS & " SET AssignmentID=" & intNewAssignmentID & _
					 " WHERE AssignmentID=" & intOldAssignmentID & " AND StaffID=" & intUserID & _
					 " AND Month(TDate)=" & intMonth & " AND EventID=1"

			If Not objDatabase.runActionQuery(strSQL) Then
				strError = objDatabase.strMessage
				tmsUpdateSubtaskforTP = strError
				Exit Function
			End If

		Else
			strError = objDatabase.strMessage
			tmsUpdateSubtaskforTP = strError
			Exit Function
		End If
				
		objDatabase.dbDisConnect()
	End If
	
End Function

'**************************************************
' Function: tmsAddsubtaskforTP
' Description: 
'			  -	Add subtask to write timesheet
'
' Parameters: - strPID			: String
'			  - strPName		: String
'			  - intParentID		: Integer
'			  - strSubTask		: String
'			  - intAssignmentID	: Integer
'			  - intMonth		: Integer
'			  - intYear			: Integer	
'			 		
' Return value: Error message if there are any errors
' Author: Phan Thi Hong
' Date: 30/07/2001
' Note:
'**************************************************

Function tmsAddsubtaskforTP(ByVal strPID, ByVal strPName, ByVal intParentID, ByVal strSubTask, ByVal intAssignmentID, ByVal intMonth, ByVal intYear)

	Dim intDayNum, intDayCol, intDayCount, intOldRow, intNewRow, ii, kk, intCurRow
	Dim fgExist, varTemp
	
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	intOldRow	= -1
	intCurRow	= -1
	fgExist		= 0
	intDayNum	= GetDay(intMonth,intYear)													' Numbers of days in a month
	intDayCol	= intDayNum + 6

	strFirstDay = FirstOfMonth(intMonth,intYear)											' Get the first day in a month				
	intDayCount	= curDayNum(intDayNum,strFirstDay)											' Numbers of days since the first day in month to now

'--------------------------------------------------
' End of initializing variables
'--------------------------------------------------

	varTimesheet = session("varTimesheet")
	
	
	If IsArray(varTimesheet) Then
		intOldRow = Ubound(varTimesheet,3)
	End If

'--------------------------------------------------
' Check if this assignmentid exists in the timesheet array or not	
'--------------------------------------------------

	If intParentID = 0 Then																	' Project no has subtask
		intCheckRow = ""
		
		If intOldRow >= 0 Then
			For ii = 0 To intOldRow
				If trim(varTimesheet(0, 0, ii)) = trim(strPID) Then
					strCheckProject = strPID
					fgExist = 1
					intCheckRow = ii
					Exit For
				End If
			Next
		End If
		
		If fgExist = 0 Then
			intNewRow = intOldRow + 1
			If intOldRow >= 0 Then
				Redim Preserve varTimesheet(intDayCol, 4, intNewRow)
			Else
				Redim varTimesheet(intDayCol, 4, intNewRow)	
			End If
			
			varTimesheet(0, 0, intNewRow)			= strPID								' Field ProjectID		
			For kk = 1 To intDayNum					
				varTimesheet(kk, 0, intNewRow)		= 0										' Initialize office hour		
				varTimesheet(kk, 1, intNewRow)		= 0										' Initialize overtime hour	
				varTimesheet(kk, 2, intNewRow)		= 0										' Initialize overtime hour
				varTimesheet(kk, 3, intNewRow)		= 0										' Initialize overtime hour	
				varTimesheet(kk, 4, intNewRow)		= ""									' Initialize note
			Next

			varTimesheet(intDayCol-5, 0, intNewRow) = 0										' Working-Hour's SubTotal										
			varTimesheet(intDayCol-4, 0, intNewRow) = intAssignmentID						' Field AssignmentID								
			varTimesheet(intDayCol-3, 0, intNewRow) = strPName								' Field Project Name
			varTimesheet(intDayCol-2, 0, intNewRow)	= "P"									' "P"-Project; "S"-SubTask; "N"-None
			varTimesheet(intDayCol-1, 0, intNewRow)	= 0										' 0 - Not Remove; 1 - Remove
			varTimesheet(intDayCol, 0, intNewRow)	= 1

			session("varTimesheet")					= varTimesheet
			Set varTimesheet = Nothing				 
		Else
			If CInt(varTimesheet(intDayCol-1, 0, intCheckRow)) = 1 Then
				varTimesheet(intDayCol-1, 0, intCheckRow) = 0
				If Trim(varTimesheet(intDayCol-2, 0, intCheckRow-1)) = "N" Then
					varTimesheet(intDayCol-1, 0, intCheckRow-1) = 0
				End If

				session("varTimesheet") = varTimesheet
			Else	
				tmsAddsubtaskforTP = strCheckProject
				Exit Function
			End if	
		End If			
	Else																					' Project has subtask
		intCheckRow = ""
		
		If intOldRow >= 0 Then	
			For ii = 0 To intOldRow
				If varTimesheet(intDayCol-4, 0, ii) = intAssignmentID Then
					strCheckProject = intAssignmentID
					fgExist = 1
					intCheckRow = ii
					Exit For
				ElseIf trim(varTimesheet(0, 0, ii)) = Trim(strPID) Then
					intCurRow = ii
					fgExist = 2	
				End If
			Next
		End If
		
		If fgExist = 0 Then																	' This Project and AssignmentID doesn't exist in varTimesheet array	

' For Project Name
			intNewRow = intOldRow + 1
			If intOldRow >= 0 Then
				Redim Preserve varTimesheet(intDayCol, 4, intNewRow)
			Else
				Redim varTimesheet(intDayCol, 4, intNewRow)	
			End If
			
			varTimesheet(0,0,intNewRow) = strPID											' Field ProjectID	

			For kk = 1 To intDayNum									
				varTimesheet(kk,0,intNewRow) = 0											' Initialize office hour
				varTimesheet(kk,1,intNewRow) = 0											' Initialize overtime hour normal rate
				varTimesheet(kk,2,intNewRow) = 0											' Initialize overtime hour night rate
				varTimesheet(kk,3,intNewRow) = 0											' Initialize overtime hour
				varTimesheet(kk,4,intNewRow) = ""											' Initialize note
			Next
			varTimesheet(intDayCol-5,0,intNewRow)	= 0										' Working-Hour's SubTotal
			varTimesheet(intDayCol-4,0,intNewRow)	= ""									' Field AssignmentID
			varTimesheet(intDayCol-3,0,intNewRow)	= strPName								' Field ProjectName
			varTimesheet(intDayCol-2,0,intNewRow)	= "N"									' "P"-Project; "S"-SubTask; "N"-None 
			varTimesheet(intDayCol-1, 0, intNewRow) = 0										' 0 - Not Remove; 1 - Remove
			varTimesheet(intDayCol, 0, intNewRow)	= 0

' For Sub-Task Name			
			intNewrow = Ubound(varTimesheet,3) + 1
			Redim Preserve varTimesheet(intDayCol, 4, intNewRow)
			
			varTimesheet(0, 0, intNewRow) = strSubTask										' Field SubTask Name
			For kk = 1 to intDayNum									
				varTimesheet(kk,0,intNewRow) = 0											' Initialize office hour
				varTimesheet(kk,1,intNewRow) = 0											' Initialize overtime hour normal rate
				varTimesheet(kk,2,intNewRow) = 0											' Initialize overtime hour night rate
				varTimesheet(kk,3,intNewRow) = 0											' Initialize overtime hour
				varTimesheet(kk,4,intNewRow) = ""											' Initialize note
			Next

			varTimesheet(intDayCol-5, 0, intNewRow) = 0										' Working-Hour's SubTotal		
			varTimesheet(intDayCol-4, 0, intNewRow) = intAssignmentID						' Field AssignmentID	
			varTimesheet(intDayCol-3, 0,intNewRow)  = strSubTask							' Field SubTask Name
			varTimesheet(intDayCol-2, 0, intNewRow) = "S"									' "P"-Project; "S"-SubTask; "N"-None 
			varTimesheet(intDayCol-1, 0, intNewRow)	= 0										' 0 - Not Remove; 1 - Remove
			varTimesheet(intDayCol, 0, intNewRow)	= 1
			
			session("varTimesheet") = varTimesheet	
					 
			Set varTimesheet = Nothing
			
		ElseIf fgExist = 2 Then																' This project has already exist but this AssignmentID does not exist
		
'--------------------------------------------------
' Divide varTimesheet array into two parts
' at the project which has subtask be added. 
' And push the first part into varTemp array
'--------------------------------------------------

'			Redim Preserve varTemp(intDayCol, 2, intCurRow)
			Redim varTemp(intDayCol, 4, intCurRow)
			For ii = 0 To intCurRow
				varTemp(0, 0, ii)	   = varTimesheet(0, 0, ii)						' Field ProjectID or SubTask Name

				For kk = 1 To intDayNum					
					varTemp(kk, 0, ii) = varTimesheet(kk, 0, ii)							' Office hour
					varTemp(kk, 1, ii) = varTimesheet(kk, 1, ii)							' OverTime hour
					varTemp(kk, 2, ii) = varTimesheet(kk, 2, ii)							' OverTime hour
					varTemp(kk, 3, ii) = varTimesheet(kk, 3, ii)							' OverTime hour
					varTemp(kk, 4, ii) = varTimesheet(kk, 4, ii)							' Note
				Next

				varTemp(intDayCol-5, 0, ii) = varTimesheet(intDayCol-5, 0, ii)				' Working-Hour's SubTotal
				varTemp(intDayCol-4, 0, ii) = varTimesheet(intDayCol-4, 0, ii)				' Field AssignmentID	
				varTemp(intDayCol-3, 0, ii) = varTimesheet(intDayCol-3, 0, ii)				' Field ProjectName		
				varTemp(intDayCol-2, 0, ii) = varTimesheet(intDayCol-2, 0, ii)				' "P"-Project; "S"-SubTask; "N"-None 
				varTemp(intDayCol-1, 0, ii)	= varTimesheet(intDayCol-1, 0, ii)				' 0 - Not Remove; 1 - Remove
				varTemp(intDayCol, 0, ii)	= varTimesheet(intDayCol, 0, ii)				' 0 - DeActivated; 1 - Activated
				
			Next

'--------------------------------------------------
' Append new subtask row to varTemp array
'--------------------------------------------------

			intNewRow = Ubound(varTemp,3) + 1
			Redim Preserve varTemp(intDayCol, 4, intNewRow)
			
			varTemp(0, 0, intNewRow) = strSubTask
			For kk = 1 To intDayNum					
				varTemp(kk, 0, intNewRow) = 0												' Office hour
				varTemp(kk, 1, ii) = 0														' OverTime hour
				varTemp(kk, 2, ii) = 0														' OverTime hour
				varTemp(kk, 3, ii) = 0														' OverTime hour
				varTemp(kk, 4, ii) = ""														' Note
			Next

			varTemp(intDayCol-5, 0, intNewRow) = 0											' Working-Hour's SubTotal
			varTemp(intDayCol-4, 0, intNewRow) = intAssignmentID							' Field AssignmentID	
			varTemp(intDayCol-3, 0, intNewRow) = strSubTask									' Field ProjectName		
			varTemp(intDayCol-2, 0, intNewRow) = "S"										' "P"-Project; "S"-SubTask; "N"-None 
			varTemp(intDayCol-1, 0, intNewRow) = 0											' 0 - Not Remove; 1 - Remove
			varTemp(intDayCol, 0, intNewRow)   = 1   	
			
'--------------------------------------------------
' Append rows which in varTimesheet array to varTemp array
'--------------------------------------------------
		
			intNRow = Ubound(varTemp,3) + (Ubound(varTimesheet,3) - intCurRow)			
			Redim Preserve varTemp(intDayCol,4, intNRow)
			For ii = intNewRow To Ubound(varTimesheet,3)
				varTemp(0, 0, ii+1)				= varTimesheet(0, 0, ii)					' Field ProjectID or SubTask Name	
				For kk = 1 To intDayNum					
					varTemp(kk, 0, ii+1)		= varTimesheet(kk, 0, ii)					' Office hour
					varTemp(kk, 1, ii+1)		= varTimesheet(kk, 1, ii)					' OverTime hour
					varTemp(kk, 2, ii+1)		= varTimesheet(kk, 2, ii)					' OverTime hour
					varTemp(kk, 3, ii+1)		= varTimesheet(kk, 3, ii)					' OverTime hour
					varTemp(kk, 4, ii+1)		= varTimesheet(kk, 4, ii)					' Note
				Next

				varTemp(intDayCol-5, 0, ii+1)	= varTimesheet(intDayCol-5, 0, ii)			' Working-Hour's SubTotal
				varTemp(intDayCol-4, 0, ii+1)	= varTimesheet(intDayCol-4, 0, ii)			' Field AssignmentID
				varTemp(intDayCol-3, 0, ii+1)	= varTimesheet(intDayCol-3, 0, ii)			' Field ProjectName	
				varTemp(intDayCol-2, 0, ii+1)	= varTimesheet(intDayCol-2, 0, ii)			' "P"-Project; "S"-SubTask; "N"-None 		
				varTemp(intDayCol-1, 0, ii+1)	= varTimesheet(intDayCol-1, 0, ii)			' 0 - Not Remove; 1 - Remove
				varTemp(intDayCol, 0, ii+1)		= varTimesheet(intDayCol, 0, ii)			' 0 - DeActivated; 1 - Activated
			Next			

'--------------------------------------------------
' Test data
'			
'	for k = 0 to ubound(varTemp,3)
'		Response.Write "<b>" & varTemp(0, 0, k) & "</b>" & "  "
'		for l=1 to intDayNum
'			Response.Write l & "date " & varTemp(l, 0, k) & "Off" & "  "
'			Response.Write varTemp(l, 1, k) & "Over" & "  "
'			Response.Write varTemp(l, 2, k) & "Notes" & "  "
'		next
'		Response.Write varTemp(intDayCol-4, 0, k) & " SubTotal" & "  " 	
'		Response.Write varTemp(intDayCol-3, 0, k) & " Assignid" & "  " 	
'		Response.Write varTemp(intDayCol-2, 0, k) & " Subtask" & "  " 	
'		Response.Write varTemp(intDayCol-1, 0, k) & " fgPro" & "  " 	
'		Response.Write varTemp(intDayCol, 0, k) & " fgDel" & "<br><br>" 	
'	next
'	Response.End
'
'--------------------------------------------------

			session("varTimesheet") = varTemp
	
		Else																				' This AssignmentID has already exist	 	
			If CInt(varTimesheet(intDayCol-1, 0, intCheckRow)) = 1 Then
				varTimesheet(intDayCol-1, 0, intCheckRow) = 0
				If Trim(varTimesheet(intDayCol-2, 0, intCheckRow-1)) = "N" Then
					varTimesheet(intDayCol-1, 0, intCheckRow-1) = 0
				End If

				session("varTimesheet") = varTimesheet
			Else	
				tmsAddsubtaskforTP = strCheckProject
				Exit Function
			End if	
		End If
	End If

End Function
</SCRIPT>
