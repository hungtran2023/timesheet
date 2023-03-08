<%
'--------------------------------------------------
'
'Task: This is a VBScript class that will generate 
' and write an HTML table filled by a RS. All you have to do is pass set its 
' connection and SQL properties. The table will be pageable and sortable by 
' clicking the column heads. You can set the column heads to text or images. You 
' can also set on of the columns to have a link (like a tradenum going to a trade 
' details page). This is a nice example of how to use a VBScript class. Pretty 
' well commented, any questions please email em to me: dgarlit@hotmail.com
'
'''''''''''''''''''''''''''''''''''''
'clsSortableTable
'
'PURPOSE:
'This is a VBScript class that will generate and write an HTML table filled by a RS.  
'All you have to do is pass set its connection and SQL properties.  The table will be pageable 
'and sortable by clicking the column heads.  You can set the column heads to text or images.  
'You can also set on of the columns to have a link (like a tradenum going to a trade details page). 
'This is a nice example of how to use a VBScript class.  Pretty well commented, any questions please
'email em to me: dgarlit@hotmail.com
'
'how to use it
'FIRST: include this file and create a new clsSortableTable object
'	dim objSortableTable	
'	set objSortableTable = new clsSortableTable  

'SECOND: set properties

' MUSTHAVES	

	'set the connection property to a connection object
'	objSortableTable.Connection = objConn

  'set the SQL property to a string of SQL
'	objSortableTable.SQL = "select mem_id, mem_last_name, mem_first_name, mem_name, mem_last_updated from mfiMember"

  'set the default sort property to a string of the field you want the default sort on
'	objSortableTable.DefaultSort = "mem_last_name"	

' OPTIONAL PROPERTIES

	'the displayedfields property can be set to an array of fields that you would like display, 
	'if you don't set this to an array, it will use all the select fields from the query
'	objSortableTable.DisplayedFields = array("mem_last_name", "mem_first_name", "mem_name", "mem_last_updated")

	'you can set this variable to the current page name, if you choose
	'not to set this variiable it will default and request the script name from the serverVariables collection
'	objSortableTable.PageName = "memberList.asp"

  'use the PageSize property to set the number of records per page, if none selected it will default to 10
'	objSortableTable.PageSize = 7

  'the TableAttributes property can be set to a string of HTML table attributes 
'	objSortableTable.TableAttributes = "border=1 cellpadding=1 cellspacing=0"

 'the HeaderRowAttributes property can be set to a string of HTML row attributes to be used for the header
'	objSortableTable.HeaderRowAttributes = "bgcolor='#cccccc'"

 'the HeaderCellAttributes property can be set to a string of HTML cell attributes to be used for the header
'	objSortableTable.HeaderCellAttributes = "align=center valign=bottom"

  'the RowAttributes property can be set to a string of HTML row attributes to be used for the records
'	objSortableTable.RowAttributes = "bgcolor=white"

  'the CellAttributes property can be set to a string of HTML cell attributes to be used for the records
'	objSortableTable.CellAttributes = "valign=top"

 'the FontTagAttributes property can be set to a string of HTML font tag attributes to be used for the records
'	objSortableTable.FontTagAttributes = "color=blue face='tahoma'"

 'the ImagePath property is the path to your image files, it defaults to "../img/" assuming
 'that the images are located in an folder called Img up a level from the current folder
'	objSortableTable.ImagePath = "../img/"

  'set the TextLink property to false if you want images,else it will default to text links that are based on
  'the DisplayedFields property  
'	objSortableTable.TextLink = true

  'PlainImages property can be set to an string array of images to be used in the header
'objSortableTable.PlainImages = array("<img src=firstname.gif>", "...

  'AscImages property can be set to an string array of images to be used in the header for Ascending sort
'objSortableTable.AscImages = array("<img src=firstname_asc.gif>", "...

	'DscImages property can be set to an string array of images to be used in the header for descending sort
'objSortableTable.DscImages = array("<img src=firstname_dsc.gif>", "...

 'NavString property can be set to an string array of images to be used for the navigation buttons
 'where index 0=first, 1=previous, 2=next, 3=last
'objSortableTable.NavString

'*the clsSortableTable has two properties it needs to sort and navigate, they are .Sort and .CurrentPage
'the sorting and navigating works by passing the querystring vairables page and sort to itself, so, if
'you were going to change these properties youd probably make em equal to request("sort") for sort and 
'request("page") for currentPage, however if you don't touch these the object will take care of it on its own.
'the point being though, that you can read these values
	
	'the LinkedColumnName property can set one of the columns to have a link in it
 'objSortableTable.LinkedColumnName = "lastname"
  
  'the Link property is the link to be used on the column that is the LinkedColumnName 
'	objSortableTable.Link = "myDetailPage.asp" 

 'the FieldToAppendToLink property is the field whos value you want passed with the link
'objSortableTable.FieldToAppendToLink = "employeeid" 'EX: &employeeid=5 will be appended

   

'THIRD: call the writeTable method where you want your table to be written 
'	objSortableTable.writeTable

'SAMPLE CALL:

	
	'dim objST 'for sortable table
	'dim objConn, strSQL
	'Set objConn = server.CreateObject("ADODB.Connection")
	'objConn.Open "northwind","",""
	'strSQL = "select employeeID, firstName, lastName, 'Northwind Traders' as comp_name, notes as comments from employees" '"Select * From workouts"	

	'set objST = new clsSortableTable
		
	'objST.Connection = objConn  
	'objST.SQL = strSQL
	'objST.DefaultSort = "lastName"	  
	'objST.DisplayedFields = array("lastname", "firstname", "comp_name", "comments")  
	'objST.PageSize = 3  
	'objST.TableAttributes = "border=1 cellpadding=1 cellspacing=0"
	'objST.HeaderRowAttributes = "bgcolor='#cccccc'"
	'objST.HeaderCellAttributes = "align=center valign=bottom"
	'objST.RowAttributes = "bgcolor=white"	
	'objST.TextLink = false
	'objST.LinkedColumnName = "lastname"
	'objST.FieldToAppendToLink = "employeeid"
	'objST.Link = "myDetailPage.asp"
	'objST.WriteTable
'
'programmer: Devin Garlit  dgarlit@hotmail.com  5/03/01
''''''''''''''''''''''''''''''''''''''''''''''''''''


	Class clsSortableTable

		'class level private variable
		Private m_objConn
		Private m_strSQL
		Private m_strDefaultSort
		Private m_strSort
		Private m_arrDisplayedFields
		Private m_intCurrentPage
		Private m_strPageName
		Private m_intPageSize
		Private m_strLinkedColumnName
		Private m_strLink
		Private m_strFieldToAppendToLink
		Private m_strTableAttributes
		Private m_strHeaderRowAttributes
		Private m_strHeaderCellAttributes
		Private m_strRowAttributes
		Private m_strCellAttributes
		Private m_strImagePath
		Private m_arrPlainImages
		Private m_arrAscImages
		Private m_arrDscImages
		Private m_arrNavString
		Private m_strFontTagAttributes
		Private m_blnTextLink
		Private m_CellStyle
		Private RS 'recordset		
		Private strTemp, field, strMoveFirst, strMoveNext, strMovePrevious, strMoveLast
		Private i, j,intTotalPages, intCurrentRecord, intTotalRecords, intTotalFields, intDisplayedFields

		'properties
		Public Property Get DefaultSort
			DefaultSort = m_strDefaultSort
		End Property

		Public Property Let DefaultSort(strDefaultSort)
			m_strDefaultSort = strDefaultSort
		End Property

		Public Property Get Sort
			Sort = m_strSort
		End Property

		Public Property Let Sort(strSort)
			m_strSort = strSort
		End Property

		Public Property Get Connection
			Connection = m_objConn
		End Property

		Public Property Let Connection(objConn)
			m_objConn = objConn
		End Property

		Public Property Get SQL
			SQL = m_strSQL
		End Property

		Public Property Let SQL(strSQL)
			m_strSQL = strSQL
		End Property

		Public Property Get CurrentPage
			CurrentPage = m_intCurrentPage
		End Property

		Public Property Let CurrentPage(intCurrentPage)
			m_intCurrentPage = intCurrentPage
		End Property

		Public Property Get DisplayedFields
			DisplayedFields = m_arrDisplayedFields
		End Property

		Public Property Let DisplayedFields(arrDisplayedFields)
			m_arrDisplayedFields = arrDisplayedFields
		End Property

		Public Property Get PageName
			PageName = m_strPageName
		End Property

		Public Property Let PageName(strPageName)
			m_strPageName = strPageName
		End Property

		Public Property Get PageSize
			PageSize = m_intPageSize
		End Property

		Public Property Let PageSize(intPageSize)
			m_intPageSize = intPageSize
		End Property
				
		Public Property Get FieldToAppendToLink
			FieldToAppendToLink = m_strFieldToAppendToLink
		End Property

		Public Property Let FieldToAppendToLink(strFieldToAppendToLink)
			m_strFieldToAppendToLink = strFieldToAppendToLink
		End Property
		
		Public Property Get LinkedColumnName
			LinkedColumnName = m_strLinkedColumnName
		End Property

		Public Property Let LinkedColumnName(strLinkedColumnName)
			m_strLinkedColumnName = strLinkedColumnName
		End Property

		Public Property Get Link
			Link = m_strLink
		End Property

		Public Property Let Link(strLink)
			m_strLink= strLink
		End Property
		
		Public Property Get TableAttributes
			TableAttributes = m_strTableAttributes
		End Property

		Public Property Let TableAttributes(strTableAttributes)
			m_strTableAttributes = strTableAttributes
		End Property
		
		Public Property Get HeaderRowAttributes
			HeaderRowAttributes = m_strHeaderRowAttributes
		End Property

		Public Property Let HeaderRowAttributes(strHeaderRowAttributes)
			m_strHeaderRowAttributes = strHeaderRowAttributes
		End Property
		
		Public Property Get HeaderCellAttributes
			HeaderCellAttributes = m_strHeaderCellAttributes
		End Property

		Public Property Let HeaderCellAttributes(strHeaderCellAttributes)
			m_strHeaderCellAttributes = strHeaderCellAttributes
		End Property
		
		Public Property Get CellAttributes
			CellAttributes = m_strCellAttributes
		End Property

		Public Property Let CellAttributes(strCellAttributes)
			m_strCellAttributes = strCellAttributes
		End Property
		
		Public Property Get RowAttributes
			RowAttributes = m_strRowAttributes
		End Property

		Public Property Let RowAttributes(strRowAttributes)
			m_strRowAttributes = strRowAttributes
		End Property
		
		Public Property Get ImagePath
			ImagePath = m_strImagePath
		End Property

		Public Property Let ImagePath(strImagePath)
			m_strImagePath = strImagePath
		End Property
		
		Public Property Get PlainImages
			PlainImages = m_arrPlainImages
		End Property

		Public Property Let PlainImages(arrPlainImages)
			m_arrPlainImages = arrPlainImages
		End Property
		
		Public Property Get DscImages
			DscImages = m_arrDscImages
		End Property

		Public Property Let DscImages(arrDscImages)
			m_arrDscImages = arrDscImages
		End Property		
		
		Public Property Get AscImages
			AscImages = m_arrAscImages
		End Property

		Public Property Let AscImages(arrAscImages)
			m_arrAscImages = arrAscImages
		End Property
		
		Public Property Get NavString
			NavString = m_strNavString
		End Property

		Public Property Let NavString(strNavString)
			m_strNavString = strNavString
		End Property
		
		Public Property Get FontTagAttributes
			FontTagAttributes = m_strFontTagAttributes
		End Property

		Public Property Let FontTagAttributes(strFontTagAttributes)
			m_strFontTagAttributes = strFontTagAttributes
		End Property

		Public Property Get TextLink
			TextLink = m_blnTextLink
		End Property

		Public Property Let TextLink(blnTextLink)
			m_blnTextLink = blnTextLink
		End Property		
		
		Public Property Get CellStyle
			CellStyle = m_CellStyle
		End Property

		Public Property Let CellStyle(strCellStyle)
			m_CellStyle = strCellStyle
		End Property
	
		Private Sub Class_Initialize()
			m_strDefaultSort = ""
			m_strSort = request("sort")
			m_arrDisplayedFields = ""
			m_intCurrentPage = request("page")
			m_strPageName = Request.ServerVariables("SCRIPT_NAME")
			m_intPageSize = 10
			m_strLinkedColumnName	= ""
			m_strLink = ""
			m_strFieldToAppendToLink = ""
			m_strTableAttributes = ""
			m_strHeaderRowAttributes = ""
			m_strHeaderCellAttributes = ""
			m_strRowAttributes = ""
			m_strCellAttributes = ""
			m_strImagePath = "../img/"
			m_arrPlainImages = ""
			m_arrAscImages = ""
			m_arrDscImages = ""
			m_arrNavString = ""
			m_strFontTagAttributes = ""
			m_blnTextLink = true
			m_cellStyle = ""					
			i = 0
			j = 0		

			set RS = server.CreateObject("adodb.recordset")
			
		End Sub
		
				

		Public sub WriteTable()			

				'if no sort then use default sort
				if m_strSort = "" then					
						m_strSort = m_strDefaultSort					
				end if

				'if no page given use page 1
				if m_intCurrentPage = "" then
					m_intCurrentPage = 1
				end if

        'open recordset
 				createRecordset

				'if no array of fields is given then use all the fields
				if isArray(m_arrDisplayedFields) then
					intDisplayedFields = ubound(m_arrDisplayedFields)
				else
					intDisplayedFields = intTotalFields - 1
					redim m_arrDisplayedFields(intDisplayedFields)
					for i = 0 to intDisplayedFields
						m_arrDisplayedFields(i) = RS.Fields(i).Name
					next
				end if
		
			  'fill image arrays
				setUpImageArrays
				
			

				Response.Write "<table " & m_strTableAttributes & " >" & vbcrlf
				'display the headers
				writeHeaders()
				'end table header

			if RS.EOF then
				Response.Write "<tr " & m_strRowAttributes & " >" & vbcrlf
				Response.Write "<td " & m_strCellAttributes & " Colspan='100%' align=center >" & vbcrlf
				Response.Write "There are no records."	
				Response.Write "</td>" & vbcrlf
				Response.Write "</tr>" & vbcrlf
			else
				'start records
				for i = intCurrentRecord to RS.PageSize
					if not RS.EOF then
						Response.Write "<tr " & m_strRowAttributes & " >" & vbcrlf

						'start record cells
						for j = 0 to intDisplayedFields
'''''''''''''''''''''''''''''''''																						
					    if lcase(m_arrDisplayedFields(j)) = lcase(m_strLinkedColumnName) then
								Response.write "<td " & m_strCellAttributes & " ><font " & m_strFontTagAttributes &  " >&nbsp;<a href='" & m_strLink & "?page="&m_intCurrentPage & _
															 "&sort=" & m_strSort &  "&" & m_strFieldToAppendToLink &  "=" & RS(m_strFieldToAppendToLink) & "' >" & RS(m_arrDisplayedFields(j)) & "</font></td>"	
							else	
								Response.write "<td " & m_strCellAttributes & " ><font " & m_strFontTagAttributes &  " >&nbsp;" & RS(m_arrDisplayedFields(j)) & "</font></td>"	
							end if	
''''''''''''''''''''''''''''''''								
						next
					 'end cells						
					Response.Write "</tr>" & vbcrlf					
					RS.MoveNext
					end if	
				next
			end if	
				'end records
				'end table
				Response.Write "</table>" & vbcrlf
				
				'write the navigation buttons
				writeNavButtons
    		
		End Sub
		
		
	 private sub Class_Terminate()	
			if RS.State = &H00000001 then 'its open
						RS.Close
			end if
			set RS = nothing
	 end sub	
	 
	'''CREATERECORDSET
	private sub createRecordset()
			with RS
				.CursorLocation=3
				.Open m_strSQL & " order by " & replace(m_strSort,"desc"," desc"), m_objConn,3
				if not rs.EOF then
					.PageSize = cint(m_intPageSize)
					intTotalPages = .PageCount
					intCurrentRecord = .AbsolutePosition
					.AbsolutePage = m_intCurrentPage
					intTotalRecords = .RecordCount
					intTotalFields = .Fields.Count
				else
					intTotalFields = 0
					intTotalRecords = 0
					intTotalPages = 0
					intCurrentRecord = 0
				end if
			end with
	end sub	
		
		
	'''WRITEHEADERS
		private sub writeHeaders()
			Response.Write "<tr " & m_strHeaderRowAttributes & ">" & vbcrlf
			for i = 0 to intDisplayedFields
				Response.Write "<td " & m_strHeaderCellAttributes & " >" & vbcrlf
				if instr(lcase(m_strSort),lcase(m_arrDisplayedFields(i))) then 'if we come across the current sort in the loop
					if instr(m_strSort, "desc")then 'if the sort is descending, write the link for ascending, but give a descending image
						Response.Write "<a href=" & m_strPageName & "?sort="& m_arrDisplayedFields(i)  &"&page="&m_intCurrentPage&">" & m_arrDscImages(i) & "</a>" & vbcrlf
					else
						Response.Write "<a href=" & m_strPageName & "?sort="& m_arrDisplayedFields(i)  &"desc&page="&m_intCurrentPage&">" & m_arrAscImages(i)  & "</a>"	& vbcrlf
					end if
				else 'if not the current sort
					Response.Write "<a href=" & m_strPageName & "?sort="& m_arrDisplayedFields(i)  &"&page="&m_intCurrentPage&">" &  m_arrPlainImages(i)  & "</a>"	& vbcrlf
				end if
				Response.Write "</td>"	& vbcrlf
			next
			Response.Write "</tr>"
		end sub			
		
		'''WRITENAVBUTTONS
		private sub writeNavButtons()
			if intTotalPages = 0 then intTotalPages = 1			
				select case cint(m_intCurrentPage)
					case cint(intTotalPages) 'last page
						if cint(intTotalPages) = 1 then 'its the only page
							strMoveFirst = m_arrNavString(0) '"<a href=" & m_strPageName & "?sort="& m_strSort &"&page=1 >"
							strMovePrevious = m_arrNavString(1) '"<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & m_intCurrentPage - 1 &  " >"
							strMoveNext = m_arrNavString(2)
							strMoveLast = m_arrNavString(3) '"<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & intTotalPages &  " >"
						else
							strMoveFirst = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=1 >"& m_arrNavString(0) &"</a>"
							strMovePrevious = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & m_intCurrentPage - 1 &  " >"& m_arrNavString(1) &"</a>"
							strMoveNext = m_arrNavString(2)
							strMoveLast = m_arrNavString(3) '"<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & intTotalPages &  " >"
						end if
					case 1 'first page
						strMoveFirst = m_arrNavString(0) '"<a href=" & m_strPageName & "?sort="& m_strSort &"&page=1 >"
						strMovePrevious = m_arrNavString(1) '"<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & m_intCurrentPage - 1 &  " >"
						strMoveNext = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & m_intCurrentPage + 1 &  " >"& m_arrNavString(2) &"</a>"
						strMoveLast = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & intTotalPages &  " >"& m_arrNavString(3) &"</a>"
					case else
						strMoveFirst = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=1 >"& m_arrNavString(0) &"</a>"
						strMovePrevious = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & m_intCurrentPage - 1 &  " >"& m_arrNavString(1) &"</a>"
						strMoveNext = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & m_intCurrentPage + 1 &  " >"& m_arrNavString(2) &"</a>"
						strMoveLast = "<a href=" & m_strPageName & "?sort="& m_strSort &"&page=" & intTotalPages &  " >"& m_arrNavString(3) &"</a>"
				end select
						

				with Response
					.Write strMoveFirst
					.Write strMovePrevious
					.Write " " & m_intCurrentPage & " of " & intTotalPages & " "
					.Write strMoveNext
					.Write strMoveLast
				end with
		 end sub		
		 
		 '''SETUPIMAGEARRAYS
		 'If the image arrays are empty Fill them with .gif images based on fieldnames
		 private sub setUpImageArrays()			
				if not isArray(m_arrPlainImages)or isnull(m_arrPlainImages) then
					redim m_arrPlainImages( intDisplayedFields )
					for i = 0 to intDisplayedFields
						if cbool(m_blnTextLink) then
							m_arrPlainImages(i) = m_arrDisplayedFields(i) 
						else
							m_arrPlainImages(i) = "<img src='" & m_strImagePath & m_arrDisplayedFields(i) & ".gif" & "' border=0>"
						end if	
					next
				end if
				if not isArray(m_arrAscImages) or isnull(m_arrAscImages) then
					redim m_arrAscImages( intDisplayedFields )
					for i = 0 to intDisplayedFields
						if cbool(m_blnTextLink) then
							m_arrAscImages(i) = m_arrDisplayedFields(i)
						else
							m_arrAscImages(i) = "<img src='" & m_strImagePath & m_arrDisplayedFields(i) & "_asc.gif"  & "' border=0>"
						end if	
					next
				end if
				if not isArray(m_arrDscImages) or isnull(m_arrDscImages) then
					redim m_arrDscImages( intDisplayedFields )
					for i = 0 to intDisplayedFields
						if cbool(m_blnTextLink) then
							m_arrDscImages(i) = m_arrDisplayedFields(i)
						else
							m_arrDscImages(i) = "<img src='" & m_strImagePath & m_arrDisplayedFields(i) & "_desc.gif"  & "' border=0>"
						end if		
					next
				end if
				if not isArray(m_arrNavString) or isNull(m_arrNavString) then
					redim m_arrNavString(3)
					m_arrNavString(0) = "<img src='" & m_strImagePath & "first.gif' border=0>"
					m_arrNavString(1) = "<img src='" & m_strImagePath & "prev.gif' border=0>"
					m_arrNavString(2) = "<img src='" & m_strImagePath & "next.gif' border=0>"
					m_arrNavString(3) = "<img src='" & m_strImagePath & "last.gif' border=0>"
				end if
		 end sub
		
		
	End Class
	
%>