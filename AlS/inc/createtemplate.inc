<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

'**************************************************
' Copyright (C) by Atlas Industries Limited
' E-mail: info@atlasindustries.com
'
' NAME:
'
'	Function ReadFromTemplate() 
'
' DESCRIPTION:
'
'	This function allows us to read the page HTML
'	from the template and store it as a string. We
'	then adjust the path of images (so that our
'	pages on the root using the template link to
'	the correct path of the template's image
'	folder. Next we replace the special @@ tags
'	in the template with the content. Finally, we
'	split the whole page string into a two-box
'	array, with the split made around the
'	@@maincontent tag. The two sections of
'	HTML will later be added above and below the
'	main content on each front end page.
'
' AUTHOR:
' DATE:
'
' NOTE:
'**************************************************

Function ReadFromTemplate(ByVal strTitle, ByVal strFunction, arrPageTemplate, ByVal strTemplatePath)

	Dim objFile, objTStream, strPathFile, strPageBaseText, strTemplateLocation

'--------------------------------------------------
' If strTemplatePath is empty we look to read the template location from a cookie.
'--------------------------------------------------
If strTemplatePath = "" then
	strTemplateLocation = Request.Cookies("templatepath")

	'--------------------------------------------------
	' If the cookie doesn't exist then we set the template to our default one.
	'--------------------------------------------------

	If Request.Cookies("templatepath") = "" Then
		strTemplateLocation = "templates/template1/"
		Response.Cookies("templatepath") = ""
	End if
Else
	strTemplateLocation = strTemplatePath
End if

'--------------------------------------------------
' Loop through header template file content
'--------------------------------------------------

	Set objFile		= Server.CreateObject("Scripting.FileSystemObject")
	strPathFile		= Server.MapPath(strTemplateLocation & "main/ats_header.htm")
	Set objTStream	= objFile.OpenTextFile (strPathFile, 1, False, False)
	
	While Not objTStream.AtEndOfStream
		strPageBaseText = strPageBaseText & objTStream.ReadLine & vbcrlf
	Wend
	
	Set objTStream = Nothing

'--------------------------------------------------
' Loop through footer template file content
'--------------------------------------------------

	strPathFile		= Server.MapPath(strTemplateLocation & "main/ats_footer.htm")
	
	Set objTStream	= objFile.OpenTextFile (strPathFile, 1, False, False)
	
	While Not objTStream.AtEndOfStream
		strPageBaseText = strPageBaseText & objTStream.ReadLine & vbcrlf
	Wend
	
	Set objTStream = Nothing

'--------------------------------------------------
' Adjust hyperlinks on template to correct path
'--------------------------------------------------

	strPageBaseText = Replace(strPageBaseText,"../","")

'--------------------------------------------------
' Fix path to template images
'--------------------------------------------------

	strPageBaseText = Replace(strPageBaseText,"images/",strTemplateLocation & "images/")

'--------------------------------------------------
' Replace @@title tag on the template page with the string we get
'--------------------------------------------------
	strPageBaseText = Replace(strPageBaseText,"@@title", strTitle)

'--------------------------------------------------
' Replace @@function tag on the template page with the string we get
'--------------------------------------------------
	
	strPageBaseText = Replace(strPageBaseText,"@@function", strFunction)
	
'--------------------------------------------------
' Create a two-box array by splitting the page HTML around the @@maincontent tag
'--------------------------------------------------
	
	arrPageTemplate = split(strPageBaseText ,"@@maincontent",-1)

	ReadFromTemplate = arrPageTemplate

END Function

'**************************************************
' Copyright (C) by Atlas Industries Limited
' E-mail: info@atlasindustries.com
'
' NAME:
'
'	Function ReadFromTemplateAll()
'
' DESCRIPTION:
'
'	This function allows us to read the specified page HTML
'	from the template and store it as a string (header and footer is default).
'	In case the variable point to the specified page HTML is empty, 
'	it means only read header and footer.
'	We then adjust the path of images (so that our
'	pages on the root using the template link to
'	the correct path of the template's image
'	folder. Next we replace the special @@ tags
'	in the template with the content. Finally, we
'	split the whole page string into a three-box
'	array, with the split made around the
'	@@maincontent tag. The three sections of
'	HTML will later be added above and below the
'	main content on each front end page.
'
' AUTHOR:
' DATE:
'
' NOTE:
'**************************************************

Function ReadFromTemplateAll(arrPageTemplate, ByVal strTemplatePath, ByVal strTemplateName)

	Dim objFile, objTStream, strPathFile, strPageBaseText, strTemplateLocation

'--------------------------------------------------
' If strTemplatePath is empty we look to read the template location from a cookie.
'--------------------------------------------------
If strTemplatePath = "" then
	strTemplateLocation = Request.Cookies("templatepath")

	'--------------------------------------------------
	' If the cookie doesn't exist then we set the template to our default one.
	'--------------------------------------------------

	If Request.Cookies("templatepath") = "" Then
		strTemplateLocation = "templates/template1/"
		Response.Cookies("templatepath") = ""
	End if
Else
	strTemplateLocation = strTemplatePath
End if

'--------------------------------------------------
' Loop through header template file content
'--------------------------------------------------

	Set objFile		= Server.CreateObject("Scripting.FileSystemObject")
	strPathFile		= Server.MapPath(strTemplateLocation & "main/ats_header.htm")
	
	Set objTStream	= objFile.OpenTextFile (strPathFile, 1, False, False)
	
	While Not objTStream.AtEndOfStream
		strPageBaseText = strPageBaseText & objTStream.ReadLine & vbcrlf
	Wend
	
	Set objTStream = Nothing

'--------------------------------------------------
' Loop through body template file content
'--------------------------------------------------

	strPathFile		= Server.MapPath(strTemplateLocation & "main/" & strTemplateName)
	
	Set objTStream	= objFile.OpenTextFile (strPathFile, 1, False, False)
	
	While Not objTStream.AtEndOfStream
		strPageBaseText = strPageBaseText & objTStream.ReadLine & vbcrlf
	Wend
	
	Set objTStream = Nothing
		
'--------------------------------------------------
' Loop through footer template file content
'--------------------------------------------------

	strPathFile		= Server.MapPath(strTemplateLocation & "main/ats_footer.htm")
	
	Set objTStream	= objFile.OpenTextFile (strPathFile, 1, False, False)
	
	strPageBaseText = strPageBaseText & "@@maincontent"
	
	While Not objTStream.AtEndOfStream
		strPageBaseText = strPageBaseText & objTStream.ReadLine & vbcrlf
	Wend
	
	Set objTStream = Nothing


'--------------------------------------------------
' Adjust hyperlinks on template to correct path
'--------------------------------------------------

	strPageBaseText = Replace(strPageBaseText, "../", "")

'--------------------------------------------------
' Fix path to template images
'--------------------------------------------------

	strPageBaseText = Replace(strPageBaseText,"images/",strTemplateLocation & "images/")
		
'--------------------------------------------------
' Create a three-box array by splitting the page HTML around the @@maincontent tag
'--------------------------------------------------
	
	arrPageTemplate = split(strPageBaseText ,"@@maincontent",-1)

	ReadFromTemplateAll = arrPageTemplate

END Function


</SCRIPT>