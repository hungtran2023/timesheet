<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

Sub GetConfigObj(objReturn, arrQuery, strFileName)
response.write(strFileName)
response.end

	Dim index
	Dim objFS, txtFile
	Dim arrName, arrParamPair
	Dim strInput
	objReturn = ""
	arrName = Split(arrQuery, ",")
	For index = 0 To UBound(arrName)
		arrName(index) = Trim(arrName(index))		
	Next
	' Create a FileSystemObject instance
	Set objFS = Server.CreateObject("Scripting.FileSystemObject")
	' Open configuration file
'	Set txtFile = objFS.OpenTextFile(Server.MapPath(strFileName))
	While txtFile.AtEndOfStream = False
		strInput = txtFile.ReadLine
		If strInput <> "" Then
			arrParamPair = Split(strInput,"=")
			For index = 0 To UBound(arrParamPair)
				arrParamPair(index) = Trim(arrParamPair(index))
			Next
			For index = 0 To UBound(arrName)
				If LCase(arrName(index)) = LCase(arrParamPair(0)) Then
					objReturn = objReturn & arrParamPair(0) & Chr(149) & arrParamPair(1) & Chr(161)
					index = UBound(arrName)
				End If
			Next								
		End If		
		
	Wend
	txtFile.Close
	If Right(objReturn,1)=Chr(161) Then objReturn = Left(objReturn, Len(objReturn)-1)
End Sub

Function Config(strName, objConfig) 
	Dim arrPairs, arrPair
	Dim index
	arrPairs = Split(objConfig, Chr(161))
	For index = 0 To UBound(arrPairs)
		arrPair = Split(arrPairs(index), Chr(149))
		If LCase(arrPair(0)) = LCase(strName) Then
			Config = arrPair(1)
		End If
	Next
End Function

</SCRIPT>
