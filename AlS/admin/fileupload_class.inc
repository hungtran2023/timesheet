<%
Class FileUpload

    ' ###############################
    ' # The ASP Emporium            #
    ' # http://www.aspemporium.com/ #
    ' # BILLYBOY's FILEUPLOAD CLASS #
    ' # Version 1.2                 #
    ' #                             #################################################
    ' #                                                                             #
    ' # PURPOSE:                                                                    #
    ' #    The FileUpload Object allows users to upload files to the server.        #
    ' #    FileUpload lets the developer control the type and size of the upload as #
    ' #    well as the upload directory and the overwrite status for already        #
    ' #    existing uploads. Currently exposes 15 properties and methods to allow   #
    ' #    an ASP programmer to manipulate and control multiple file uploads to     #
    ' #    their server without using FTP.                                          #
    ' #                                                                             #
    ' #                                                                             #
    ' # UPGRADES:                                                                   #
    ' #    The latest and greatest release version of the FileUpload Class can be   #
    ' #    found here:                                                              #
    ' #      http://www.aspEmporium.com/aspEmporium/examples/fileupload_class.asp   #
    ' #                                                                             #
    ' #    Joining the ASP Emporium mailing list is the best way to find out when   #
    ' #    the FileUpload class is updated officially as you'll get a nifty email   #
    ' #    reminder in your inbox when you get to work that day (ad free, baby!)    #
    ' #      http://www.aspEmporium.com/aspEmporium/mailingList.asp                 #
    ' #                                                                             #
    ' #                                                                             #
    ' # VERSION HISTORY:                                                            #
    ' #    1.2  [February 23, 2001]                                                 #
    ' #        - Abandoned web site documentation for this class and put it all in  #
    ' #          this file instead. At least now the documentation goes wherever    #
    ' #          this class does.                                                   #
    ' #        - Added a bunch of new properties and methods based on user feedback #
    ' #        - Supports multiple file downloads at the same time. Supports        #
    ' #          looping when gathering and saving multiple file uploads if         #
    ' #          necessary.                                                         #
    ' #        - Supports gathering of other form inputs as well as file inputs     #
    ' #           at the same time.                                                 #
    ' #        - Added minimum upload size requirement                              #
    ' #        - Added properties to prevent uploading (acts like a stop            #
    ' #          command - if you want to disallow uploading temporarily while      #
    ' #          you are admin-ing the site or something) as well as to make        #
    ' #          uploading optional (for example a used car system that had a user  #
    ' #          fill out a form and if desired, optionally attach a picture of the #
    ' #          car...) Both of these properties work best if there are other      #
    ' #          form inputs along with the file input.                             #
    ' #        - Added read only properties that can be used to determine where     #
    ' #          files were uploaded to on the server as well as to determine if    #
    ' #          an upload was successfully completed. These properties could be    #
    ' #          used to log uploads if necessary.                                  #
    ' #        - Miscellaneous improvements were made class-wide but basically      #
    ' #          version 1.2 is a serious hack (I meant "extension") of version 1.1 #
    ' #          and just goes to show how much you can re-work something and       #
    ' #          almost get a whole new program out of it. For this class I doubled #
    ' #          the number of properties and methods available in just under a     #
    ' #          couple of hours. Thanks to all who wrote in the forum requesting   #
    ' #          new functionality. I think I did em all and included the lot in    #
    ' #          this release.                                                      #
    ' #        - The built in upload form is a little more interactive based        #
    ' #          on the entered property settings.                                  #
    ' #                                                                             #
    ' #    1.1  [July 25, 2000]                                                     #
    ' #        Initial Release                                                      #
    ' #                                                                             #
    ' #                                                                             #
    ' # WARNINGS:                                                                   #
    ' #    Allowing users to upload files can pose a significant security risk      #
    ' #    unless strictly controlled...                                            #
    ' #    Please recognize the potential danger that this system can cause.        #
    ' #    Please review the list below for warnings.                               #
    ' #                                                                             #
    ' #    1.) From a security standpoint, giving users the power to upload         #
    ' #        executable files or scripts can give them the same power as a        #
    ' #        developer on that website, allowing remote users to upload           #
    ' #        potentially hazardous files and then execute the code using a web    #
    ' #        browser. With that in mind, the following file extensions should     #
    ' #        never be allowed to be uploaded: (This is by no means a complete     #
    ' #        list...)                                                             #
    ' #+-----------------+--------------------------------+------------------------+#
    ' #| Executables and | Scripts, especially those that | Files that may contain |#
    ' #| system files:   | can be executed by a browser:  | harmful macros:        |#
    ' #+-----------------+--------------------------------+------------------------+#
    ' #|      *.exe      |              *.asp             |          *.doc         |#
    ' #|      *.dll      |              *.shtml           |          *.xls         |#
    ' #|      *.bat      |              *.js              |          *.mdb         |#
    ' #|      *.com      |              *.vbs             |                        |#
    ' #|      *.dos      |              *.cfm             |                        |#
    ' #|      *.dat      |              *.cgi             |                        |#
    ' #|      *.class    |              *.pl              |                        |#
    ' #|      *.sys      |              *.jsp             |                        |#
    ' #|      *.ini      |              *.wsc             |                        |#
    ' #|      *.cab      |              *.sql             |                        |#
    ' #+-----------------+--------------------------------+------------------------+#
    ' #                                                                             #
    ' #    2.) This system is slow. Therefore allowable sizes of more that          #
    ' #        1024000 b (1 mb) can result in a server.scripttimeout error. File    #
    ' #        transfers of this size should be sent using the File Transfer        #
    ' #        Protocol (FTP) and not HTTP. Use caution when setting the maximum    #
    ' #        allowable file size.                                                 #
    ' #                                                                             #
    ' #    3.) If possible, set the upload directory to a path that is not          #
    ' #        available from the public server. This way you will have some        #
    ' #        protection against executable and script executions as those files   #
    ' #        cannot be reached via the web (upload directory is not in the shared #
    ' #        server space).                                                       #
    ' #                                                                             #
    ' #    4.) Only enable the File overwrite if you don't care about the data in   #
    ' #        the upload directory.                                                #
    ' #                                                                             #
    ' #                                                                             #
    ' # CREDITS:                                                                    #
    ' #    The last three functions in this class                                   #
    ' #    (getByteString, getString, BuildUploadRequest)                           #
    ' #    I got from:                                                              #
    ' #                                                                             #
    ' #    Philippe Collignon's article on ASP Today:                               #
    ' #    http://www.asptoday.com/articles/20000316.htm                            #
    ' #	                                                                            #
    ' #    It's a nice bit of reusable code. All I really did was	            #
    ' #    further encapsulate it into a class with some more options ...           #
    ' #                                                                             #
    ' #    The actual code that performs the parse and upload of the                #
    ' #    chosen file is all from Philippe Collignon's article.	            #
    ' #                                                                             #
    ' #                                                                             #
    ' # OTHER CREDITS:                                                              #
    ' #    Anyone who complained about the first release or suggested ideas         #
    ' #    for this newest release. Without you people, I would have no motivation  #
    ' #    to release new versions.                                                 #
    ' #                                                                             #
    ' #                                                                             #
    ' # PROPERTIES:                                                                 #
    ' #    The class has the following properties, all of which are optional:       #
    ' #                                                                             #
    ' #    Extensions                                                               #
    ' #       Expects an array of allowable file extensions.                        #
    ' #       All file extensions should begin with "." and should                  #
    ' #       be four characters including the ".".                                 #
    ' #       NOTE:                                                                 #
    ' #       If you want to allow all files (not recommended) use an asterisk      #
    ' #       instead of an array as input. This is the only instance in which      #
    ' #       the Extensions property can be other than an array.                   #
    ' #       [default: ".txt", ".htm", ".html", ".zip", ".inc"]                    #
    ' #                                                                             #
    ' #    UploadDirectory                                                          #
    ' #       string representing the absolute path to the upload directory         #
    ' #       [default: / (root folder of web site)]                                #
    ' #                                                                             #
    ' #    AllowOverWrite                                                           #
    ' #       Expects a boolean indicating whether or not already existing          #
    ' #       uploads should be replaced by new uploads of the same name.           #
    ' #       boolean - True overwrites, False doesn't                              #
    ' #       [default: False]                                                      #
    ' #                                                                             #
    ' #    UploadSize                                                               #
    ' #       double representing the maximum file size allowed ( in bytes )        #
    ' #       [default: 100,000 (97.7 k)]                                           #
    ' #                                                                             #
    ' #    MininumRequiredByteCountForUpload                                        #
    ' #       long representing the minimum file size allowed (in bytes)            #
    ' #       [default: 1,024 (1.0 k)]                                              #
    ' #                                                                             #
    ' #    AllowUserToBypassUpload                                                  #
    ' #       boolean indicating whether or not to force the user to upload.        #
    ' #       True allows a user to not upload a file but still allows other form   #
    ' #       inputs to be read. False forces an upload to occur.                   #
    ' #       [default: False]                                                      #
    ' #                                                                             #
    ' #    BypassUploadIfMinimumByteCountThresholdNotMet                            #
    ' #       boolean indicating whether or not to allow a user to not upload if    #
    ' #       it appears that they did not choose anything to upload. In other      #
    ' #       words, if you want the file upload part to be optional                #
    ' #       [default: True]                                                       #
    ' #                                                                             #
    ' #    Version                                                                  #
    ' #       READ ONLY - string representing the major.minor version number        #
    ' #       of this class                                                         #
    ' #                                                                             #
    ' #    UploadSuccessful                                                         #
    ' #       READ ONLY - boolean indicating whether or not the upload was          #
    ' #       successful.                                                           #
    ' #                                                                             #
    ' #    AbsolutePathToUploadedFile                                               #
    ' #       READ ONLY - contains a string with the fully qualified path to the    #
    ' #       uploaded file. Only contains valid data if the file was uploaded      #
    ' #       successfully.                                                         #
    ' #                                                                             #
    ' #                                                                             #
    ' # METHODS:                                                                    #
    ' #    The class has only one method which must be called to perform an upload. #
    ' #    After the upload method is called, a second method called                #
    ' #    FormInput(input) can be called to read other form inputs that may have   #
    ' #    also been passed along with the blob using the multipart/form enctype.   #
    ' #                                                                             #
    ' #    NoUploadMode()                                                           #
    ' #       prevents uploads. This is the equivalent of setting the               #
    ' #       AllowUserToBypassUpload property to true. Must be called before       #
    ' #       the upload() method.                                                  #
    ' #                                                                             #
    ' #    OptionalUploadMode()                                                     #
    ' #       Allows the upload to be optional. Equivalent of setting the           #
    ' #       BypassUploadIfMinimumByteCountThresholdNotMet property to True.       #
    ' #       Must be called before the Upload() method.                            #
    ' #                                                                             #
    ' #    Upload()                                                                 #
    ' #       uploads a file or displays the update form (automatically determined) #
    ' #       The upload method is required and must be called.                     #
    ' #                                                                             #
    ' #    FormInput(inputname)                                                     #
    ' #       Returns a string representing the value of a form input.              #
    ' #       This function behaves similar to Request.Form(inputname)              #
    ' #       and is used in it's place when running the fileupload class           #
    ' #       (microsoft's rule, not mine...). This does not output a collection    #
    ' #       only a string representing the value of inputname.                    #
    ' #       - example: instead of Request.Form("info") use                        #
    ' #           object.FormInput("info") to get the same information but          #
    ' #           without the error!                                                #
    ' #       Must be called after the Upload() method.                             #
    ' #                                                                             #
    ' #    SaveAnotherBlob(inputname)                                               #
    ' #       Returns void. Saves another binary file, other than the input         #
    ' #       named "blob". Use this method to save more than one file at a time.   #
    ' #       Certain rules apply:                                                  #
    ' #        1.) inputname must the the form input name to be saving and must     #
    ' #            be of type "file".                                               #
    ' #        2.) the combined byte count of all files to be uploaded at the       #
    ' #            same time must fall under the UploadSize property setting.       #
    ' #        3.) the AbsolutePathToUploadedFile property will hold the path       #
    ' #            to the most recently uploaded file until the upload method       #
    ' #            is called again.                                                 #
    ' #        4.) the byte count check will have already been performed before     #
    ' #            you can call this method. This check will not be repeated when   #
    ' #            you call this method.                                            #
    ' #        5.) all other rules apply to this upload as well. Rules such as      #
    ' #            allowable file types and any other property settings are applied #
    ' #            to this upload.                                                  #
    ' #        6.) if you are using the same file to display the form and to        #
    ' #            handle the upload, you must check the UploadSuccessful property  #
    ' #            to make sure that the first file uploaded successfully otherwise #
    ' #            you will get an "ObjectRequired" error.                          #
    ' #        7.) The first file must not fail during upload. If it does, you      #
    ' #            will not be able to call this method successfully.               #
    ' #        8.) Must be called after the Upload() method.                        #
    ' #                                                                             #
    ' #                                                                             #
    ' # SYNTAX:                                                                     #
    ' #    Set variable = New FileUpload                                            #
    ' #                                                                             #
    ' #                                                                             #
    ' # SUGGESTING CHANGES, GETTING HELP, POSTING BUGS/UPDATES/WORKAROUNDS:         #
    ' #    A thread has developed in the ASP Developer's forum about the            #
    ' #    FileUpload class.  If you are seeking more information about this class  #
    ' #    or are experiencing difficulty using the class, please post your         #
    ' #    questions there:                                                         #
    ' #    http://www.aspEmporium.com/aspEmporium/forum/display_forum.asp?fid=10    #
    ' #                                                                             #
    ' #    You can also search for posts relating to the FileUpload Class:          #
    ' #  http://www.aspEmporium.com/aspEmporium/forum/search.asp?keyword=fileupload #
    ' #                                                                             #
    ' #                                                                             #
    ' # REQUIREMENTS:                                                               #
    ' #    VBScript 5.5 or better                                                   #
    ' #    Scripting Run Time Library (Scrrun.dll)                                  #
    ' #    Adequate read/write/modify permissions in the specified upload directory #
    ' #    IIS 4 or 5 (preferably 5)                                                #
    ' #    ASP 2.0 or 3.0 (preferably 3.0)                                          #
    ' #                                                                             #
    ' #                                                                             #
    ' # FOR MORE INFORMATION AND FRESH DOWNLOADS:                                   #
    ' #    Detailed Information and Downloads Relating To This Class Can be found   #
    ' #    here:                                                                    #
    ' #      http://www.aspEmporium.com/aspEmporium/examples/fileupload_class.asp   #
    ' #                                                                             #
    ' #                                                                             #
    ' # DISCLAIMER:                                                                 #
    ' #   You are bound to the ASP Emporium license when you turn on and/or restart #
    ' #   your computer or look at the monitor for any reason... blah.              #
    ' #   http://www.aspemporium.com/aspEmporium/help/license.asp                   #
    ' #                                                                             #
    ' #                                                                             #
    ' ###############################################################################





	 ' ############################################
	 ' # CLASS DECLARATIONS                       #
	 ' # variables declared at the class level    #
	 ' # can be used freely in any routine within #
	 ' # the class delimiters                     #
	 ' ############################################

	Private strThisPage, lngUploadSize, boolAllowOverwrite
	Private UploadRequest, arrExt, strUploadDir, bCantCallInputs
	Private bMin, bByPass, bByPassMin, bTmpUploadSuccess
	Private sTmpUploadFilePath


	 ' ###################################################
	 ' # PROPERTY SETTINGS                               #
	 ' # Properties are set by a developer and enable    #
	 ' # and disable features of the class, as well as   #
	 ' # to control the overall useability of the system #
	 ' ###################################################

	Public Property LET Extensions(byVal arrayInput)
		arrExt = arrayInput
	End Property

	Public Property LET UploadDirectory(byVal strInput)
		strUploadDir = strInput
	End Property

	Public Property LET AllowOverWrite(byVal boolInput)
		boolAllowOverwrite = boolInput
	End Property

	Public Property LET UploadSize(byVal lngInput)
		lngUploadSize = CLng( lngInput )
	End Property

	Public Property LET MininumRequiredByteCountForUpload(byVal lngInput)
		bMin = lngInput
	End Property

	Public Property LET AllowUserToBypassUpload(byval boolInput)
		bByPass = boolInput
	End Property

	Public Property LET BypassUploadIfMinimumByteCountThresholdNotMet(byval boolInput)
		bByPassMin = boolInput
	End Property

	Public Property GET BypassUploadIfMinimumByteCountThresholdNotMet
		On Error Resume Next
		If Len(bByPassMin) < 4 or Len(bByPassMin) > 5 then
			bByPassMin = True
		Else
			bByPassMin = CBool(bByPassMin)
		End If
		BypassUploadIfMinimumByteCountThresholdNotMet = bByPassMin
	End Property

	Public Property GET AllowUserToBypassUpload()
		On Error Resume Next
		If Len(bByPass) < 4 or Len(bByPass) > 5 then
			bByPass = False
		Else
			bByPass = CBool(bByPass)
		End If
		AllowUserToBypassUpload = bByPass
	End Property

	Public Property GET MininumRequiredByteCountForUpload()
		On Error Resume Next
		if isnumeric(bMin) and bMin > 0 then
			MininumRequiredByteCountForUpload = cLng(bMin)
		else
			 ' default is 1k minimum
			MininumRequiredByteCountForUpload = 1024
		end if
	End Property

	Public Property GET Extensions()
		 ' allowable extensions
		If IsArray( arrExt ) then
			Extensions = arrExt
		ElseIf Len(arrExt) = 1 AND arrExt = "*" Then
			Extensions = "*"
		Else
			 ' default extensions
			Extensions = Array(".txt", ".htm", ".jpg", ".gif", ".inc")
		End If
	End Property

	Public Property GET UploadDirectory()
		 ' upload directory path
		If strUploadDir = Empty then
			 ' default upload path
			UploadDirectory = Server.Mappath("/")
		Else
			UploadDirectory = strUploadDir
		End If
	End Property

	Public Property GET AllowOverWrite()
		 ' allow file overwrite
		AllowOverWrite = CBool( boolAllowOverwrite )
	End Property

	Public Property GET UploadSize()
		 ' set maximum file size
		If lngUploadSize = 0 then
			 ' default maximum file size 97.7 K
			UploadSize = 100000
		Else
			UploadSize = lngUploadSize
		End If
	End Property


	' read-only property
	Public Property Get Version()
		Version = "1.2"
	End Property

	Public Property Get UploadSuccessful()
		UploadSuccessful = bTmpUploadSuccess
	End Property

	Public Property Get AbsolutePathToUploadedFile()
		AbsolutePathToUploadedFile = sTmpUploadFilePath
	End Property


	 ' ########################################
	 ' # EVENTS                               #
	 ' # Event code is automatically executed #
	 ' # when the specified event occurs      #
	 ' ########################################

	Private Sub Class_Initialize()
		strThisPage   = request.serverVariables("PATH_INFO")
		Call ClearVariables()
		Call ClearDictionary()
		bCantCallInputs = True
		bTmpUploadSuccess = False
	End Sub

	Private Sub Class_Terminate()
		strThisPage   = ""
		Call ClearVariables()
		Call ClearDictionary()
		Call TellWhereYouGotIt()
	End Sub


	 ' #################################################################
	 ' # PUBLIC METHODS                                                #
	 ' # Public methods are called by a developer to perform an action #
	 ' #################################################################

	Public Sub NoUploadMode()
		AllowUserToBypassUpload = True
	End Sub

	Public Sub OptionalUploadMode()
		BypassUploadIfMinimumByteCountThresholdNotMet = True
	End Sub

	Public Default Sub Upload()
		 ' public method
		Dim lngTotalBytes

		On Error Resume Next
		 ' total size of the file being uploaded
		lngTotalBytes = Request.TotalBytes
		 ' determine what the class should do:
		If lngTotalBytes <= 0 Then
			 ' no bytes means no form input 
			 ' so display the upload form
			Call ShowForm()
		ElseIf lngTotalBytes > 0 And lngTotalBytes <= UploadSize Then
			 ' if the total size of the file is 
			 ' within the set parameters, upload it.

			Call UploadBinary( lngTotalBytes )
		ElseIf lngTotalBytes > UploadSize Then
			 ' if the total size is larger than 
			 ' the set maximum make an error.
			Call FileTooBig_Error()
		Else
			 ' if something squeaks through, 
			 ' show the form because I don't know what it is
			Call ShowForm()
		End If
		On Error GoTo 0
	End Sub

	Public Function FormInput(sFormInput)
		If bCantCallInputs Then
			FormInput = ""
		Else
			FormInput = UploadRequest.Item(sFormInput).Item("Value")
		End If
	End Function

	Public Sub SaveAnotherBlob(byVal sFormInput)
		PrepForUpload sFormInput
	End Sub


	 ' ###############################################################
	 ' # UTILITY ROUTINES                                            #
	 ' # These routines handle and control the class file and upload #
	 ' ###############################################################

	Private Sub ClearVariables()
		lngUploadSize = 0
		arrExt        = Empty
		strUploadDir  = Empty
		boolAllowOverwrite = False
	End Sub

	Private Sub ClearDictionary()
		 ' clear dictionary
		Set UploadRequest = Server.CreateObject("Scripting.Dictionary")
		UploadRequest.RemoveAll
		Set UploadRequest = Nothing
	End Sub

	Private Sub QuitUpload()
		 ' show the link to the upload form
		Response.Write "<P><CENTER><A HREF=""" & strThisPage & """>"
		Response.Write "<B>Try Again</B></A></CENTER></P>"
	End Sub

	Private Sub UploadBinary(byVal MaxBytes)
		Dim byteCount, RequestBin

		 ' total size of the file
		byteCount = MaxBytes

		 ' read upload file
		RequestBin = Request.BinaryRead(byteCount)
		Set UploadRequest = Server.CreateObject("Scripting.Dictionary")
		 ' call BuildUploadRequest to parse binary info
		BuildUploadRequest  RequestBin

		bCantCallInputs = False
		 ' build upload request returns info in the'
		 ' scripting dictionary

		PrepForUpload "blob"
	End Sub

	Private Sub PrepForUpload(byval frmBlobInputName)
		Dim blobL, boolAllowed, filepathext
		Dim ContentType, filePathName, filename
		
		boolAllowed = False
		If Not AllowUserToBypassUpload Then
			 ' content type of the file (I don't check this - 
			 ' instead I check the extension)
			contentType = UploadRequest.Item(frmBlobInputName).Item("ContentType")
			 ' file path
			filepathname = UploadRequest.Item(frmBlobInputName).Item("FileName")
			 ' remove path and just get the file name...
			filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
			 ' get the extension
			filepathext = Right(filepathname, 4)
			blobL = LenB(UploadRequest.Item(frmBlobInputName).Item("Value")) 
			if blobL > MininumRequiredByteCountForUpload then 
				If IsArray(Extensions) Then
					 ' check the extension
					boolAllowed = ExtensionCheck( filepathext )
				ElseIf Extensions = "*" Then 
					boolAllowed = True
				End If
			else
				If Not BypassUploadIfMinimumByteCountThresholdNotMet Then
					Call MinimumByteThresholdNotMet_Error()
				Else
					bTmpUploadSuccess = True
				End If
				Exit Sub
			end if 
			If boolAllowed Then
				 ' get the upload directory and append the 
				 ' name of the file
				filename = UploadDirectory & "\" & "logo" & filepathext	'filename
				 ' create the new file on the server
				Call WriteUploadFile( filename, UploadRequest.Item( _
					frmBlobInputName).Item("Value") )
			Else
				Call BadExtension_Error()
			End If
		Else
			bTmpUploadSuccess = True
		End If
	End Sub

	Private Sub WriteUploadFile(byVal NAME, byVal CONTENTS)
		 ' create the file on the server
		dim ScriptObject, i, NewFile
		On error resume next
		Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
		If ScriptObject.FolderExists( UploadDirectory ) Then
			If NOT AllowOverWrite AND ScriptObject.FileExists( NAME ) Then
				 ' don't allow file overwrite if AllowOverWrite is False
				Set ScriptObject = Nothing
				Call FileAlreadyExists_Error()
				Exit Sub
			End If
			Set NewFile = ScriptObject.CreateTextFile( NAME )
			For i = 1 to LenB( CONTENTS )
				 ' translate binary data into ASCII 
				 ' characters and write them into the file.
				NewFile.Write chr( AscB( MidB( CONTENTS, i, 1) ) )
			Next
			 ' clean up and inform the user of successful upload
			NewFile.Close
			Set NewFile = Nothing
			sTmpUploadFilePath = NAME
			Call Success( NAME )
		Else
			 ' catch directory not found error
			Set ScriptObject = Nothing
			Call UploadDirNotFound_Error()
			Exit Sub
		End If
		Set ScriptObject = Nothing
	End Sub

	Private Function ExtensionCheck(byVal EXT)
		 ' check the input extension against 
		 ' all allowable extensions
		Dim Item

		For each item in Extensions
			if LCase( EXT ) = LCase( Item ) Then 
				ExtensionCheck = True
				Exit Function
			end if
		Next
		ExtensionCheck = False
	End Function


	 ' ###################################################
	 ' # HTTP WRITE ROUTINES                             #
	 ' # These routines write to the user's browser and  #
	 ' # allow the user to interact with the application #
	 ' ###################################################

	Private Sub Success(byVal FILE)
		bTmpUploadSuccess = True

		 ' display success message
'		Dim strRoot
'		strRoot = LCase( Server.Mappath("/") )
'		FILE = LCase( FILE )
'		FILE = Replace( Replace( FILE, strRoot, "" ), "\", "/" )
'		Response.Write "<H1>Upload Successful</H1>"
'		Response.Write "<P><BIG><B>" & FILE & "</B></BIG></P>"
'		Response.Write "<P><B><A HREF=""" & strThisPage & _
'			""">Upload Again</A></B></P>"
	End Sub

	Private Sub ShowForm()
		On Error Resume Next
		bTmpUploadSuccess = False

		 ' display the upload form and let the 
		 ' user know what they can and cannot upload
		Dim tmp, item
		With Response
			.Write("<H1>File Upload <CODE>v" & Version & "</CODE></H1>")
			.Write("<P>You can currently add any file of type: ")
			tmp = ""
			If IsArray(Extensions) Then
				For Each Item In Extensions
					tmp = tmp & "<CODE>*" & Item & "</CODE>, "
				Next
				tmp = left( tmp, Len(tmp) - 2 )
			Else
				tmp = "<CODE>All file types allowed</CODE>"
			End If
			.Write(tmp & "<BR>")
			.Write("Having a maximum size of: <CODE>~ ")
			.Write(Round( UploadSize / 1024, 1 ) & " k</CODE> ")
			.Write("and a minimum size of: <CODE>~ ")
			.Write(FormatNumber(Round( MininumRequiredByteCountForUpload _
				/ 1024, 1 ), 1) & " k.</CODE></P>")
			.Write("</P>")

			.Write("<FORM ENCTYPE=""multipart/form-data"" ACTION=""")
			.Write(strThisPage & """ METHOD=""POST"">" & vbCrLf)

			.Write("Please select a file to upload ")
			if AllowUserToBypassUpload Then
				.Write("from your computer [upload is disabled]:<BR>" & vbCrLf)
				.Write("<INPUT TYPE=FILE NAME=""blob"" DISABLED><BR><BR>" & vbCrLf)
			Else
				.Write("from your computer:")
				If BypassUploadIfMinimumByteCountThresholdNotMet Then
					.Write(" [Upload is optional]")
				Else
					.Write(" [upload is mandatory]")
				End If
				.Write("<BR>" & vbCrLf)
				.Write("<INPUT TYPE=FILE NAME=""blob""><BR><BR>" & vbCrLf)
			End If

			.write("<INPUT TYPE=FILE NAME=""blob2""><BR>")

			.Write("Please enter your full name:<BR>" & vbCrLf)
			.Write("<INPUT TYPE=TEXT NAME=""myName""><BR><BR>" & vbCrLf)
			.Write("<INPUT TYPE=SUBMIT VALUE=""Upload File"">" & vbCrLf)
			.Write("</FORM>" & vbCrLf)
		End With
	End Sub

	Private Sub TellWhereYouGotIt()
		With Response
			.Write vbCrLf
			.Write "<!--"
			.Write vbCrLf
			.Write "Created and Made Possible By The FileUpload Class"
			.Write vbCrLf
			.Write "http://www.aspEmporium.com/"
			.Write vbCrLf
			.Write "-->"
			.Write vbCrLf
		End With
	End Sub


	 ' ##########################################
	 ' # ERROR ROUTINES                         #
	 ' # error routines display if the uploaded # 
	 ' # file doesn't meet the set or default   #
	 ' # criteria.                              #
	 ' ##########################################

	Private Sub MinimumByteThresholdNotMet_Error()
		 ' didn't meet the minimum byte count required to upload.
		Response.Write "<H1>Upload File Too Small...</H1>"
		Response.Write "<P>You must upload a minimum size of <CODE>"
		On Error Resume Next
		Response.Write Round( MininumRequiredByteCountForUpload _
			/ 1024, 1 ) & " k.</CODE></P>"
		On Error GoTo 0
		Call QuitUpload()
	End Sub

	Private Sub FileTooBig_Error()
		 ' file too large error
		Response.Write "<H1>Upload File Too Large...</H1>"
		Response.Write "<P>You are currently limited to "
		Response.Write "uploading a maximum of <CODE>"
		On Error Resume Next
		Response.Write Round( UploadSize / 1024, 1 ) & " k</CODE>.</P>"
		On Error GoTo 0
		Call QuitUpload()
	End Sub

	Private Sub UploadDirNotFound_Error()
		 ' upload directory not found error
		response.write UploadDirectory
		Response.Write "<H1>Upload Directory Not Found ...</H1>"
		Response.Write "<P>The Specified Upload Directory Cannot Be Found.</P>"
		Call QuitUpload()
	End Sub

	Private Sub FileAlreadyExists_Error()
		 ' file already exists error
		Response.Write "<H1>Specified File Already Exists ...</H1>"
		Response.Write "<P>File OverWrite Disabled.</P>"
		Response.Write "<P>Rename the file on your computer and re-upload...</P>"
		Call QuitUpload()
	End Sub

	Private Sub BadExtension_Error()
		 ' bad file type error
		Dim Item, tmp

		Response.Write "<H1>Bad File Type...</H1>"
		Response.Write "<P>You are currently limited to "
		Response.Write "uploading only the following types of files:</P>"
		Response.Write "<P>"
		tmp = ""
		For Each Item In Extensions
			tmp = tmp & "<CODE>*" & Item & "</CODE>, "
		Next
		tmp = left( tmp, Len(tmp) - 2 )
		Response.Write "<P>" & tmp & "</P>"
		Call QuitUpload()
	End Sub


	 ' #########################################################
	 ' # UPLOAD ROUTINES                                       #
	 ' # For detailed information about these routines, go to: #
	 ' # http://www.asptoday.com/articles/20000316.htm         #
	 ' #########################################################

	Private Function getByteString(StringStr)
		 ' For detailed information about this routine, go to:
		 ' http://www.asptoday.com/articles/20000316.htm
		dim char, i

		For i = 1 to Len(StringStr)
			char = Mid(StringStr,i,1)
			getByteString = getByteString & chrB(AscB(char))
		Next
	End Function

	Private Function getString(StringBin)
		 ' For detailed information about this routine, go to:
		 ' http://www.asptoday.com/articles/20000316.htm
		dim intCount

		getString =""
		For intCount = 1 to LenB(StringBin)
			getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
		Next
	End Function

	Private Sub BuildUploadRequest(RequestBin)
		 ' For detailed information about this routine, go to:
		 ' http://www.asptoday.com/articles/20000316.htm
		dim PosBeg,PosEnd,boundary,boundaryPos,Pos,Name,PosFile
		dim PosBound,FileName,ContentType,Value

		PosBeg = 1
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
		boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
		boundaryPos = InstrB(1,RequestBin,boundary)
		Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
			Dim UploadControl
			Set UploadControl = Server.CreateObject("Scripting.Dictionary")
			Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
			Pos = InstrB(Pos,RequestBin,getByteString("name="))
			PosBeg = Pos+6
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
			PosBound = InstrB(PosEnd,RequestBin,boundary)
			If  PosFile<>0 AND (PosFile<PosBound) Then
				PosBeg = PosFile + 10
				PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
				FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "FileName", FileName
				Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
				PosBeg = Pos+14
				PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
				ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "ContentType",ContentType
				PosBeg = PosEnd+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Else
				Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
				PosBeg = Pos+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			End If
			UploadControl.Add "Value" , Value	
			UploadRequest.Add name, UploadControl	
			BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
		Loop
	End Sub
End Class
%>