'==============================================================================
' NAME: SpaceMailer.vbs
'
' AUTHOR: Brian Sheaffer
' COMMENT: Gathers drive space info for each local hard drive
'			Will Log to DB, email only if alert threshold is breached.
'
' USE:	by default, alert will be issued if drive freespace is below 20%
'
' Command line switches:
'		/s		= Output to screen, not email
'		#		= a number for the freespace alert threshold
'		/d		= debug mode - writes to "dev" db and emails brian
'		email	= email (or emails) to replace default destination
'
'	ex. cscript spacemailer.vbs 15 email@address.3ld email2@address.3ld
'			15 will be the new alert threshold
'			email will go to email and email2 (well, prolly not actually)
'
' History:	
'	2007/07/27	BDS		Created the script
'	2007/10/26	BDS		Added Database Section
'	2009/02/24	BDS		Adding Server Comment and hardware info
'						Disabling Database writing (not being rolled-out yet)
'	2010/02/08	BDS		Added disk labels into report
'	2010/04/05	BDS		Added disks without drive letters
'						Fixed disk sizing/space rounding error (divide by 1024, not 1000)
'	2010/04/07	BDS		Added command line email support (blat.exe)
'	2010/04/08	BDS		Added sorting to the drive list
'						Added AD Lookup for comment
'   2018/05/10  BDS     Genericized and published to github
'==============================================================================

Option Explicit
On Error Resume Next

'DECLARE ALL GLOBAL VARIABLES
Dim oWMI, colItems, oItem, oNet, oDRS, oRS, oConn, wShell
Dim EMAILTo, msgBody, sB, sC, sName, sLabel, dbServer, dbFile
Dim AlertThresholdPercent, iPercent, iWidth, iWidthHolder
Dim bToScreen, bToEmail, bDebug, bDataSet, bAlert, bDoDB

'SET ALL CONSTANTS
Const sComputer = "."
Const sAsterisks = "*********************************************************************"
Const EMAILFrom = "######"
Const EMAILServer = "######"
Const EMAILSubject = "DiskSpace Report for "
Const ScriptVer = "3.4"
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Const adVarChar = 200
Const MaxCharacters = 60
Const adFieldIsNullable = 32

'SET ALL INITIAL VARIABLE VALUES (might be over-ridden by command line switches)
EMAILTo = "######"
AlertThresholdPercent = 20
bToScreen = False
bToEmail = True
bAlert = False
bDoDB = False
bDataSet = True
sC = ""
iWidth = 12
iWidthHolder = 12

'The database for PRODUCTION
'(overridden via the -D command line switch
DBServer = "srp005"
DBFile = "WebEng"

ParseCLI

If bToScreen Then bDoDB = False

'Get the WMI Object for all uses coming up...
Set oWMI = GetObject("winmgmts:\\" & sComputer & "\root\CIMV2")

'The wShell object gives us access to various pieces of system info....
'It also lets us write to the event Log
Set wShell = WScript.CreateObject("WScript.Shell")

'Set up the network object to pull the computer name
Set oNet = CreateObject("WScript.NetWork") 

'Establish our disconnected recordset for sorting
Err.Clear
Set oDRS = CreateObject("ADODB.Recordset")
If Err.Number = 0 Then
	oDRS.Fields.Append "Disk", adVarChar, MaxCharacters, adFieldIsNullable
	oDRS.Fields.Append "Format", adVarChar, MaxCharacters, adFieldIsNullable
	oDRS.Fields.Append "Size", adVarChar, MaxCharacters, adFieldIsNullable
	oDRS.Fields.Append "Free", adVarChar, MaxCharacters, adFieldIsNullable
	oDRS.Fields.Append "Percent", adVarChar, MaxCharacters, adFieldIsNullable
	oDRS.Fields.Append "Label", adVarChar, MaxCharacters, adFieldIsNullable
	Err.clear
	oDRS.Open
	If Err.Number <> 0 Then bDataSet = False
Else
	bDataSet = False
	Err.Clear
End If

'MAIN CODE
If bDoDB then OpenDB
Echo "System:      " & UCase(oNet.ComputerName)
Echo "Description: " & GetComment(oNet.ComputerName) & ""
Echo "Script Run:  " & Now()
Echo "Hardware:    " & GetModel
Echo "OpSys:       " & GetOS
Echo "Last Boot:   " & GetUptime
Echo "Processor:   " & GetProcessor
Echo sAsterisks

'Let's get the hard drives available 
Err.Clear
Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_Volume WHERE DriveType = 3", "WQL", _
		wbemFlagReturnImmediately + wbemFlagForwardOnly)
If Err.Number = 0 Then
	For Each oItem In colItems	
		'Seems some cluster drives aren't really there... this tests their validity.
		If IsNumeric(oItem.Freespace) then
			iPercent = oItem.Freespace / oItem.Capacity 'iPercent = oItem.Freespace / oItem.Size
			If CInt(Round(iPercent * 100, 0)) < CInt(AlertThresholdPercent) Then bAlert = True
			
			sName = ""
			sName = oItem.Name
			If Len(Trim(sName)) = 0 Then sName = "No Name"
			If Trim(sName) = "" Then sName = "No Name"
			If Left(sName, 4) = "\\?\" Then sName = "{DirMount}"
			If Right(sName, 1) = "\" Then sName = Left(sName, Len(sName) -1 )
			
			'This helps setup our max width for the first column
			if iWidthHolder < Len(sName) Then iWidthHolder = Len(sName)
			
			sB = PadLeft(sName) ' & " (" & oItem.FileSystem & ")") 
			sB = sB & GBNum(oItem.Capacity) & GBNum(oItem.FreeSpace) & PercNum(iPercent)
				
			sLabel = ""
			sLabel = oItem.Label
			If Len(Trim(sLabel)) = 0 Then sLabel = "No Label"
			If Trim(sLabel) = "" Then sLabel = "No Label"
			sB = sB & "   " & sLabel

			If bDataSet Then
				oDRS.AddNew
				oDRS("Disk") = sName
				oDRS("Format") = oItem.Filesystem
				oDRS("Size") = GBNum(oItem.Capacity)
				oDRS("Free") = GBNum(oItem.FreeSpace)
				oDRS("Percent") = PercNum(iPercent)
				oDRS("Label") = sLabel
				oDRS.Update
			End If 				
			sC = sC & sB & VbCrLf
			
			If bDoDB Then WriteDB oNet.ComputerName, oItem.DeviceId, oItem.FileSystem, oItem.Size, oItem.FreeSpace
		End If
	Next

	if bDataSet Then iWidth = iWidthHolder	
	Echo PadLeft("Disk") & PadRight("Size") & PadRight("Free") & PadPercent("Free") & "   " & "Label"

	If bDataSet Then
		oDRS.Sort = "Disk Asc"
		oDRS.MoveFirst
		Do Until oDRS.EOF
			sB = PadLeft(oDRS.Fields.Item("Disk"))
			sB = sB & PadRight(oDRS.Fields.Item("Size"))
			sB = sB & PadRight(oDRS.Fields.Item("Free"))
			sB = sB & PadPercent(oDRS.Fields.Item("Percent"))
			sB = sB & "   " & oDRS.Fields.Item("Label")
			If Len(Trim(Replace(sB, vbNull, ""))) > 0 Then Echo sB
			oDRS.MoveNext
		Loop
	Else
		sC = Left(sC, Len(sC) - 1)
		Echo sC
	End If
	
	
Else
	WriteEventLog "An Error occured using WMI: " & Err.Description
End If

Echo sAsterisks
Echo VbCrLf
If bDoDB then CloseDB
If bAlert Then 
	If not bToScreen then bToEmail = True
End If
If bToEmail = True Then EmailIt
WScript.Quit

Function GetModel()
	Const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20

	Dim sM, colItems, oItem

	Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", _
                  wbemFlagReturnImmediately + wbemFlagForwardOnly)

	For Each oItem In colItems
		sM = Trim(oItem.Model)
		if Len(sM) > 0 then exit for
	Next
	GetModel = TrimNoData(sM)
End Function

Function GetOS()
'This is where we gather the OS and version. This requires a lot of prettying up.
Dim colItems, OpSys, oItem, sServPack
Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each oItem In colItems
	If oItem.Caption <> "" Then
		OpSys = oItem.Caption ' & " " & oItem.CSDVersion
		If Not IsNull(oItem.OtherTypeDescription) Then
			If oItem.OtherTypeDescription = "R2" Then 
				OpSys = Replace(OpSys, "Server 2003", "Server 2003 R2")
			End if
		End If
 		sServPack = "SP" & oItem.ServicePackMajorVersion
' 		sOSInstallDate = oItem.InstallDate
' 		sOSInstallDate = TrimNoData(sOSInstallDate)
' 		If sOSInstallDate <> "No Data" Then sOSInstallDate = WMIDateStringToDate(sOSInstallDate)
		Exit For
	End If
Next
	OpSys = Replace(OpSys, "®", "")
	OpSys = Replace(OpSys, "(R)", "")
	OpSys = Replace(OpSys, "Microsoft ", "")
	OpSys = Replace(OpSys, " Edition", ",")
	OpSys = Replace(OpSys, "Service Pack ", "SP")
	OpSys = Replace(OpSys, "Server 2003 R2", "2003 R2 Server")
	OpSys = Replace(OpSys, "Server 2003", "2003 Server")
	OpSys = Replace(OpSys, "Server 2008", "2008 Server")
	OpSys = Replace(OpSys, "Server, Enterprise,", "Server Enterprise")
	OpSys = Replace(OpSys, "Server, Standard,", "Server Standard")
	OpSys = Replace(OpSys, "Enterprise Server", "Server Enterprise")
	OpSys = Replace(OpSys, "Standard Server", "Server Standard")
	OpSys = Replace(OpSys, " XP Professional ", " XP ")
	'OpSys = Replace(OpSys, "Windows ", "")
	OpSys = Trim(OpSys) & ", " & Trim(sServPack)
	If Right(OpSys, 1) = "," Then OpSys = Left(OpSys, Len(OpSys) - 1)
	OpSys = Replace(OpSys, ",,", ",")
	OpSys = Replace(OpSys, ",,", ",")
	GetOS = TrimNoData(OpSys)
End Function

Function GetComment(sName)
	Dim colItems, oItem, sValue, oRegistry, sKeyPath, sValueName
	Dim sDN
	Const HKEY_LOCAL_MACHINE = &H80000002

	Set oRegistry = GetObject("winmgmts:\\" & _ 
    	sName & "\root\default:StdRegProv")
 
	sKeyPath = "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
	sValueName = "srvcomment"
	oRegistry.GetStringValue HKEY_LOCAL_MACHINE,sKeyPath,sValueName,sValue

	If IsNull(sValue) Then 
		sDN = GetDN(sName)
		If sDN = "No Data" Then
			sValue = ""
		Else
			sValue = GetADComment(sDN)
		End If
	End If
	
	sValue = TrimNoData(sValue)
	If sValue = "No Data" Then sValue = "Comment not found on the System nor in AD"
	GetComment = sValue
End Function

Function GetADComment(sPath)
	Dim oComp, sADComment
	On Error Resume Next
	sADComment = ""
	Set oComp= GetObject("LDAP://" & sPath)
    If Err.Number <> 0 Then
      	WriteEventLog("Error looking up the comment in AD: Computer Not found  or: " & Err.Description)
    Else
		sADComment = oComp.Get("description")
		If Err.Number <> 0 Then sADComment = "No Data"
	End If
	GetADComment = TrimNoData(sADComment)
	Set oComp = Nothing
End Function

Function GetDN(sComputerName)
' Use the NameTranslate object to convert the NT name of the computer To
' the Distinguished name required for the LDAP provider. Computer names
' must end with "$". Returns comma delimited string to calling code.
	Dim oTrans, oDomain, sRetVal
	sRetVal = ""
	' Constants for the NameTranslate object.
	Const ADS_NAME_INITTYPE_GC = 3
	Const ADS_NAME_TYPE_NT4 = 3
	Const ADS_NAME_TYPE_1779 = 1
	Set oTrans = CreateObject("NameTranslate")
	Set oDomain = getObject("LDAP://rootDse")
	oTrans.Init ADS_NAME_INITTYPE_GC, ""
	oTrans.Set ADS_NAME_TYPE_NT4, oNet.UserDomain & "\" & sComputerName & "$"
	sRetVal = oTrans.Get(ADS_NAME_TYPE_1779)
	GetDN = TrimNoData(sRetVal)
End Function

Function GetProcessor()
	Dim colItems, oItem, sCPUName, sCPUSpeed
	On Error Goto 0
	Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_Processor", "WQL")
	For Each oItem In colItems
		sCPUSpeed = oItem.MaxClockSpeed
		sCPUName = oItem.Name 
		sCPUName = Replace(sCPUName, "Intel(R) ", "")
		sCPUName = Replace(sCPUName, "Pentium(R)", "Pentium")
		sCPUName = Replace(sCPUName, " CPU", "")
		sCPUName = Replace(sCPUName, "(TM)", "")
		sCPUName = Replace(sCPUName, "(R)", "")
		sCPUName = Replace(sCPUName, "  ", " ")
		sCPUName = Replace(sCPUName, "  ", " ")
		sCPUName = Replace(sCPUName, "  ", " ")
		sCPUName = Replace(sCPUName, "  ", " ")
		sCPUName = Replace(sCPUName, "  ", " ")
		If InStr(sCPUName, "@") > 1 Then sCPUName = Left(sCPUName, InStr(sCPUName, "@") - 1)
		Exit For
	Next
	sCPUName = colItems.Count & " @ " & Trim(sCPUName)
	sCPUSpeed = (sCPUSpeed / 1000) & "ghz"
	GetProcessor = sCPUName & " @ " & sCPUSpeed
End Function

Function GetUptime()
Dim oOS, cOS, dtmBootup, dtmLastBootUpTime, dtmSystemUptime
Set cOS = oWMI.ExecQuery("Select * from Win32_OperatingSystem")
For Each oOS in cOS
    dtmBootup = oOS.LastBootUpTime
    dtmLastBootUpTime = WMIDateStringToDate(dtmBootup)
    'dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
Next
GetUptime = dtmLastBootUpTime

End Function

Function PercNum(sNumber)
	'Here, we format a 0.xxxxxx number into a percent
	Dim sRet
	sRet = FormatNumber(sNumber * 100, 1) & "%"
	PercNum = PadPercent(sRet)
End Function

Function GBNum(sNumber)
	'Here, we format an number to a drive size (so 123 becomes 124b for bytes; 27232130 becomes 27mb)
	Dim sRet, iLen
 	If Left(sNumber / 1024, 1) = "9" Then
 		iLen = Len(sNumber) - 1
 	Else
		iLen = Len(sNumber)
 	End If
	Select Case iLen
		Case 0, 1, 2, 3
			'bytes
			sRet = FormatNumber(sNumber, 2) & "B"
		Case 4, 5, 6
			'kilobytes
			sRet = FormatNumber(sNumber / 1024, 2) & "kb"
		Case 7, 8, 9
			'megabytes
			sRet = FormatNumber(sNumber / 1048576, 2) & "mb"
		Case 10, 11, 12
			'gigabytes
			sRet = FormatNumber(sNumber / 1073741824, 2) & "gb"
		Case 13, 14, 15
			'terabytes
			sRet = FormatNumber(sNumber / 1099511627776, 2) & "tb"
		Case 16, 17, 18
			'petabytes
			sRet = FormatNumber(sNumber / 1125899906842624, 2) & "pb"
		Case 19, 20, 21
			'exabytes
			sRet = FormatNumber(sNumber / 1.152921504606847e+18, 2) & "eb"
		Case 22, 23, 24
			'zettabytes
			sRet = FormatNumber(sNumber / 1.180591620717411e+21, 2) & "zb"
		Case 25, 26, 27
			'yottabytes
			sRet = FormatNumber(sNumber / 1.208925819614629e+24, 2) & "yb"		
		Case Else
			'bytes
			sRet = FormatNumber(sNumber, 2) & "B"

	End Select
	sRet = PadRight(sRet)
	GBNum = sRet
End Function

Function PadRight(sOrig)
	'right justify a column of text so "123" becomese "   123"
	Dim sRet
	sRet = String(12 - Len(sOrig), " ") & sOrig
	PadRight = sRet
End Function

Function PadPercent(sOrig)
	'right justify a column of text so "123" becomese "   123"
	Dim sRet
	sRet = String(8 - Len(sOrig), " ") & sOrig
	PadPercent = sRet
End Function

Function PadLeft(sOrig)
	'left justify a column - padding the end with spaces
	Dim sRet
	If iWidth < Len(sOrig) Then iWidth = Len(sOrig) + 2
	sRet = sOrig & String(iWidth - Len(sOrig), " ")
	PadLeft = sRet
End Function

Sub Echo(sText)
	'Set up the output - either to the screen or to email
	If bToScreen Then WScript.Echo sText
	msgBody = msgBody & sText & VbCrLf
End Sub

Sub EmailIt()
	'Create an email object and send it
 	Dim objEmail, oNetwork, sMachineName, sSubject, bUseCMDEmailer
	Dim oApp, sCLI
	On Error Resume Next
 	
 	Set oNetwork = WScript.CreateObject("WScript.Network")
 	sMachineName = oNetwork.ComputerName
	sSubject = EMAILSubject & sMachineName
 	If bAlert Then sSubject = "ALERT! " & sSubject

	bUseCMDEmailer = False

	Err.Clear
	Set objEmail = CreateObject("CDO.Message")

	If objEmail Is Nothing Then bUseCMDEmail = True

	If Err.Number <> 0 Then 
		WriteEventLog("Error Sending Email: " & Err.Description)
		bUseCMDEmailer = True
	Else
		Err.Clear
		objEmail.AutoGenerateTextBody = False
		objEmail.From = EMAILFrom
		objEmail.To = EMAILTo
		objEmail.Subject = sSubject
		objEmail.Textbody = msgBody
		objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EMAILServer 
		objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objEmail.Configuration.Fields.Update
		objEmail.Send
		If Err.number <> 0 Then bUseCMDEmailer = True
		Set objEmail = Nothing
	End If

	If bUseCMDEmailer Then
		Err.Clear
		msgBody = Replace(msgBody, vbcrlf, "|")
		sCLI = "blat - -t " & EMAILTo & " -s " & Chr(34) & sSubject & Chr(34) & " -body " & Chr(34) & msgBody & Chr(34) & " -server " & EMAILServer & " -f " & EMAILFrom
		Set oApp = wShell.exec(sCLI)
		If Err.Number <> 0 Then
			WriteEventLog("Error found: " & Err.Description)
			Err.Clear
		End If
		'and wait for it To Exit
		do until oApp.status = 1 : wscript.sleep 10 : Loop
		'WriteToEventLog "Application Complete: " & sCLI
	End If
End Sub

Sub ParseCLI()
	'This reads throught the command line switches and acts accordingly
	Dim oArgs, i, sA, sEmailTos
	sEmailTos = ""
	Set oArgs = WScript.Arguments
	For i = 0 to oArgs.Count - 1
		sA = oArgs(i)
		'Let's see if this is a number
		If IsNumeric(sA) Then 
			'aha, if it's between 0 and 100 then it's our new Alert Threshold
			If (sA < 100) And (sA > 0) Then AlertThresholdPercent = sA
		End If
		If IsEmail(sA) Then
			sEmailTos = Trim(sA) & "," & Trim(sEmailTos)
		End If
		
		Select Case Left(UCase(sA), 2)
			Case "/?", "-?", "/H", "-H", "-HELP"
				DoHelp
			Case "/s", "/S", "-s", "-S"
				'Send the output to the screen, rather than an email
				bToScreen = True
				bToEmail = False
			Case "/d", "/D", "-d", "-D"
				bDebug = True
				'The database server for TESTING
				DBServer = "infosysdb\infosys"
				DBFile = "WebEng"
				EMAILTo = "me@address.3ld"
		End Select
	Next
	'Clean up the email addresses if necessary, and assign them as the destination.
	If Right(sEmailTos, 1) = "," Then sEmailTos = Left(sEmailTos, Len(sEmailTos) - 1)
	If Len(sEmailTos) > 0 Then EMAILTo = sEmailTos
End Sub

Function IsEmail(sText)
	'Rudimentary email address format verifier
	Dim bRet, iAt, iDot
	bRet = False
	iAt = InStr(sText, "@")
	iDot = InStr(sText, ".")
	'Basically, if there's an @ symbol, followed by a . then it's an email address
	If iAt > 0 Then 
		If iDot > iAt Then bRet = True
	End If
	IsEmail = bRet
End Function

Sub DoHelp
	Dim sB
	sB = ""
	sB = sB & string(76, "-") & VbCrLf
	sB = sB & "SpaceMailer" & VbCrLf
	sB = sB & string(76, "-") & VbCrLf
	sB = sB & "Gathers drive space information for all local hard drives and emails" & VbCrLf
	sB = sB & "that information in a report to admins." & VbCrLf
	sB = sB & VbCrLf
	sB = sB & "Usage: spacemailer [percent] [email1] [email2] [email...]" & VbCrLf
	sB = sB & VbCrLf
	sB = sB & "Options: "  & VbCrLf
	sB = sB & VbCrLf
	sB = sB & "percent    if freespace is below this percent, issues an alert (default=20)" & VbCrLf
	sB = sB & "email      email addresses for the report (separate by spaces; default=Is)" & VbCrLf
	sB = sB & VbCrLf
	sB = sB & VbCrLf
	
	WScript.Echo sB
	WScript.quit
End Sub

Sub OpenDB()
	Const adOpenStatic = 3
	Const adLockOptimistic = 3
	Const adUseClient = 3

	'The next 2 lines create the Connection and RecordSet objects 
	Set oConn = CreateObject("ADODB.Connection")
	Set oRS = CreateObject("ADODB.Recordset")

	'Just in case something failed earlier, we'll clear the error to be fresh
	Err.Clear

	'The next line opens the Connection to the specified server and database, using TCP/IP (DBMSS0CN)
	oConn.Open "DRIVER={SQL Server}; server=" & DBServer & "; database=" & DBFile & "; Network=DBMSSOCN; User id=LoginOutWriter; password=3v3nl3sss3cur3!"

	'If the Connection was successful, Then
	If Err.Number = 0 Then
		'Create a local recordset with fields based on the actual DB Table (much quicker for large DBs)
		oRS.CursorLocation = adUseClient
		oRS.Open "SELECT TOP 0 * FROM SystemSpace", oConn, adOpenStatic, adLockOptimistic
		If Err.Number <> 0 Then
			WriteEventLog "An Error occured selecting columns from the Database: " & Err.Description
			Err.Clear
		End If
	Else
		bToEmail = True
		WriteEventLog "An Error occured connecting to the Database: " & Err.Description
		Err.Clear
	End If
End Sub

Sub CloseDB
	oRS.Close
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing
End Sub 

Sub WriteDB(sSystem, sDrive, sFormat, sSize, sFree)
	'Now the database stuff.
	'Create a new row in the local recordset and populate our new info
	oRS.AddNew
	oRS("systemname") = sSystem
	oRS("driveletter") = Left(sDrive, 1)
	oRS("driveformat") = sFormat
	oRS("drivesizeb") = sSize
	oRS("drivefreeb") = sFree
	oRS("coldate") = Now()
	oRS("sversion") = ScriptVer

	'Send our new row to the server and save it
	oRS.Update
	If Err.Number <> 0 Then
		'Write the error to the event Log
		WriteEventLog "An Error occured updating the Database: " & Err.Description
		Err.Clear
	End If
End Sub

Sub WriteEventLog(sText)
	wShell.LogEvent 1, "SpaceMailer.vbs: " & sText
End Sub

Function TrimNoData(s)
	'Just a little function to clean up the strings a little
	'This strips any extra spaces and sets empty strings to "No Data"
	Dim sRet
	sRet = Trim(s)
	If Len(sRet) = 0 Then sRet = "No Data"
	TrimNoData = sRet
End Function

Function StringIsEmpty(s)
	'If the string is empty, this returns TRUE, otherwise FALSE
	StringIsEmpty = CBool(Len(Trim(s)) = 0)
End Function

Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate = CDate(Mid(dtmBootup, 5, 2) & "/" & _
         Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
         & " " & Mid (dtmBootup, 9, 2) & ":" & _
         Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, _
         13, 2))
End Function
