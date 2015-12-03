# $language = "VBScript"
# $interface = "1.0"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                                                                                  '''
'''		AUDIT TOOL Version 1.1   ONERAN VERSION                                                                    '''
'''		DATE: 15/10/2015                                                                             '''
'''		DEVELOPED BY: Manuel Rivera (manuel.rivera@ericsson.com)                                     '''
'''		COMMENTS: Script developed in order to create site logs while perfomring an audit site }     '''
'''               configuration agains a defined customer standard, and with the intention of        '''
'''               increase the quality of GSM Mexico Deliveries                                      '''
'''		                                                                                             '''
'''     Script developed as a region and technology customization of the script NodeAudit from       '''                                                                                                
'''		Marco Antonio Fuentes working for sprint. (marco.antonio.fuentes@ericsson.com)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' GLOBAL VARIABLES DEFINITION
Dim signum, siteName, technology, rncName, TabRbs, TabRnc, TabRncPosition, numberOfTabs, xmlloaded, rncOutput, rbsOutput, logprompt
Dim promptstring, auxindex, customer, customerName, profileProperties, comandsProperties, cluster, LogsPath, RemotePathA, RemotePathB, RemotePathC
Dim TemplatePath, CmdPath, CVFlag, CVName, CVUser, Subject , Recipient, DateFormated, TimeFormated
Dim isIntegration, rbsCVComment, rncCVComment, shift, logtype, result, CurrentFolder, nodetype, commandString
Dim rncStrLog, rbsStrLog, rncLogName, rbsLogName, strlogtype, strCapture, no_cvs, ossstatus
Dim FileSys, g_Shell, xmlDoc,fsl, ConFile, CurrentPath, ToolPath, strreport, stractivity, remotefolderflag, rncstrUTC
Dim connectionString, sqlcommandString, auditid, strtechnology, commandElements, commandlist, rbsmodel, commandCustomer, RemoteUrl
Dim logflag, translationflag, translationString, validateFlag, validateType, regexpString, validateString, swname, sectorslist, numberofboards, txboards

'''CONSTANTS INICIALIZATION
profileProperties = 13
comandsProperties = 9

'''Variables initialization
remotefolderflag = false

'''DEFINING SCRIPTING OBJECTS TO BE USED
Set g_Shell = CreateObject("WScript.Shell")
Set FileSys = CreateObject("Scripting.FileSystemObject")
set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "false"

'''Start of Main Process
Sub Main
	
	''''constante de prueba
	'siteName = "WPDA1_SP"
	
	Set TabRbs = crt.GetScriptTab
	TabRbs.Screen.IgnoreEscape = True
	TabRbs.Screen.Synchronous = True
	
	'''GETTING USER ID
	signum = g_shell.expandenvironmentstrings("%USERNAME%")
	
	rncName = "NA"
	siteName = "NA"
	productName = "NA"
	'''GETTING SITE ID
	TabRbs.Screen.Send "pv nodename" & VbCr
	strCapture = TabRbs.Screen.ReadString(">")
	siteName = RegExpTest("nodename = (\w+)",strCapture)(0).SubMatches(0)
	
	'''GETTING SITE TECHNOLOGY AND INITIALIZING TABS FOR AUDITING NODES
	strCapture = ReadValueRbs("get 0 mimInfo")
	nodetype = RegExpTest("1.mimName =[ ]+(\w+)",strCapture)(0).SubMatches(0)
	strCapture = ReadValueRbs("get 0 productName")
	rbsmodel = RegExpTest("ManagedElement=1\s+productName\s+(\S+)",strCapture)(0).SubMatches(0)
	If InStr(nodetype,"ERBS") > 0 Then
		technology = false
	ElseIf InStr(nodetype,"RBS") > 0 Then
		technology = true
		TabRncPosition = crt.Dialog.prompt("Please type the position of the SecureCRT RNC Tab: " & vbCrlf & "Tab numeration begin in 1")
		
		If TabRncPosition = "" Then
			MsgBox("Not RNC Tab specified")
			Exit Sub
		End If
		
		numberOfTabs = crt.GetTabCount
		result = numberOfTabs - TabRncPosition
		If result < 0 Then
			Exit Sub
		Else
			Set TabRnc = crt.GetTab(TabRncPosition)
			TabRnc.Screen.IgnoreEscape = True
			TabRnc.Screen.Synchronous = True

			TabRnc.Screen.Send("pv nodename" & vbCr)
			strCapture = TabRnc.Screen.ReadString(">")
			rncName = RegExpTest("nodename = (\w+)",strCapture)(0).SubMatches(0)
		
			strCapture = ReadValueRnc("get 0 mimInfo")
			nodetype = RegExpTest("1.mimName =[ ]+(\w+)",strCapture)(0).SubMatches(0)
			If InStr(nodetype,"RNC") = 0 Then
				MsgBox("RNC Tab was not found")
				Exit Sub 
			End If
		
		End If
		
	Else
		Exit Sub
	End If
	
	'''SETTING CUSTOMER PROFILE
	xmlloaded = xmlDoc.Load("C:\RICC\PROFILES\customerprofiles.xml")
		
	If xmlloaded Then
		promptstring = "Pick a corresponding number for the customer: "
		auxindex = 1
		Set profileNodes=xmlDoc.selectNodes("/Profilesfile/Customerprofile")
		
		If profileNodes.length > 0 Then
			For Each objNode in profileNodes
				promptstring = promptstring & auxindex & "-" & objNode.getAttribute("customer") & " "
				auxindex = auxindex + 1
			Next
			
			customer = crt.Dialog.prompt(promptstring)
			
			If Not  IsNumeric(customer) then
				MsgBox("Invalid input for customer selection")
				Exit Sub
			End If
			
			customer = CInt(customer)
			
			If ((customer > 0) And (customer < auxindex)) Then
				customerName = profileNodes(customer-1).getAttribute("customer")
				If profileNodes(customer-1).hasChildNodes Then
				
					set propertyNodes = profileNodes(customer-1).ChildNodes
					
					If ((propertyNodes.length>0) And (propertyNodes.length <= profileProperties)) Then
						
						cluster = propertyNodes(0).text
						'MsgBox(cluster)
						LogsPath = propertyNodes(1).text
						'MsgBox(LogsPath)
						RemoteUrl = propertyNodes(2).text
						'MsgBox(RemoteUrl)
						RemotePathA = propertyNodes(3).text
						'MsgBox(RemotePathA)
						RemotePathB = propertyNodes(4).text
						'MsgBox(RemotePathB)
						RemotePathC = propertyNodes(5).text
						'MsgBox(RemotePathC)
						TemplatePath = propertyNodes(6).text
						'MsgBox(TemplatePath)
						CmdPath = propertyNodes(7).text
						'MsgBox(CmdPath)
						CVFlag = propertyNodes(8).text
						'MsgBox(CVFlag)
						CVName = propertyNodes(9).text
						'MsgBox(CVName)
						CVUser = propertyNodes(10).text
						'MsgBox(CVUser)
						Subject = propertyNodes(11).text
						'MsgBox(Subject)
						Recipient = propertyNodes(12).text
						'MsgBox(Recipient)
						
					Else
						MsgBox("Profile Not in Valid Format")
						Exit Sub
					End If
				Else
					MsgBox("Not Valid Profiles File")
					Exit Sub
				End If
			Else
				MsgBox("Not Valid Customer Specified")
				Exit Sub
			End If
		Else
			MsgBox("Not Valid Profiles File")
			Exit Sub
		End If
	Else
		MsgBox("Not Profiles File Found")
		Exit Sub
	End If
	xmlloaded = false
	
	If ((Len(RemoteUrl) > 0 ) And (RemoteUrl <> "NA")) Then
		strCmd = "net use G: " & RemoteUrl & " /persistent:yes /YES"	
		Set objExec = g_Shell.Exec(strCmd)
		remotefolderflag = true
	Else
		remotefolderflag = false
	End If
	
	'''CAPTURING LOG TYPE INFORMATION
	activityprompt = "Which activity are you performing?" & VbCrLf & _
						"1 - New RBS Node" & VbCrLf & _
						"2 - New RBS Node Mixed Mode" & VbCrLf & _
						"3 - Reconfiguration" & VbCrLf & _
						"4 - Cabinet/Board/Carrier Expansion" & VbCrLf & _
						"5 - Migration" & VbCrLf & _
						"6 - Rehome" & VbCrLf & _
						"7 - Swap" & VbCrLf & _
						"8 - License Load" & VbCrLf & _
						"9 - Support Services (Alarm Verification, Close Pendings)" & VbCrLf & _
						"10 - RNC Support Activities" & VbCrLf & _
						"11 - Trial" & VbCrLf & _
						"12 - Site Delete" & VbCrLf & _
						"13 - Feature Activation" & VbCrLf & _
						"14 - Antenna Swap"
	
	activityoption = crt.Dialog.prompt(activityprompt)
	
	If Not IsNumeric(activityoption) then
		MsgBox("Invalid input activity type")
		Exit Sub
	End If
	
	activityoption = Cint(activityoption)
	
	If ((activityoption > 0) And (activityoption < 15)) Then
	
		Select Case activityoption
			Case 1
				stractivity = "New RBS Node"
				isIntegration = crt.Dialog.prompt("New RBS Integration Finished Without Pendings? 1-YES 2-NO")
			Case 2
				stractivity = "New RBS Node Mixed Mode"
				isIntegration = crt.Dialog.prompt("New RBS Integration Finished Without Pendings? 1-YES 2-NO")
			Case 3
				stractivity = "Reconfiguration"
				isIntegration = 2
			Case 4
				stractivity = "Cabinet/Board/Carrier Expansion"
				isIntegration = 2
			Case 5
				stractivity = "Migration"
				isIntegration = 2
			Case 6
				stractivity = "Rehome"
				isIntegration = 2
			Case 7
				stractivity = "Swap"
				isIntegration = 2
			Case 8
				stractivity = "License Load"
				isIntegration = 2
			Case 9
				stractivity = "Support Services (Alarm Verification, Close Pendings)"
				isIntegration = 2
			Case 10
				stractivity = "RNC Support Activities"
				isIntegration = 2
			Case 11
				stractivity = "Trial"
				isIntegration = 2
			Case 12
				stractivity = "Site Delete"
				isIntegration = 2
			Case 13
				stractivity = "Feature Activation"
				isIntegration = 2
			Case 14
				stractivity = "Antenna Swap"
				isIntegration = 2
			Case Else
				MsgBox("Not Correct Option case")
				Exit Sub
		End Select
		
	Else
		MsgBox("Not Correct Option")
		Exit Sub
	End If
	
	
	
	If isIntegration = 1 Then
		If UCase(CVFlag) = "YES" Then
			rbsCVComment = "Final Integration"
			'stractivity = "Integration"
			If technology Then
				rncCVComment = crt.Dialog.prompt("Add a description for the RNC log?")
				If Len(rncCVComment) = 0 Then
					rncCVComment = "Not Comments Added"
				End If
			End If
		End If
		strlogtype = "INTEGRATION"
		CurrentPath = CurrentPath & "INTEGRATION\"
	ElseIf isIntegration=2 Then
		'stractivity = crt.Dialog.prompt("Which activity are you performing?")
							
		logtype = crt.Dialog.prompt("Before or After Log? 1-BEFORE  2-AFTER")
		If Not  IsNumeric(logtype) then
			MsgBox("Invalid input for log type selection")
			Exit Sub
		Else
			logtype = CInt(logtype)
			If logtype = 1 Then
				CurrentPath = CurrentPath & "PRECHECK\"
				strlogtype = "PRECHECK"
			ElseIf logtype = 2 Then
				CurrentPath = CurrentPath & "POSTCHECK\"
				strlogtype = "POSTCHECK"
			Else
				MsgBox("Invalid input for log type selection")
				Exit Sub
			End If
		End If	
		
		If UCase(CVFlag) = "YES" Then
			rbsCVComment = crt.Dialog.prompt("Add a description for the RBS log?")
			If Len(rbsCVComment) = 0 Then
				rbsCVComment = "Not Comments Added"
			End If
			If technology Then
				rncCVComment = crt.Dialog.prompt("Add a description for the RNC log?")
				If Len(rncCVComment) = 0 Then
					rncCVComment = "Not Comments Added"
				End If
			End If
		End If
	Else
		MsgBox("Invalid input - activity type does not exist")
		Exit Sub
	End If
	
	shift = crt.Dialog.prompt("Day or night maintenance window? 1-DAY 2-NIGHT 3-JME")
	If Not IsNumeric(shift) then
		MsgBox("Invalid input for Shift selection")
		Exit Sub
	Else
		shift = CInt(shift)
		If shift < 1 and shift > 3 Then
			MsgBox("Invalid input for Shift selection")
			Exit Sub
		End If
	End If

	'''SETTING FOLDERS FOR LOGS TRACKING
	y = Year(Date) 
    m = Month(Date) : If Len(m)=1 Then m = "0" & m : End If
    d = Day(Date) : If Len(d)=1 Then d = "0" & d : End If
	
	DateFormated  = y & "-" & m & "-" & d
	CurrentDate = y & m & d
	CurrentPath = CurrentPath & CurrentDate & "\"
	
	If  Not FileSys.FolderExists(LogsPath & CurrentPath) Then
		FileSys.CreateFolder(LogsPath & CurrentPath)
	End If
	
	CurrentPath = CurrentPath & siteName & "\"
	
	If  Not FileSys.FolderExists(LogsPath & CurrentPath) Then
		FileSys.CreateFolder(LogsPath & CurrentPath)
	End If
	
	'''GENERATING LOG FILES NAME
	If isIntegration = 1 Then
		rbsLogName = LogsPath & CurrentPath & DatePart("yyyy", Now)-2000 & Right("0" & DatePart("m",Now), 2) & Right("0" & DatePart("d",Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now) , 2)  & "_" & customerName & "_" & siteName & ".txt"	
	Else
		rbsLogName = LogsPath & CurrentPath & DatePart("yyyy", Now)-2000 & Right("0" & DatePart("m",Now), 2) & Right("0" & DatePart("d",Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now) , 2)  & "_" & strlogtype & "_" & customerName & "_" & siteName & ".txt"	
	End If
	
	If technology Then
		rncLogName = LogsPath & CurrentPath & DatePart("yyyy", Now)-2000 & Right("0" & DatePart("m",Now), 2) & Right("0" & DatePart("d",Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now) , 2)  & "_" & strlogtype & "_" & customerName & "_" & rncName & ".txt"
	End If
	
	
	''' SETING DATABASE VARIABLES
	connectionString = "Driver={MySQL ODBC 5.1 Driver};Server=146.250.116.41; Database=rlam_audit_tool; Uid=rlamuser; Pwd=rlamuser; Allow User Variables=True;"
	'connectionString = "Driver={MySQL ODBC 5.1 Driver};Server=127.0.0.1; Database=audittoolrlam; Uid=root; Pwd=; Allow User Variables=True;"
	
	CHour = DatePart("h" ,now) : if Len(CHour) = 1 then CHour = "0" & CHour : End if
	CMin = DatePart("n" ,now) : if Len(CMin) = 1 then CMin = "0" & CMin : End if
	CSeg = DatePart("s" ,now) : if Len(CSeg) = 1 then CSeg = "0" & CSeg : End if
	TimeFormated = CHour & ":" & CMin & ":" & CSeg
	CurrentTime = CHour & CMin & CSeg
	auditid = siteName & CurrentDate & CurrentTime
	auditid = UCase(Replace(auditid,"_",""))
	''' CALLING AUDIT FUNCTION
	
	sendRbsCommand("confb+")
	sendRbsCommand("gs+")
	If technology Then
		sendRncCommand("confb+")
		sendRncCommand("gs+")
	End If
	Call DU_StatusCheck()
	sendRbsCommand("gs-")
	sendRbsCommand("confb-")
	If technology Then
		sendRncCommand("gs-")
		sendRncCommand("confb-")
	End If
	
	MsgBox("AUDIT COMPLETED")
	
	'''DELETING OBJECTS
	Set FileSys = Nothing
	set xmlDoc = Nothing
	Set g_Shell = Nothing
	
	
End Sub

Sub DU_StatusCheck()
	
	'''INITIATING VALIDATION
	
	'''LOADING ALL MOs IN THE NODES
	sendRbsCommand("lt all")
	
	If technology Then
		sendRncCommand("lt all")
	End If
	
	''' VERIFYING AVAILABILITY FOR NEW CVs
	If UCase(CVFlag) = "YES" Then
		strCapture = ReadValueRbs("cvls")
		no_cvs = CInt(RegExpTest(">>> Total: (\d+) CV's", strCapture)(0).SubMatches(0))
		If no_cvs > 43 Then
			MsgBox "The Site has more than 43 CVs, please leave with less than de 43",0,"More than 43 CVs"
			Exit Sub	
		End If
		If technology Then
			strCapture = ReadValueRnc("cvls")
			no_cvs = CInt(RegExpTest(">>> Total: (\d+) CV's", strCapture)(0).SubMatches(0))
			If no_cvs > 43 Then
				MsgBox "The RNC has more than 43 CVs, please leave with less than 43",0,"More than 43 CVs"
				Exit Sub	
			End If
		End If
	End If
	
	''' VERIFYING SITE TX BOARDS
	strCapture = ReadValueRbs("st tx")
	'MsgBox(RegExpTest("^\s*Total:\s+(\d+)\s+MOs\s*$", strCapture)(0).SubMatches(0))
	
	If technology Then
		numberofboards = CInt(RegExpTest("^\s*Total:\s+(\d+)\s+MOs\s*$", strCapture)(0).SubMatches(0))
	Else
		numberofboards = 2
	End If

	''' VERIFYING IF SITE IS SINCHRONIZED TO START LOGING
	strutcstatus = "OK"
	TabRbs.Screen.Send "readclock" & VbCr
	strUTC = TabRbs.Screen.ReadString("Please enter Node Password: ", siteName & "> ")
	If TabRbs.Screen.MatchIndex = 1 Then
		TabRbs.Screen.Send "rbs" & VbCr
		strUTC = strUTC & TabRbs.Screen.ReadString(siteName & ">")
	End If
	
	If technology Then
		TabRnc.Screen.Send "readclock" & VbCr
		rncstrUTC = TabRnc.Screen.ReadString("Please enter Node Password: ", rncName & "> ")
		If TabRnc.Screen.MatchIndex = 1 Then
			TabRnc.Screen.Send "rnc" & VbCr
			rncstrUTC = rncstrUTC & TabRnc.Screen.ReadString(rncName & ">")
		End If
	End If
	
	strutcstatus = "OK"
	If isIntegration = 1 Then
		If InStr(1,strUTC,"UTC",1) = 0 Then
			MsgBox "!SITE STILL NOT UTC SYNCHRONIZED!" & vbCr & "Integration Log Can´t be created",20,"UTC FAIL"
			Exit Sub
		End If
	Else
		If InStr(1,strUTC,"UTC",1) = 0 Then
			If MsgBox("!SITE STILL NOT UTC SYNCHRONIZED!" & vbCr & "Are you sure you want to create the log?",20,"UTC FAIL") = vbNo Then
				Exit Sub
			End If
			strutcstatus = "NOK"
		End If
	End If
	
	
	'LICENCIA
	Dim licenseserver, integrationlicense, strlicense, lkfd, lkfe, lkfi, lkfstate, strlkf
	lkfd = 0
	lkfe = 0
	lkfi = 0
	strlicense = ""
	lkfstate = 0
	
	licenseserver = ReadValueRbs("license server")
	set Matches = RegExpTest("^License\s+Key\s+File.*$|^Emergency\s+Status.*$",licenseserver)
	
	sArray = split(licenseserver,vbCr,-1,1)
	uLimit = UBound(sArray)
	
	If Matches.Count = 2 Then
		If InStr(1,Matches(0),"Not Installed",1) > 0 Then
			If InStr(1,Matches(1),"Deactivated",1) > 0 Then
				'Dim integrationlicense
				integrationlicense = ReadValueRbs("license iu status")
				set Matches = RegExpTest("^State.*$",integrationlicense)
				If Matches.Count > 0 Then
					If InStr(1,Matches(0),"ACTIVATED",1) > 0 Then
						lkfi = 2
					End If
				End If
			Else
				lkfe = 3
			End If
		Else
			lkfd = 1 
			If InStr(1,Matches(1),"Deactivated",1) <= 0 Then
				lkfe = 2
			End If
		End If
	End If
	
	lkfstat = lkfd +lkfi + lkfe

	Select Case lkfstat
		Case 1 
		  strlicense = "FINAL LICENSE"
		Case 2 
		  strlicense = "INTEGRATION LICENSE"
		Case 3 
		  strlicense = "EMERGENCY LICENSE"
		Case Else
		  strlicense = "NOT LICENSE"
	End Select
	
	''' VERIFYING OSS CREATION
	TabRbs.Screen.Send "!smocpp" & VbCr
	
	nResult = TabRbs.Screen.WaitForStrings ("smocpp>","CORBA system exception: org.omg.CORBA.TIMEOUT: Connection Timed out  vmcid: OMG  minor code: 0  completed: No",20)
		
	If nResult = 1 Then
		TabRbs.Screen.Send "findnode -ne " & siteName & VbCr
		strCapture = TabRbs.Screen.ReadString("smocpp>")
		TabRbs.Screen.Send "exit" & VbCr
		TabRbs.Screen.WaitForString (siteName & ">")
		
		set Matches = RegExpTest("^\s*Operation\s+terminated\.\s+Node\(s\)\s+non-existing:(?:\s+(\S+)\s*)+$",strCapture)
		If Matches.Count > 0 Then
			ossstatus = "NOK"
		Else
			ossstatus = "OK"
		End If
		
	Else
		ossstatus = "NOK"
		TabRbs.Screen.Send chr(asc("C") - 64) & VbCr
		TabRbs.Screen.WaitForString (siteName & ">")
	End If
	

	''' CREATING LOG CV
	If UCase(CVFlag) = "YES" Then
		'strCapture = ReadValueRbs("cvcu")
		'swname = Replace(RegExpTest(".*UpgradePackage=([^\s]*)",strCapture)(0).SubMatches(0),"/","%")
		
		sendRbsCommand("get ConfigurationVersion=1 currentUpgradePackage > $current")
		sendRbsCommand("get $current administrativedata > $currentid1")
		sendRbsCommand("$currentid2 = $currentid1[productNumber] -s / -r %")
		sendRbsCommand("$currentid = $currentid1[productRevision] -s / -r % -g\r")
		sendRbsCommand("$userid = " & signum)
		'sendRbsCommand("$swname = " & swname)
		sendRbsCommand("$date = `date +%y%m%d`")
		sendRbsCommand("$time = `date +%H%M`")
		
		vWaitFors = Array("Are you Sure [y/n] ?",siteName & "> ",rncName & "> ")
		TabRbs.Screen.Send "lset ConfigurationVersion=1 autoCreatedCVIsTurnedOn true" & VbCr
		nResult = TabRbs.Screen.WaitForStrings(vWaitFors)
		
		If nResult = 1 Then
			TabRbs.Screen.Send "y" & VbCr
			TabRbs.Screen.WaitForStrings(vWaitFors)
		End If
		
		TabRbs.Screen.Send "lset ConfigurationVersion=1 rollbackOn true" & VbCr
		nResult = TabRbs.Screen.WaitForStrings(vWaitFors)
		If nResult = 1 Then
			TabRbs.Screen.Send "y" & VbCr
			TabRbs.Screen.WaitForStrings(vWaitFors)
		End If
		
		TabRbs.Screen.Send "cvms " & CVName & " " & CVUser & " " & rbsCVComment & VbCr
		TabRbs.Screen.WaitForString ("Total:")
		TabRbs.Screen.WaitForString ("Total:")
		TabRbs.Screen.WaitForString siteName & "> "
		
		If technology Then
		
			'sendRncCommand("cvcu")
			'swname = Replace(RegExpTest(".*UpgradePackage=([^\s]*)",strCapture)(0).SubMatches(0),"/","%")
			
			sendRncCommand("get ConfigurationVersion=1 currentUpgradePackage > $current")
			sendRncCommand("get $current administrativedata > $currentid1")
			sendRncCommand("$currentid2 = $currentid1[productNumber] -s / -r %")
			sendRncCommand("$currentid = $currentid1[productRevision] -s / -r % -g\r")
			
			sendRncCommand("$userid = " & signum)
			'sendRncCommand("$swname = " & swname)
			sendRncCommand("$date = `date +%y%m%d`")
			sendRncCommand("$time = `date +%H%M`")
			
			TabRnc.Screen.Send "lset ConfigurationVersion=1 autoCreatedCVIsTurnedOn true" & VbCr
			nResult = TabRnc.Screen.WaitForStrings(vWaitFors)
			If nResult = 1 Then
				TabRnc.Screen.Send "y" & VbCr
				TabRnc.Screen.WaitForStrings(vWaitFors)
			End If
			
			TabRnc.Screen.Send "lset ConfigurationVersion=1 rollbackOn true" & VbCr
			nResult = TabRnc.Screen.WaitForStrings(vWaitFors)
			If nResult = 1 Then
				TabRnc.Screen.Send "y" & VbCr
				TabRnc.Screen.WaitForStrings(vWaitFors)
			End If
			
			TabRnc.Screen.Send "cvms " & CVName & " " & CVUser & " " & rncCVComment & VbCr
			TabRnc.Screen.WaitForString ("Total:")
			TabRnc.Screen.WaitForString ("Total:")
			TabRnc.Screen.WaitForString rncName & "> "
		
		End If
	
	End IF
	
	''' GETTING COMANDS LIST AND LOOPING COMANDS EXECUTION
	strtechnology = "4G"
	If technology Then
		strtechnology = "3G"
		sectorslist = getSectors
		sendRncCommand("$sectors = " & sectorslist)
	End If
	
	xmlloaded = xmlDoc.Load(CmdPath)	
	If xmlloaded Then
		Set commandlist = xmlDoc.selectNodes("/commandsfile/commandlist[@technology='" & strtechnology & "']/command")	
		If commandlist.length > 0 Then
			For Each objNode in commandlist
				If objNode.hasChildNodes Then
					Set commandElements = objNode.ChildNodes
					If (commandElements.length = comandsProperties) Then
						
						commandString = commandElements(0).text
						nodetype = commandElements(1).text
						logflag = commandElements(2).text
						translationflag = commandElements(3).text
						translationString = commandElements(4).text
						validateFlag = commandElements(5).text
						validateType = commandElements(6).text
						txboards = CInt(commandElements(7).text)
						commandCustomer = commandElements(8).text
						
						Call processCommand()
						
					Else
						MsgBox("Command Not In Valid Format")
						Exit Sub
					End IF
				Else
					MsgBox("Command Not In Valid Format")
					Exit Sub
				End If
			Next
		Else
			MsgBox("Not Command List Found")
			Exit Sub
		End If
	Else
		MsgBox("Not Commands File Found")
		Exit Sub
	End If
	xmlloaded = false
	
	'''UPDATING AUDIT SITE REGISTRY
	
	strengineername = DatabaseReadValue("SELECT EngineerName FROM engineerlist WHERE signum='" & signum & "';")
	strcluster = DatabaseReadValue("SELECT Cluster FROM engineerlist WHERE signum='" & signum & "';") 
	cmd = 	"INSERT INTO `data` " & _
			"(AuditID, ReportType, `DATE`, `TIME`, Cluster, Customer, SiteName, RncName, Technology, Activity, UTCStatus, License, OSSStatus, EngineerUserId, EngineerName, EmailStatus, AuditStatus)" & _
			"VALUES ('" & auditid & "', '" & strlogtype & "', '" & DateFormated & "', '" & TimeFormated & "', '" & strcluster & "', '" & customerName & "', '" & SiteName & _
			"', '" & RncName & "', '" & strtechnology & "', '" &  stractivity & "', '" &  strutcstatus & "', '" &  strlicense & "', '" &  ossstatus & "', '" & signum & "', '" & strengineername & "', 'UNDEF', 'UNDEF');"
	'MsgBox(cmd)
	DatabaseUpdateCommand cmd
	
	'''CREATING LOGS
	If Len(rbsStrLog)>0 Then
		rbsStrLog = "LOG CREATED BY " & signum & " AT " & DateFormated & VbCrLf & VbCrLf & rbsStrLog
		Call createLogs(rbsLogName,rbsStrLog)
	End If
	
	If Len(rncStrLog)>0 Then
		rncStrLog = "LOG CREATED BY " & signum & " AT " & DateFormated & VbCrLf & VbCrLf & rncStrLog
		Call createLogs(rncLogName,rncStrLog)
	End If
	
	If remotefolderflag Then
		If FileSys.FileExists(rbsLogName)  Then
			If ((isIntegration=1) And (Len(RemotePathA) > 0)) Then
				FileSys.CopyFile rbsLogName, RemotePathA , True
			Else
				If shift = 1 Then
					If (Len(RemotePathB) > 0) Then
						FileSys.CopyFile rbsLogName, RemotePathB , True
					End If
				Else
					If (Len(RemotePathC) > 0) Then
						FileSys.CopyFile rbsLogName, RemotePathC , True
					End If
				End If
			End If				 
		End If
		
		If technology Then
			If shift = 1 Then
				If (Len(RemotePathB) > 0) Then
					FileSys.CopyFile rncLogName, RemotePathB , True
				End If
			Else
				If (Len(RemotePathC) > 0) Then
					FileSys.CopyFile rncLogName, RemotePathC , True
				End If
			End If
		End If
	End If
	
	If remotefolderflag Then
		If isIntegration = 1 Then
			If (Len(RemotePathA) > 0) Then
				g_Shell.Run "explorer.exe /e," & RemotePathA
			End If
		End If
		
		If technology Then
		
			If shift = 1 Then
				If (Len(RemotePathB) > 0) Then
					g_Shell.Run "explorer.exe /e," & RemotePathB
				End If
			Else
				If (Len(RemotePathC) > 0) Then
					g_Shell.Run "explorer.exe /e," & RemotePathC
				End If
			End If

		End If
	End If
	
	''' CLOSING VALIDATION
	'TabRbs.Screen.Send "b" & VbCr
End Sub


Sub processCommand
	rbsOutput = ""
	rncOutput = ""
	If ((commandCustomer="ALL") OR (inStr(1,commandCustomer,customerName,1) > 0)) Then
		If txboards <= numberofboards Then
			Select Case UCase(nodetype)
				Case "ANY"
					rbsOutput = ReadValueRbs(commandString)
					validatecommand rbsOutput,"RBS"
					If technology Then
						rncOutput = ReadValueRnc(commandString)
						validatecommand rncOutput,"RNC"
					End If
					If UCase(translationflag) = "YES" Then
						rbsOutput = Replace(rbsOutput,commandString,translationString)
						If technology Then
							rncOutput = Replace(rncOutput,commandString,translationString)
						End If
					End If
				Case "RBS"
					rbsOutput = ReadValueRbs(commandString)
					validatecommand rbsOutput,"RBS"
					If UCase(translationflag) = "YES" Then
						rbsOutput = Replace(rbsOutput,commandString,translationString)
					End If
				Case "RNC"
					If technology Then
						rncOutput = ReadValueRnc(commandString)
						validatecommand rncOutput,"RNC"
					End If
					If UCase(translationflag) = "YES" Then
						If technology Then
							rncOutput = Replace(rncOutput,commandString,translationString)
						End If
					End If
				Case Else
					MsgBox("Node Type for command" & commandString & " not specified")
					Exit Sub
			End Select
			
			If UCase(logflag) = "YES" Then
				rbsStrLog = rbsStrLog & rbsOutput
				If technology Then
					rncStrLog = rncStrLog & rncOutput
				End If
			End If
		End If
	End If
End Sub

Sub validatecommand(strtovalidate,strnodetype)
	If UCase(validateFlag) = "YES" Then
		Select Case CInt(validateType)
			Case 1
				validationtypea strtovalidate,strnodetype
			Case 2
				validationtypeb strtovalidate,strnodetype
			Case 3
				validationtypec strtovalidate,strnodetype
			Case 4
				validationtyped strtovalidate,strnodetype
			Case 5
				validationtypee strtovalidate,strnodetype
			Case Else
				Exit Sub
		End Select
	End If
End Sub

Sub validationtypea(strtovalidate,strnodetype)
	If Len(strtovalidate)>0 Then
		set Matches = RegExpTest("^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})\s+(\D{1})\s+(.*)\s+(\S+)\s+\((.*)\)$",strtovalidate)
		For Each Match In Matches
			if (Match.SubMatches(3) <> "UplinkBaseBandPool_UlHwLessThanUlCapacity") And (Match.SubMatches(3) <> "DownlinkBaseBandPool_DlHwLessThanDlCapacity") Then
				cmd = "INSERT INTO alarms " & _
				"(AuditID, ReportType, NodeType, AlarmDate, AlarmTime, Severity, SpecificProblem, MO, Cause) VALUES " & _
				"('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "', '" & Match.SubMatches(0) & "', '" & Match.SubMatches(1) & "', '" & Match.SubMatches(2) & _ 
				"', '" & Match.SubMatches(3) & "', '" & Match.SubMatches(4) & "', '" & Match.SubMatches(5) & "');"
				
				If strnodetype = "RNC" Then
					iubSiteName = split(siteName,"_",-1,1)
					If InStr(1,Match.SubMatches(4),iubSiteName(0),1) > 0 Then
						DatabaseUpdateCommand cmd
					End If
				Else
					DatabaseUpdateCommand cmd
				End If
			End If
		Next
	End If
End Sub

Sub validationtypeb(strtovalidate,strnodetype)
	If Len(strtovalidate)>0 Then
		set Matches = RegExpTest("^\s+\d+(?:\s+|\s+\d{1}\s+\((\S+)\)\s+)\d{1}\s+\((\S+)\)\s+(\S+)$",strtovalidate)
		For Each Match In Matches
			'MsgBox(Match.Value)
			adstatus = Match.SubMatches(0)
			If Len(Match.SubMatches(0)) = 0 Then
				adstatus = "NA"
			End If
			cmd = 	"INSERT INTO mosstatus (AuditID, ReportType, NodeType, MO, AdminStatus, OperatStatus, AuditStatus)" & _ 
					"VALUES ('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "', '" & Match.SubMatches(2) & "', '" & adstatus & "', '" & Match.SubMatches(1) & "',IF(((AdminStatus!='UNLOCKED' AND AdminStatus!='NA') OR (OperatStatus!='ENABLED')),'NOK','OK'));"
			'MsgBox(cmd)
			DatabaseUpdateCommand cmd
		Next		
	End If
End Sub

Sub validationtypec(strtovalidate,strnodetype)
	If Len(strtovalidate)>0 Then
		set Matches = RegExpTest("^\s*(\S+)\s+(\S+)\s+((?:\S+|\S+\s+\(\S+\)))\s*$",strtovalidate)
		For Each Match In Matches
			If Match.SubMatches(0) <> "MO" And Match.SubMatches(0) <> "Total:" Then
				mo = Match.SubMatches(0)
				attr = Match.SubMatches(1)
				value = Match.SubMatches(2)
				
				Dim cn, rs, value
				set cn = CreateObject("ADODB.Connection")
				cn.connectionstring = connectionString'"DSN=TEST;DriverId=1046;MaxBufferSize=2048;PageTimeout=5;"
				cn.open
				set rs = CreateObject("ADODB.Recordset")
				rs.open "SELECT " & customerName & ", MO FROM standardization WHERE Attribute='" & attr & "';", cn , 3, 4
				
				if not rs.eof Then
					rs.MoveFirst
					while not rs.eof
						validtype = ""
						if rs(1)<>"ALL" Then
							If(InStr(mo,rs(1))>0 ) Then
								validtype = DatabaseReadValue("SELECT CompType FROM standardization WHERE MO='" & rs(1) & "' AND Attribute='" & attr & "';")
							
								rulestring=""
								arrayvalues = Split(rs(0),",")
								limit = UBound(arrayvalues)
									
								If rs(0) <> "Not Found" Then
									Select Case CInt(validtype)
										Case 1
											validstr = "!="
										Case 2
											validstr = "<"
										Case Else
											validstr = "NA"
									End Select
								End If
									
								For i = 0 To limit
									If i = limit Then
										rulestring = rulestring & "(Value" & validstr & "'" & arrayvalues(i) & "')"
									Else
										rulestring = rulestring & "(Value" & validstr & "'" & arrayvalues(i) & "') AND " 
									End If
								Next
									
								If rs(0) <> "Not Found" Then
									cmd = 	"INSERT INTO attributes (AuditID, ReportType, NodeType, MO, Attribute, Value, CorrectValue, AuditStatus) " & _
											"VALUES ('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "', '" & mo & "', '" & attr & "', '" & value & "', '" & rs(0) & "', IF((" & rulestring & "),'NOK','OK'));"
									'MsgBox(cmd)
									DatabaseUpdateCommand cmd
								End If
							End If
						Else
							validtype = DatabaseReadValue("SELECT CompType FROM standardization WHERE Attribute='" & attr & "';")
							rulestring=""
							arrayvalues = Split(rs(0),",")
							limit = UBound(arrayvalues)
									
							If rs(0) <> "Not Found" Then
								Select Case CInt(validtype)
									Case 1
										validstr = "!="
									Case 2
										validstr = "<"
									Case Else
										validstr = "NA"
								End Select
							End If
									
							For i = 0 To limit
								If i = limit Then
									rulestring = rulestring & "(Value" & validstr & "'" & arrayvalues(i) & "')"
								Else
									rulestring = rulestring & "(Value" & validstr & "'" & arrayvalues(i) & "') AND " 
								End If
							Next
									
							If rs(0) <> "Not Found" Then
								cmd = 	"INSERT INTO attributes (AuditID, ReportType, NodeType, MO, Attribute, Value, CorrectValue, AuditStatus) " & _
										"VALUES ('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "', '" & mo & "', '" & attr & "', '" & value & "', '" & rs(0) & "', IF((" & rulestring & "),'NOK','OK'));"
								'MsgBox(cmd)
								DatabaseUpdateCommand cmd
							End If
							
						End If
						
						rs.MoveNext						
					wend
				end if
				rs.close
				cn.close
	
				set rs.ActiveConnection = Nothing
				
				set cn = nothing
				set rs = nothing
				
			End If
		Next		
	End If
End Sub

Sub validationtyped(strtovalidate,strnodetype)
	If Len(strtovalidate)>0 Then
		set Matches = RegExpTest("^\s+\d+\s+(\d+)\s+(\S+)\s+(ON|OFF|16HZ|5HZ)\s+(ON|OFF|16HZ|5HZ)\s+(ON|OFF|16HZ|5HZ)(?:\s+|\s+\S+\s+)(\S+)\s+(\S+)\s+(\S+)\s+(\d+)\s+(?:(\S+C)\s+(\S+))*$",strtovalidate)
		For Each Match In Matches
			cmd = 	"INSERT INTO boards (AuditId, ReportType, NodeType, Apn, Board, FaultLed, OperLed, MaintLed, ProductNumber, Rev, Serial, Date, Temp, BoardStatus, AuditStatus) " & _
					"VALUES ('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "', " & Match.SubMatches(0) & ", '" & Match.SubMatches(1) & "', '" & Match.SubMatches(2) & "', '" & Match.SubMatches(3) & _
					"', '" & Match.SubMatches(4) & "', '" & Match.SubMatches(5) & "', '" & Match.SubMatches(6) & "', '" & Match.SubMatches(7) & "', '" & Match.SubMatches(8) & _
					"', '" & Match.SubMatches(9) & "', '" & Match.SubMatches(10) & "', IF((FaultLed!='OFF' OR OperLed!='ON' OR MaintLed!='OFF'),'NOK','OK'));"
			DatabaseUpdateCommand cmd
		Next
		auxsiteName = split(siteName,"_",-1,1)
		set Matches = RegExpTest("^\s*\d+\s+\d+\s+(port\S+|BXP\S+)(\s+|\s+R\S+\s+)(ON|OFF|16HZ|5HZ)\s+(ON|OFF|16HZ|5HZ)\s+(ON|OFF|16HZ|5HZ)\s+\S+\s+\S+\s+\S+\s+\d+.*(?:S\d+C\d+\s*|" & auxsiteName(0) & "\S+\s*)+$",strtovalidate)
		
		For Each Match In Matches
			cmd = 	"INSERT INTO vswr (AuditID, ReportType, NodeType, Port, VSWR1, VSWR2, VSWR3, VSWR4, RU, FaultLedStatus, OperationalLedStatus, MaitenanceLedStatus, AuditStatus) " & _
					"VALUES ('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "', '" & Match.SubMatches(0) & "', -1, -1, -1, -1, '" & Match.SubMatches(1) & "', '" & Match.SubMatches(2) & "', '" & Match.SubMatches(3) & "', '" & Match.SubMatches(4) & "', 'UNDEF');"
			'MsgBox(cmd)
			DatabaseUpdateCommand cmd
		Next
		
		set Matches = RegExpTest("^\s*\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+(?:TX(\d+)\(W/dBm\)\s+)+(?:VSWR(\d+)\s+\(RL\d+\)\s+)+.*$",strtovalidate)

		txcount = 0
		vswrcount = 0
		regexpString = ""
		
		
		For Each Match In Matches
			txcount = CInt(Match.SubMatches(0))
			vswrcount = CInt(Match.SubMatches(1))
		
			For count=1 to txcount
				regexpString = regexpString & "(?:\d+\s+|N/A\s+|-\s+|\d+\.\d+\s+\(\d+\.\d+\)\s+)"
			Next
			
			For count=1 to vswrcount
				regexpString = regexpString & "(\d+\s+|N/A\s+|-\s+|\d+\.\d+\s+\(\d+\.\d+\)\s+)"
			Next
			
			regexpString =  "^\s*\d+\s+\d+\s+(port\S+|BXP\S+)(?:\s+R\S+\s+|\s+)\S+\s+" & regexpString & "(?:S\d+C\d+\s+|" & auxsiteName(0) & "\S+\s+\S+\s*)+$"'regexpString & "(?:S\d+C\d+\s*|" & auxsiteName(0) & "\S+\s+\S+\s*)+$"
		
		
			set Matchesb = RegExpTest(regexpString,strtovalidate)
			For Each Matchb In Matchesb
				results = CInt(Matchb.SubMatches.count)-1
				cmd = 	"UPDATE vswr SET " 
				vswrcmd = ""
				For count=1 to vswrcount	
					If Len(Matchb.SubMatches(count)) > 0 Then
						vswrstring = Trim(Matchb.SubMatches(count))
						
						Set Matchesc = RegExpTest("(?:(\d+$)|(N/A)|(-)|\d+\.\d+\s+\((\d+\.\d+)\))",vswrstring)
						If (Len(Matchesc(0).SubMatches(0)) > 0) Then
							vswrcmd = vswrcmd & "VSWR" & count & "=" & Matchesc(0).SubMatches(0) & " "
						ElseIf Len(Matchesc(0).SubMatches(1)) > 0 Then
							vswrcmd = vswrcmd & "VSWR" & count & "=-1 "
						ElseIf Len(Matchesc(0).SubMatches(2)) > 0 Then
							vswrcmd = vswrcmd & "VSWR" & count & "=-1 "
						ElseIf Len(Matchesc(0).SubMatches(3)) > 0 Then
							vswrcmd = vswrcmd & "VSWR" & count & "=" & Matchesc(0).SubMatches(3) & " "
						Else 
							vswrcmd = vswrcmd & "VSWR" & count & "=-1 "
						End If
						
					End If
				Next
				
				vswrcmd = Replace(RTrim(vswrcmd)," ",",")
			
				vswrvalue = DatabaseReadValue("SELECT " & customerName & " FROM standardization WHERE Attribute='vswr';") 
				
				cmd = cmd & vswrcmd & ", AuditStatus=IF(((RU='') OR (FaultLedStatus!='OFF' OR OperationalLedStatus!='ON' OR MaitenanceLedStatus!='OFF') OR (VSWR1>0 AND VSWR1<" & vswrvalue & ") OR (VSWR2>0 AND VSWR2<" & vswrvalue & ") OR (VSWR3>0 AND VSWR3<" & vswrvalue & ") OR (VSWR4>0 AND VSWR4<" & vswrvalue & ")),'NOK','OK') WHERE AuditID='" & auditid & "' AND Port='" & Matchb.SubMatches(0) & "';"
				
				'MsgBox(cmd)
				DatabaseUpdateCommand cmd
				
			Next
		Next

		If technology Then
			
			set Matches = RegExpTest("^\s*(\S+)\s+-(?:\d+\.\d+|\d+)\s+-(?:\d+\.\d+|\d+)\s+((?:\d+\.\d+|\d+))\s+\S+\s+\S+\s*$",strtovalidate)
			For Each Match In Matches
				
				delta =  DatabaseReadValue("SELECT " & customerName & " FROM standardization WHERE Attribute='delta';")
				cmd = 	"INSERT INTO rssi (AuditID, ReportType, NodeType, Cell, delta, AuditStatus) "& _
						"VALUES ('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "','" & Match.SubMatches(0) & "', " & Match.SubMatches(1) & ", IF(delta>"& delta &",'NOK','OK'));"
				'MsgBox(cmd)
				DatabaseUpdateCommand cmd

			Next
		End If
		
	End If
End Sub

Sub validationtypee(strtovalidate,strnodetype)
	If Len(strtovalidate)>0 Then
		set Matches = RegExpTest("^\s*Time\s+Counter\s+(.*)\s*$",strtovalidate)
		
		Dim cells
		limit=0
		regexpaux ="^\d+:\d+\s+Int_RadioRecInterferencePwr\s+"
		For Each Match In Matches
			cells = Split(Match.SubMatches(0)," ")
			limit = UBound(cells)
			For i=0 To limit
				If i = limit Then
					regexpaux = regexpaux & "(\d+|-\d+\.\d+)\s*$"
				Else
					regexpaux = regexpaux & "(\d+|-\d+\.\d+)\s+" 
				End If
			Next
		Next
		
		numcells = CInt(limit)+1
		
		ReDim Values(numcells)
		
		For i=0 To limit
			Values(i) = 0
		Next
		
		set Matches = RegExpTest(regexpaux,strtovalidate)
		
		For Each Match In Matches
			For i=0 To limit
				auxvalue = (-1)*CDbl(Match.SubMatches(i))
				If auxvalue > 0 Then
					Values(i) = auxvalue
				End If
			Next
		Next
		
		rssi =  DatabaseReadValue("SELECT " & customerName & " FROM standardization WHERE Attribute='rssi';")
		
		For i=0 To limit
			cmd = 	"INSERT INTO rssi (AuditID, ReportType, NodeType, Cell, rssi, AuditStatus) "& _
						"VALUES ('" & auditid & "', '" & strlogtype & "', '" & strnodetype & "','" & cells(i) & "', " & Values(i) & ", IF(rssi<"& rssi &",'NOK','OK'));"
			DatabaseUpdateCommand cmd
		Next
		
	End If
End Sub

Sub DatabaseUpdateCommand(command)
	Dim cn, value
	set cn = CreateObject("ADODB.Connection")
	cn.connectionstring = connectionString
	cn.open
	cn.Execute = command
	cn.Close
End Sub

Function DatabaseReadValue(command)
	Dim cn, rs, value
	set cn = CreateObject("ADODB.Connection")
	set rs = CreateObject("ADODB.Recordset")
	cn.connectionstring = connectionString'"Server=146.250.116.41:3306; Database=rlam_audit_tool; Uid=rlamuser; Pwd=rlamuser; Allow User Variables=True;"
	cn.open
	rs.open command, cn
	if rs.eof Then
		DatabaseReadValue = "Not Found"
		cn.close
	else
		rs.MoveFirst
		while not rs.eof
			value = rs(0)
			rs.MoveNext
		wend
		cn.close
		DatabaseReadValue = value
	end if
End Function

Sub createLogs(logname,logstring)
	Set logFile = FileSys.CreateTextFile(logname, True)
	logFile.Write logstring
	logFile.Close
End Sub

Function RegExpTest(patrn, strng)
   Dim regEx, Match, Matches
   Set regEx = New RegExp
   regEx.Pattern = patrn
   regEx.Multiline = True
   regEx.IgnoreCase = True
   regEx.Global = True
   Set RegExpTest = regEx.Execute(strng)
End Function

Function ReadValueRbs(Command)
	TabRbs.Screen.Send Command & VbCr
	ReadValueRbs = siteName & ">" & TabRbs.Screen.ReadString(siteName & "> ")
End Function

Sub sendRbsCommand(Command)
	TabRbs.Screen.Send Command & VbCr
	TabRbs.Screen.WaitForString siteName & "> "
End Sub

Function ReadValueRnc(Command)
	TabRnc.Screen.Send Command & VbCr
	ReadValueRnc = rncName & ">" & TabRnc.Screen.ReadString(rncName & "> ")
End Function

Sub sendRncCommand(Command)
	TabRnc.Screen.Send Command & VbCr
	TabRnc.Screen.WaitForString rncName & "> "
End Sub

Function getSectors
	strSectors = ""
	If technology Then
		iubSiteName = split(siteName,"_",-1,1)
		iubName = "Iub_" & iubSiteName(0)
				
		strCapture = ReadValueRnc("get " & iubName)
		set sectors = RegExpTest("^.*UtranCell=(\S+)$",strCapture)
		
		For each sector in sectors
			strSectors = strSectors & sector.SubMatches(0) & " "
		Next
		getSectors = iubName & "|" & Replace(RTrim(strSectors)," ","|")
	Else		
		strCapture = ReadValueRbs("st eut")
		set sectors = RegExpTest("^.*EUtranCellFDD=(\w+)$",strCapture)
		
		For each sector in sectors
			strSectors = strSectors & sector.SubMatches(0) & " "
		Next
		getSectors = Replace(RTrim(strSectors)," ","|")
	End If
End Function