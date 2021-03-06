set filesys=CreateObject("Scripting.FileSystemObject")
Dim strCurDir
strCurDir = filesys.GetParentFolderName(Wscript.ScriptFullName)
set xmlhttp = createobject("msxml2.xmlhttp.3.0")
dim WPData 'Web page text - for MAC address CSV
Dim OIUdata 'Data from MAC address CSV file
Dim CSVdata 'Data from CSV
Dim outputl 'Email body
Dim adoconn
Dim rs
Dim str
Dim i 'Counter
Dim Response 'For answers to prompts
Dim PSSchema, PSTbl 'Define schema and table names
PSSchema = "pingscan"
PSTbl = "pingscan2"

outputl = ""

'Gather variables from psapp.ini or prompt for them and save them for next time
If filesys.FileExists(strCurDir & "\psapp.ini") then
	'Database
	CSVPath = ReadIni(strCurDir & "\psapp.ini", "Database", "CSVPath" )
	DBLocation = ReadIni(strCurDir & "\psapp.ini", "Database", "DBLocation" )
	DBUser = ReadIni(strCurDir & "\psapp.ini", "Database", "DBUser" )
	DBPass = ReadIni(strCurDir & "\psapp.ini", "Database", "DBPass" )
	
	'Email - Defaults to anonymous login
	RptToEmail = ReadIni(strCurDir & "\psapp.ini", "Email", "RptToEmail" )
	RptFromEmail = ReadIni(strCurDir & "\psapp.ini", "Email", "RptFromEmail" )
	EmailSvr = ReadIni(strCurDir & "\psapp.ini", "Email", "EmailSvr" )
	'Additional email settings found in Function SendMail()
	
	'Location Specific information for scanning
	Building = ReadIni(strCurDir & "\psapp.ini", "LocationSpecific", "Building" )
	SubnetDotZero = ReadIni(strCurDir & "\psapp.ini", "LocationSpecific", "SubnetDotZero" )
	SubnetStart = ReadIni(strCurDir & "\psapp.ini", "LocationSpecific", "SubnetStart" )
	SubnetEnd = ReadIni(strCurDir & "\psapp.ini", "LocationSpecific", "SubnetEnd" )
	DaysBeforeUntrusted = ReadIni(strCurDir & "\psapp.ini", "LocationSpecific", "DaysBeforeUntrusted" )
	PingExceptions = ReadIni(strCurDir & "\psapp.ini", "LocationSpecific", "PingExceptions" ) 'Manually add this entry to ini file if needed. MAC Addresses should be pipe separated.
	
	'WebGUI
	EditURL = ReadIni(strCurDir & "\psapp.ini", "WebGUI", "EditURL" )
	
	'MAC CSV (aka OUI.CSV)
	ConsistencyMAC = ucase(ReadIni(strCurDir & "\psapp.ini", "MACCSV", "ConsistencyMAC" ))
	OUIURL = ReadIni(strCurDir & "\psapp.ini", "MACCSV", "OUIURL" )
	OUIUpdateAfter = ReadIni(strCurDir & "\psapp.ini", "MACCSV", "OUIUpdateAfter" )
	OUIDaysToUpdate = ReadIni(strCurDir & "\psapp.ini", "MACCSV", "OUIDaysToUpdate" )
else
	msgbox "INI file not found at: " & strCurDir & "\psapp.ini" & vbCrlf & "You will now be prompted with questions to create it."
	
	'Database
	CSVPath = inputbox("Enter the location where the PingScan data should be saved during processing (UNC path recommended):", "PingScan 2.0", strCurDir & "\PingScan.csv")
	DBLocation = inputbox("Enter the IP address or hostname for the location of the database:", "PingScan 2.0", "localhost")
	DBUser = inputbox("Enter the user name to access database on " & DBLocation & ":", "PingScan 2.0", "user")
	DBPass = inputbox("Enter the password to access database on " & DBLocation & ":", "PingScan 2.0", "P@ssword1")
	
	'Check to see if DB exists
	CheckForTables
	
	'Email - Defaults to anonymous login
	RptToEmail = inputbox("Enter the report email's To address:", "PingScan 2.0", "admin@company.com")
	RptFromEmail = inputbox("Enter the report email's From address:", "PingScan 2.0", "admin@company.com")
	EmailSvr = inputbox("Enter the FQDN or IP address of email server:", "PingScan 2.0", "mail.server.com")
	msgbox "Additional email settings found in Function SendMail()"
	
	'Location Specific information for scanning
	Building = inputbox("Enter the location (or building) of this scanner:", "PingScan 2.0", "Main Office")
	SubnetDotZero = inputbox("Enter the subnet IP address to scan ending in zero (0):", "PingScan 2.0", "192.168.1.0")
	SubnetStart = inputbox("Enter the first IP to scan (last octet only):", "PingScan 2.0", "1")
	SubnetEnd = inputbox("Enter the last IP to scan (last octet only):", "PingScan 2.0", "254")
	DaysBeforeUntrusted = inputbox("Enter the amount of days before a trusted computer is considered untrusted:", "PingScan 2.0", "7")
	
	'WebGUI
	EditURL = inputbox("Enter the URL to be used for editing a devices details (public version not available yet so this can be left blank):", "PingScan 2.0", "http://www.intranet.com/pingscan/update_device.asp?ID=")
	
	'MAC CSV (aka OUI.CSV)
	ConsistencyMAC = "B827EB" 'Used to test CSV file
	Response = msgbox("Would you like to set up the MAC address CSV to be download automatically on a reoccurring basis?", vbYesNo) 'Ask whether we should download
	if Response = vbYes then
		OUIURL = inputbox("Enter the URL where we can download the current MAC ""database"" in CSV format:", "PingScan 2.0", "http://standards-oui.ieee.org/oui/oui.csv")
		OUIDaysToUpdate = inputbox("Enter how often (amount of days) to check for an updated CSV file from the website provided:", "PingScan 2.0", "30")
		OUIUpdateAfter = format(Date() + OUIDaysToUpdate, "YYYYMMDD") 'calculate date based on days
	else
		OUIURL = "http://standards-oui.ieee.org/oui/oui.csv"
		OUIDaysToUpdate = 0
		OUIUpdateAfter = ""
		
		Response = msgbox("You've chosen not to update on a regular bases. Would you like to download it one time now from " & OUIURL & "?", vbYesNo) 'download once?
		if Response = vbYes then
			xmlhttp.open "get", OUIURL, false
			xmlhttp.send
			WPData = xmlhttp.responseText
			WriteUTF strCurDir & "\oui.csv", WPData
			if instr(1,WPData,ConsistencyMAC,1) = 0 then msgbox "Consistency check on the CSV failed. Please verify that the data looks correct at: " & strCurDir & "\oui.csv"
		end if
	end if
	
	'Write the data to INI file
	WriteIni strCurDir & "\psapp.ini", "Database", "CSVPath", CSVPath
	WriteIni strCurDir & "\psapp.ini", "Database", "DBLocation", DBLocation
	WriteIni strCurDir & "\psapp.ini", "Database", "DBUser", DBUser
	WriteIni strCurDir & "\psapp.ini", "Database", "DBPass", DBPass
	WriteIni strCurDir & "\psapp.ini", "Email", "RptToEmail", RptToEmail
	WriteIni strCurDir & "\psapp.ini", "Email", "RptFromEmail", RptFromEmail
	WriteIni strCurDir & "\psapp.ini", "Email", "EmailSvr", EmailSvr
	WriteIni strCurDir & "\psapp.ini", "LocationSpecific", "Building", Building
	WriteIni strCurDir & "\psapp.ini", "LocationSpecific", "SubnetDotZero", SubnetDotZero
	WriteIni strCurDir & "\psapp.ini", "LocationSpecific", "SubnetStart", SubnetStart
	WriteIni strCurDir & "\psapp.ini", "LocationSpecific", "SubnetEnd", SubnetEnd
	WriteIni strCurDir & "\psapp.ini", "LocationSpecific", "DaysBeforeUntrusted", DaysBeforeUntrusted
	WriteIni strCurDir & "\psapp.ini", "WebGUI", "EditURL", EditURL
	WriteIni strCurDir & "\psapp.ini", "MACCSV", "ConsistencyMAC", ConsistencyMAC
	WriteIni strCurDir & "\psapp.ini", "MACCSV", "OUIURL", OUIURL
	WriteIni strCurDir & "\psapp.ini", "MACCSV", "OUIUpdateAfter", OUIUpdateAfter
	WriteIni strCurDir & "\psapp.ini", "MACCSV", "OUIDaysToUpdate", OUIDaysToUpdate
end if

'Check to see if MAC address table exists and if so, use it
If filesys.FileExists(strCurDir & "\oui.csv") then
	OIUdata = ReadUTF(strCurDir & "\oui.csv")
	OIUdata = replace(OIUdata,"'","") 'Replace any single quotes in the MAC address CSV as the database doesn't like them
else
	OIUdata = ""
end if

'If option is selected, check the MAC address CSV for updates
if format(Date(), "YYYYMMDD") => OUIUpdateAfter and OUIDaysToUpdate > 0 then
	xmlhttp.open "get", OUIURL, false
	xmlhttp.send
	WPData = xmlhttp.responseText
	
	if replace(WPData,"'","") = OIUdata then
		'msgbox "Awesome!"
	else
		'msgbox len(OIUdata) & " --> " & len(WPData)
		'msgbox left(WPData,50)
		if instr(1,WPData,ConsistencyMAC,1) = 0 then 'Consistency Check
			outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}</style> </head><body> There were consistency errors detected while updating the MAC address CSV on the <b>" & _ 
				Building & "</b> network. Please make sure that the URL """ & OUIURL & """ is accessible from the server at that location. Updating has been delayed for another " & OUIDaysToUpdate &  " days"
			SendMail RptToEmail, "PingScan - OUI CSV Update Failure"
			outputl = ""
		else 'Success, update CSV
			WriteUTF strCurDir & "\oui.csv", WPData
		end if
	end if
	
	WriteIni strCurDir & "\psapp.ini", "MACCSV", "OUIUpdateAfter", format(Date() + OUIDaysToUpdate, "YYYYMMDD")
end if

'msgbox "This is the last stop"

'Make Powershell Script and run it
MakePSScript
Set objShell = Wscript.CreateObject("Wscript.Shell") 
objShell.Run "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe .\Start-PingScan.ps1", SHOW_ACTIVE_APP, True ' The script will continue until it is closed.
'objShell.Run "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe """ & strCurDir & "\Start-PingScan.ps1""", SHOW_ACTIVE_APP, True ' The script will continue until it is closed.
filesys.DeleteFile strCurDir & "\Start-PingScan.ps1", force

'Get data from CSV file
CSVdata = getfile(CSVPath)
CSVdata = right(CSVdata,len(CSVdata)-81)
'msgbox """" & left(CSVdata,50) & """"

Set adoconn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & DBLocation & ";" & _
			"Database=" & PSSchema & "; User=" & DBUser & "; Password=" & DBPass & ";"

Check_Down_Hosts 'Check first to see if any reported hosts that were up are now down
Get_PingScan_Data 'Read the CSV file and compare if to the database
Mark_Untrusted_Hosts 'Mark any trusted hosts as untrusted after a period of time elapses (set in INI)

'Clean up
filesys.DeleteFile CSVPath, force



Function Check_Down_Hosts()
	str = "Select * from " & PSTbl & " where PingStatus='Y' and LocalBuilding='" & Building & "' order by PingMAC;"
	rs.CursorLocation = 3 'adUseClient
	rs.Open str, adoconn, 3, 3 'OpenType, LockType
	
	if not rs.eof and len(CSVdata) > 10 then 
		rs.MoveFirst
		do while not rs.eof
			if instr(1,CSVdata,rs("PingMAC"),1) < 1 Then
				rs("PingStatus") = "N"
				rs.Update
			End If
			rs.MoveNext
		loop
	end if
	rs.close
End Function

Function Get_PingScan_Data()
	Dim PingIPAdd, PingHost, PingMAC 'The three columns from the scan
	Dim HWType, OIUQTH 'Manufacturer of the network equipment scanned
	Dim TrustedChange 'Temporary holding text for trusted computers that change IP
	
	TrustedChange = ""
	
	Do while len(CSVdata) > 10
		'Get IP Address
		if left(CSVdata,1)="""" then
			PingIPAdd = mid(CSVdata,2,instr(1,CSVdata,""",",1)-2)
			CSVdata = right(CSVdata,len(CSVdata)-instr(1,CSVdata,""",",1)-1)
		else
			PingIPAdd = mid(CSVdata,1,instr(1,CSVdata,",",1)-1)
			CSVdata = right(CSVdata,len(CSVdata)-instr(1,CSVdata,",",1))
		end if
		'msgbox PingIPAdd
		'Get Hostname
		if left(CSVdata,1)="""" then
			PingHost = mid(CSVdata,2,instr(1,CSVdata,""",",1)-2)
			CSVdata = right(CSVdata,len(CSVdata)-instr(1,CSVdata,""",",1)-1)
		else
			PingHost = mid(CSVdata,1,instr(1,CSVdata,",",1)-1)
			CSVdata = right(CSVdata,len(CSVdata)-instr(1,CSVdata,",",1))
		end if
		'msgbox PingHost
		'Get MAC Address
		PingMAC = mid(CSVdata,2,instr(1,CSVdata,vbCrlf,1)-3)
		CSVdata = right(CSVdata,len(CSVdata)-instr(1,CSVdata,vbCrlf,1)-1)
		'PingMAC = replace(PingMAC,"-","")
		'msgbox PingMAC
		'msgbox """" & left(CSVdata,50) & """"
		
		if PingMAC = "unknown" then 'Don't add unknown MAC addresses to the table because that's the primary key
			outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}</style> </head><body> The MAC address was unresolved for the following device that is now active on the <b>" & Building & "</b> network. The device was not added to the database but will continue to trigger this alert for security reasons. Details below:"
			outputl = outputl & vbCrlf & vbCrlf & "<br><br>IP: " & PingIPAdd & vbCrlf & "<br>Name: <b>" & PingHost & vbCrlf & "</b><br>MAC: " & PingMAC & vbCrlf & "<br>Type: Unknown"
			SendMail RptToEmail, "PingScan - Unknown MAC Address for Device"
			outputl = ""
		elseif instr(1,PingExceptions,PingMAC,1) = 0 then 'Ignore any device exceptions
			str = "Select * from " & PSTbl & " where PingMAC='" & PingMAC & "';"
			rs.CursorLocation = 3 'adUseClient
			rs.Open str, adoconn, 3, 3 'OpenType, LockType
			
			if not rs.eof then
				rs.MoveFirst
				if not rs("PingIP") = PingIPAdd and rs("HWTrusted") = "Y" then 'if a trusted computer changes it's IP
					TrustedChange = TrustedChange & vbCrlf & vbCrlf & "<br><br>IP: " & PingIPAdd & " (" & Building & ")<br>Previous IP: " & rs("PingIP") & " (" & rs("LocalBuilding") & ")<br>Name: <b>" & rs("HostName") & vbCrlf & "</b><br>MAC: " & PingMAC & vbCrlf & "<br>Type: " & rs("HWType")
				end if
				rs("PingIP") = PingIPAdd
				if len(rs("HostName") & "") = 0 then
					rs("HostName") = PingHost
				else
					if rs("HostName") = "unknown" and not PingHost = "unknown" then rs("HostName") = PingHost
				end if
				'If hardware type was priviously unknown look it up again
				if rs("HWType") = "Unknown" then
					OIUQTH = instr(1,OIUdata,left(replace(PingMAC,"-",""),6),1)
					if OIUQTH > 0 then
						if mid(OIUdata,OIUQTH+6,2) =  ",""" then
							'msgbox "Quote: " & mid(OIUdata,OIUQTH+8,instr(OIUQTH+8,OIUdata,""",",1)-OIUQTH-8)
							rs("HWType") = mid(OIUdata,OIUQTH+8,instr(OIUQTH+8,OIUdata,""",",1)-OIUQTH-8)
						else
							'msgbox "Comma: " & mid(OIUdata,OIUQTH+7,instr(OIUQTH+7,OIUdata,",",1)-OIUQTH-7)
							rs("HWType") = mid(OIUdata,OIUQTH+7,instr(OIUQTH+7,OIUdata,",",1)-OIUQTH-7)
						end if
					end if
				end if
				if rs("HWTrusted") = "N" and rs("PingStatus") = "N" then 'If an untrusted device comes back on the network
					if EditURL = "" then
						outputl = outputl & vbCrlf & vbCrlf & "<br><br>IP: " & PingIPAdd & vbCrlf & "<br>Name: <b>" & rs("HostName") & vbCrlf & "</b><br>MAC: " & PingMAC & vbCrlf & "<br>Type: " & rs("HWType")
					else
						outputl = outputl & vbCrlf & vbCrlf & "<br><br>IP: " & PingIPAdd & vbCrlf & "<br>Name: <b>" & rs("HostName") & vbCrlf & "</b><br>MAC: " & PingMAC & vbCrlf & "<br>Type: " & rs("HWType") & vbCrlf & "<br><br><a href=""" & EditURL & PingMAC & """>Click here to edit</a>"
					end if
				end if
				rs("PingStatus") = "Y"
				rs("LastDate") = format(date(), "YYYY-MM-DD")
				rs("LastTime") = format(Time, "HH:MM:SS")
				rs("LocalBuilding") = Building
				
				rs.Update
				rs.close
			else
				rs.close
				
				OIUQTH = instr(1,OIUdata,left(replace(PingMAC,"-",""),6),1)
				if OIUQTH > 0 then
					if mid(OIUdata,OIUQTH+6,2) =  ",""" then
						'msgbox "Quote: " & mid(OIUdata,OIUQTH+8,instr(OIUQTH+8,OIUdata,""",",1)-OIUQTH-8)
						HWType = mid(OIUdata,OIUQTH+8,instr(OIUQTH+8,OIUdata,""",",1)-OIUQTH-8)
					else
						'msgbox "Comma: " & mid(OIUdata,OIUQTH+7,instr(OIUQTH+7,OIUdata,",",1)-OIUQTH-7)
						HWType = mid(OIUdata,OIUQTH+7,instr(OIUQTH+7,OIUdata,",",1)-OIUQTH-7)
					end if
				else
					HWType = "Unknown"
				end if
				
				str = "INSERT INTO " & PSTbl & "(PingMAC,PingIP,HostName,HWType,FirstDate,LastDate,LastTime,PingStatus,HWTrusted,LocalBuilding,NearPhone) values('" & PingMAC & "','" & PingIPAdd & "','" & PingHost & "','" & HWType & "','" & format(date(), "YYYY-MM-DD") & "','" & format(date(), "YYYY-MM-DD") & "','" & format(Time, "HH:MM:SS") & "','Y','N','" & Building & "','');"
				adoconn.Execute(str)
				
				if EditURL = "" then
					outputl = outputl & vbCrlf & vbCrlf & "<br><br>IP: " & PingIPAdd & vbCrlf & "<br>Name: <b>" & PingHost & vbCrlf & "</b><br>MAC: " & PingMAC & vbCrlf & "<br>Type: " & HWType
				else
					outputl = outputl & vbCrlf & vbCrlf & "<br><br>IP: " & PingIPAdd & vbCrlf & "<br>Name: <b>" & PingHost & vbCrlf & "</b><br>MAC: " & PingMAC & vbCrlf & "<br>Type: " & HWType & vbCrlf & "<br><br><a href=""" & EditURL & PingMAC & """>Click here to edit</a>"
				end if
			end if
		end if
	Loop
	
	if outputl <> "" then
		outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}</style> </head><body> There are devices(s) that are now active on the <b>" & Building & "</b> network. Details below:" & outputl
		SendMail RptToEmail, "PingScan - Untrusted Devices Found"
		outputl = ""
	end if
	if TrustedChange <> "" then
		outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}</style> </head><body> Trusted device(s) have changed their IP addresses. See info below::" & TrustedChange
		SendMail RptToEmail, "PingScan - Trusted Computer IP Change"
		outputl = ""
	end if
End Function

Function Mark_Untrusted_Hosts()
	str = "Select * from " & PSTbl & " where HWTrusted='Y' and LocalBuilding='" & Building & "' and LastDate < '" & format(date()-DaysBeforeUntrusted, "YYYY-MM-DD") & "' order by PingIP;"
	rs.CursorLocation = 3 'adUseClient
	rs.Open str, adoconn, 3, 3 'OpenType, LockType
	
	if not rs.eof then 
		rs.MoveFirst
		do while not rs.eof
			rs("HWTrusted") = "N"
			outputl = outputl & vbCrlf & vbCrlf & "<br><br>IP: " & rs("PingIP") & vbCrlf & "<br>Name: <b>" & rs("HostName") & vbCrlf & "</b><br>MAC: " & rs("PingMAC") & vbCrlf & "<br>Type: " & rs("HWType")
			rs.Update
			rs.MoveNext
		loop
	end if
	rs.close
	
	if outputl <> "" then
		outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}</style> </head><body> Trusted device(s) have been inactive on the network for more than " & DaysBeforeUntrusted & " days. It will now be marked as untrusted. See info below:" & outputl
		SendMail RptToEmail, "PingScan - Inactive Trusted Device"
		outputl = ""
	end if
End Function

Function MakePSScript()
	Dim PSInfo
	
	'Two Powershell scripts courtesy of:
	'Test-ComputerConnection: Kreloc (https://www.reddit.com/r/PowerShell/comments/3rnrj9/faster_testconnection/)
	'Test-Subnet: Jeff Hicks (https://www.petri.com/building-ping-sweep-tool-powershell-part-4)
	
	PSInfo = "Function Test-ComputerConnection" & vbCrlf & _
	"{" & vbCrlf & _
		"[CmdletBinding()]" & vbCrlf & _
		"param" & vbCrlf & _
		"(" & vbCrlf & _
			"[Parameter(Mandatory=$True," & vbCrlf & _
			"ValueFromPipeline=$True, ValueFromPipelinebyPropertyName=$true)]" & vbCrlf & _
			"[alias(""CN"",""MachineName"",""Device Name"")]" & vbCrlf & _
			"[string]$ComputerName	" & vbCrlf & _
		")" & vbCrlf & _
		"Begin" & vbCrlf & _
		"{" & vbCrlf & _
		"	[int]$timeout = 20" & vbCrlf & _
		"	[switch]$resolve = $true" & vbCrlf & _
		"	[int]$TTL = 128" & vbCrlf & _
		"	[switch]$DontFragment = $false" & vbCrlf & _
		"	[int]$buffersize = 32" & vbCrlf & _
		"	$options = new-object system.net.networkinformation.pingoptions" & vbCrlf & _
		"	$options.TTL = $TTL" & vbCrlf & _
		"	$options.DontFragment = $DontFragment" & vbCrlf & _
		"	$buffer=([system.text.encoding]::ASCII).getbytes(""a""*$buffersize)" & vbCrlf & _	
		"}" & vbCrlf & _
		"Process" & vbCrlf & _
		"{" & vbCrlf & _
		"	$ping = new-object system.net.networkinformation.ping" & vbCrlf & _
		"	try" & vbCrlf & _
		"	{" & vbCrlf & _
		"		$reply = $ping.Send($ComputerName,$timeout,$buffer,$options)" & vbCrlf & _	
		"	}" & vbCrlf & _
		"	catch" & vbCrlf & _
		"	{" & vbCrlf & _
		"		$ErrorMessage = $_.Exception.Message" & vbCrlf & _
		"	}" & vbCrlf & _
		"	if ($reply.status -eq ""Success"")" & vbCrlf & _
		"	{" & vbCrlf & _
		"		return $True" & vbCrlf & _
		"	}" & vbCrlf & _
		"	else" & vbCrlf & _
		"	{" & vbCrlf & _
		"		return $False" & vbCrlf & _
		"	}" & vbCrlf & _
		"}" & vbCrlf & _
	"}" & vbCrlf & _
	"" & vbCrlf & _
	"Function Test-Subnet {" & vbCrlf & _
	 "" & vbCrlf & _
	"[cmdletbinding()]" & vbCrlf & _
	"Param(" & vbCrlf & _
	"[Parameter(Position=0,HelpMessage=""Enter an IPv4 subnet ending in 0."")]" & vbCrlf & _
	"[ValidatePattern(""\d{1,3}\.\d{1,3}\.\d{1,3}\.0"")]" & vbCrlf & _
	"[string]$Subnet= ((Get-NetIPAddress -AddressFamily IPv4).Where({$_.InterfaceAlias -notmatch ""Bluetooth|Loopback""}).IPAddress -replace ""\d{1,3}$"",""0"")," & vbCrlf & _
	 "" & vbCrlf & _
	"[ValidateRange(1,255)]" & vbCrlf & _
	"[int]$Start = 1," & vbCrlf & _
	 "" & vbCrlf & _
	"[ValidateRange(1,255)]" & vbCrlf & _
	"[int]$End = 254," & vbCrlf & _
	 "" & vbCrlf & _
	"[ValidateRange(1,10)]" & vbCrlf & _
	"[Alias(""count"")]" & vbCrlf & _
	"[int]$Ping = 1" & vbCrlf & _
	")" & vbCrlf & _
	 "" & vbCrlf & _
	"Write-Verbose ""Pinging $subnet from $start to $end""" & vbCrlf & _
	"Write-Verbose ""Testing with $ping pings(s)""" & vbCrlf & _
	 "" & vbCrlf & _
	"#a hash table of parameter values to splat to Write-Progress" & vbCrlf & _
	 "$progHash = @{" & vbCrlf & _
	 "Activity = ""Ping Sweep""" & vbCrlf & _
	 "CurrentOperation = ""None""" & vbCrlf & _
	 "Status = ""Pinging IP Address""" & vbCrlf & _
	 "PercentComplete = 0" & vbCrlf & _
	"}" & vbCrlf & _
	 "" & vbCrlf & _
	"#How many addresses need to be pinged?" & vbCrlf & _
	"$count = ($end - $start)+1" & vbCrlf & _
	 "" & vbCrlf & _
	"<#" & vbCrlf & _
	"take the subnet and split it into an array then join the first" & vbCrlf & _
	"3 elements back into a string separated by a period." & vbCrlf & _
	"This will be used to construct an IP address." & vbCrlf & _
	"#>" & vbCrlf & _
	 "" & vbCrlf & _
	"$base = $subnet.split(""."")[0..2] -join "".""" & vbCrlf & _
	 "" & vbCrlf & _
	"#Initialize a counter" & vbCrlf & _
	"$i = 0" & vbCrlf & _
	 "" & vbCrlf & _
	"#get local IP" & vbCrlf & _
	"$local = (Get-NetIPAddress -AddressFamily IPv4).Where({$_.InterfaceAlias -notmatch ""Bluetooth|Loopback""})" & vbCrlf & _
	 "" & vbCrlf & _
	"#loop while the value of $start is <= $end" & vbCrlf & _
	"while ($start -le $end) {" & vbCrlf & _
	"  #increment the counter" & vbCrlf & _
	"  write-Verbose $start" & vbCrlf & _
	"  $i++" & vbCrlf & _
	"  #calculate % processed for Write-Progress" & vbCrlf & _
	"  $progHash.PercentComplete = ($i/$count)*100" & vbCrlf & _
	 "" & vbCrlf & _
	"  #define the IP address to be pinged by using the current value of $start" & vbCrlf & _
	"  $IP = ""$base.$start"" " & vbCrlf & _
	 "" & vbCrlf & _
	"  #Use the value in Write-Progress" & vbCrlf & _
	"  $proghash.currentoperation = $IP" & vbCrlf & _
	"  Write-Progress @proghash" & vbCrlf & _
	  "" & vbCrlf & _
	"  #test the connection" & vbCrlf & _
	"  if (Test-ComputerConnection -ComputerName $IP) {" & vbCrlf & _
	"	#if the IP is not local get the MAC" & vbCrlf & _
	"	if ($IP -ne $Local.IPAddress) {" & vbCrlf & _
	"		#get MAC entry from ARP table" & vbCrlf & _
	"		Try {" & vbCrlf & _
	"			$arp = (arp -a $IP | % {$_.replace($Local.IPAddress,""LocalIP"")} | where {$_ -match $IP}).trim() -split ""\s+""" & vbCrlf & _
	"			$MAC = $arp[1]" & vbCrlf & _
	"		}" & vbCrlf & _
	"		Catch {" & vbCrlf & _
	"			#this should never happen but just in case" & vbCrlf & _
	"			Write-Warning ""Failed to resolve MAC for $IP""" & vbCrlf & _
	"			$MAC = ""unknown""" & vbCrlf & _
	"		}" & vbCrlf & _
	"	}" & vbCrlf & _
	"	else {" & vbCrlf & _
	"		#get local MAC" & vbCrlf & _
	"		$MAC = ($local | Get-NetAdapter).MACAddress" & vbCrlf & _
	"	} " & vbCrlf & _
	"	#attempt to resolve the hostname" & vbCrlf & _
	"	Try {" & vbCrlf & _
	"		$iphost = (Resolve-DNSName -Name $IP -ErrorAction Stop).Namehost" & vbCrlf & _
	"	}" & vbCrlf & _
	"	Catch {" & vbCrlf & _
	"		Write-Verbose ""Failed to resolve host name for $IP""" & vbCrlf & _
	"		#set a value" & vbCrlf & _
	"		$iphost = ""unknown""" & vbCrlf & _
	"	}" & vbCrlf & _
	"	Finally {" & vbCrlf & _
	"		#create a custom object" & vbCrlf & _
	"	   [pscustomobject]@{" & vbCrlf & _
	"	   IPAddress = $IP" & vbCrlf & _
	"	   Hostname = $iphost" & vbCrlf & _
	"	   MAC = $MAC.ToLower()" & vbCrlf & _
	"	   }" & vbCrlf & _
	"	}" & vbCrlf & _
	"  } #if test ping" & vbCrlf & _
	 "" & vbCrlf & _
	"  #increment the value $start by 1" & vbCrlf & _
	"  $start++" & vbCrlf & _
	"} #close while loop" & vbCrlf & _
	 "" & vbCrlf & _
	"} #end function" & vbCrlf & _
	"" & vbCrlf & _
	"test-subnet -subnet " & SubnetDotZero & " -Start " & SubnetStart & " -End " & SubnetEnd & " | export-CSV """ & CSVPath & """"

	WriteFile strCurDir & "\Start-PingScan.ps1", PSInfo
End Function

Function SendMail(TextRcv,TextSubject)
  Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
  Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

  Const cdoAnonymous = 0 'Do not authenticate
  Const cdoBasic = 1 'basic (clear-text) authentication
  Const cdoNTLM = 2 'NTLM

  Set objMessage = CreateObject("CDO.Message") 
  objMessage.Subject = TextSubject 
  objMessage.From = RptFromEmail 
  objMessage.To = TextRcv
  objMessage.HTMLBody = outputl

  '==This section provides the configuration information for the remote SMTP server.

  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

  'Name or IP of Remote SMTP Server
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailSvr

  'Type of authentication, NONE, Basic (Base64 encoded), NTLM
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoAnonymous

  'Server port (typically 25)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

  'Use SSL for the connection (False or True)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

  'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

  objMessage.Configuration.Fields.Update

  '==End remote SMTP server configuration section==

  objMessage.Send
End Function

Function Format(vExpression, sFormat)
  Dim nExpression
  nExpression = sFormat
  
  if isnull(vExpression) = False then
    if instr(1,sFormat,"Y") > 0 or instr(1,sFormat,"M") > 0 or instr(1,sFormat,"D") > 0 or instr(1,sFormat,"H") > 0 or instr(1,sFormat,"S") > 0 then 'Time/Date Format
      vExpression = cdate(vExpression)
	  if instr(1,sFormat,"AM/PM") > 0 and int(hour(vExpression)) > 12 then
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression)-12,2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)-12) '1 character hour
		nExpression = replace(nExpression,"AM/PM","PM") 'If if its afternoon, its PM
	  else
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression),2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)) '1 character hour
		if int(hour(vExpression)) = 12 then nExpression = replace(nExpression,"AM/PM","PM") '12 noon is PM while anything else in this section is AM (fixed 04/19/2019 thanks to our HR Dept.)
		nExpression = replace(nExpression,"AM/PM","AM") 'If its not PM, its AM
	  end if
	  nExpression = replace(nExpression,":MM",":" & right("00" & minute(vExpression),2)) '2 character minute
	  nExpression = replace(nExpression,"SS",right("00" & second(vExpression),2)) '2 character second
	  nExpression = replace(nExpression,"YYYY",year(vExpression)) '4 character year
	  nExpression = replace(nExpression,"YY",right(year(vExpression),2)) '2 character year
	  nExpression = replace(nExpression,"DD",right("00" & day(vExpression),2)) '2 character day
	  nExpression = replace(nExpression,"D",day(vExpression)) '(N)N format day
	  nExpression = replace(nExpression,"MMM",left(MonthName(month(vExpression)),3)) '3 character month name
	  if instr(1,sFormat,"MM") > 0 then
	    nExpression = replace(nExpression,"MM",right("00" & month(vExpression),2)) '2 character month
	  else
	    nExpression = replace(nExpression,"M",month(vExpression)) '(N)N format month
	  end if
    elseif instr(1,sFormat,"N") > 0 then 'Number format
	  nExpression = vExpression
	  if instr(1,sFormat,".") > 0 then 'Decimal format
	    if instr(1,nExpression,".") > 0 then 'Both have decimals
		  do while instr(1,sFormat,".") > instr(1,nExpression,".")
		    nExpression = "0" & nExpression
		  loop
		  if len(nExpression)-instr(1,nExpression,".") >= len(sFormat)-instr(1,sFormat,".") then
		    nExpression = left(nExpression,instr(1,nExpression,".")+len(sFormat)-instr(1,sFormat,"."))
	      else
		    do while len(nExpression)-instr(1,nExpression,".") < len(sFormat)-instr(1,sFormat,".")
			  nExpression = nExpression & "0"
			loop
	      end if
		else
		  nExpression = nExpression & "."
		  do while len(nExpression) < len(sFormat)
			nExpression = nExpression & "0"
		  loop
	    end if
	  else
		do while len(nExpression) < sFormat
		  nExpression = "0" and nExpression
		loop
	  end if
	else
      msgbox "Formating issue on page. Unrecognized format: " & sFormat
	end if
	
	Format = nExpression
  end if
End Function

'Read text file
function GetFile(FileName)
  If FileName<>"" Then
    Dim FS, FileStream
    Set FS = CreateObject("Scripting.FileSystemObject")
      on error resume Next
      Set FileStream = FS.OpenTextFile(FileName)
      GetFile = FileStream.ReadAll
  End If
End Function

'Write string As a text file.
function WriteFile(FileName, Contents)
  Dim OutStream, FS

  on error resume Next
  Set FS = CreateObject("Scripting.FileSystemObject")
    Set OutStream = FS.OpenTextFile(FileName, 2, True)
    OutStream.Write Contents
End Function

'Read UTF file
Function ReadUTF(FileName)
	Dim objStream
	Set objStream = CreateObject("ADODB.Stream")
	
	objStream.CharSet = "utf-8"
	objStream.Open
	objStream.LoadFromFile(FileName)
	ReadUTF = objStream.ReadText()
End Function

'Write UTF file
Function WriteUTF(FileName, Contents)
	Dim objStream
	Set objStream = CreateObject("ADODB.Stream")
	
	objStream.CharSet = "utf-8"
	objStream.Open
	objStream.WriteText Contents
	objStream.SaveToFile FileName, 2
End Function

Function ReadIni( myFilePath, mySection, myKey ) 'Thanks to http://www.robvanderwoude.com
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function

Sub WriteIni( myFilePath, mySection, myKey, myValue ) 'Thanks to http://www.robvanderwoude.com
    ' This subroutine writes a value to an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be written
    ' myValue     [string]  the value to be written (myKey will be
    '                       deleted if myValue is <DELETE_THIS_VALUE>)
    '
    ' Returns:
    ' N/A
    '
    ' CAVEAT:     WriteIni function needs ReadIni function to run
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre, Johan Pol and Rob van der Woude
	Dim WshShell
	Set WshShell = CreateObject("WScript.Shell")


    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni
    Dim strFilePath, strFolderPath, strKey, strLeftString
    Dim strLine, strSection, strTempDir, strTempFile, strValue

    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    strValue    = Trim( myValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
End Sub

'Check to see if database and tables exist
Function CheckForTables()
	Dim CreatePS2DB 'Boolean for DB creation
	CreatePS2DB = False
	
	Set adoconn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & DBLocation & ";" & _
			"User=" & DBUser & "; Password=" & DBPass & ";"
			
	str = "SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '" & PSSchema & "'"
	rs.CursorLocation = 3 'adUseClient
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
	if rs.eof then
		Response = msgbox("The database does not exist. Would you like to create it now? (Make sure the user """ & DBUser & """ has permission to do so)", vbYesNo)
		if Response = vbYes then
			CreatePS2DB = True
		else
			WScript.Quit
		end if
		rs.close
	else
		'msgbox "DB exists"
		rs.close
		
		'Double check to make sure table is also there
		str = "SELECT * FROM information_schema.tables WHERE table_schema = '" & PSSchema & "' AND table_name = '" & PSTbl & "' LIMIT 1;"
		rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
		if rs.eof then
			Response = msgbox("The database exists but the table does not exist. Would you like to create it now?", vbYesNo)
			if Response = vbYes then
				CreatePS2DB = True
			else
				WScript.Quit
			end if
			rs.close
		else
			'msgbox "Table exists"
			rs.close
		end if
	end if
	
	'Create schema and/or table if needed
	if CreatePS2DB = True then
		'Create schema if not there
		str = "CREATE DATABASE IF NOT EXISTS " & PSSchema & ";"
		adoconn.Execute(str)
		
		'Create table
		str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, PingMAC TEXT, PingIP TEXT, HostName TEXT, HWType TEXT, FirstDate DATE,  LastDate DATE, LastTime TEXT, PingStatus TINYTEXT, HWTrusted TINYTEXT, LocalBuilding TEXT, NearPhone TEXT);"
		adoconn.Execute(str)
	end if
	
	Set adoconn = Nothing
	Set rs = Nothing
End Function