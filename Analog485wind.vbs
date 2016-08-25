' outdoor wind sensor with 4-20ma output (corresponding to 0-32.4 m/s = 0-72.48 mph) is connected to 
' the analog input of a SuperLogics 4014D A/D converter (#8)
' RS-485 output from 4014D is connected to RS-485 input of a Lantronix SDS1100 Serial-to-Network converter
' SDS1100 is connected to LAN via Ethernet Bridge
' PC connected to LAN gets current value of 4014D by reading a virtual serial port
' via Lantronix CPR Manager, set up to talk to SDS1100 as COM8

Option Explicit
Dim MsComm1, objFSO, objFSX, mph, objLogFile, objLogSum, instring, done, OurDate
Dim oldLowCO, oldhiWind, loWind, hiWind, avgWind, flgDate, samp, ttlWind
Dim tSamps, tmph, pmph, ppmph

const path = "C:\ProgramData\Sensors\"
const pSum = "C:\ProgramData\Sensors\Summaries\"
const logfile = "Analog485wind.log"
const logSum = "Analog485wind.log" 
const comEvSend = 1         ' enumeration of comm events
const comEvReceive = 2
const comEvCTS = 3
const comEvDSR = 4
const comEvCD = 5
const comEvRing = 6
const comEvEOF = 7
const comInputModeText = 0       ' enumeration of input mode constants
const comInputModeBinary = 1

ppmph = CSng(0)
hiWind = CSng(0)
loWind = CSng(90)
avgWind = CSng(0)
ttlWind = CSng(0)
samp = CInt(0)
flgDate = Date


Do
	scomOpen ()  '  Open the serial port
	
	On Error Resume Next
	If MSComm1.PortOpen = False Then  ' did it open?
		WScript.Sleep (30000) ' if not, wait 3 seconds
		scomOpen ()	' and try one more time
		WScript.Sleep (500)
		If MSComm1.PortOpen = False Then  ' did it open this time?
			Call LogIt (Now & " - closing down program because can't open MSComm1")
			Exit Do ' No?  Then give up and quit the program
		End If
	End If

	tSamps = 0			' zeroize 10-sample counter
	tmph = 0			' variable to hold total of 10 readings
	pmph = 0			' variable to hold peak wind speed
	Do While tSamps < 10  ' read analog value 10 times then average the total

		MSComm1.Output = Chr(35) & "08" & Chr(13) 		' send '#08' to module 08 to read analog data	 	  
		WScript.Sleep (100) ' give a chance for the module to respond

		mph = CInt(0)
		done = False
		instring = ""		' variable to hold data read from serial port
		Do While Not done   ' read entire contents of serial buffer
			Do While MSComm1.InBufferCount > 0
				instring = instring & MSComm1.Input
			Loop
			If MSComm1.InBufferCount = 0 Then
				done = True
			End If
		Loop

		' convert Analog Data to wind speed in mph
		On Error Resume Next
		If Len(instring) >3 Then		' 4014D outputs 4-1/2 digits
			mph = (CSng(Right(instring,Len(instring) - 4))) * 4.5298  ' 2.025 m/s / ma * 2.23694 mph / m/s
			If mph > pmph Then  ' keep track of peak wind reading 
				pmph = mph
			End If
			tmph = tmph + mph		' add value to 10-sample total
			tSamps = tSamps + 1
		End If

		WScript.Sleep (2500)  ' wait 5 secs for next 1 of 10 reads
	Loop
	
	mph = tmph / 10		' divide total by 10 to get average reading
	
	'------------------------ highs and lows --------------------------
	samp = samp + 1
	ttlWind = ttlWind + mph
	avgWind = ttlWind / samp
	avgWind = avgWind
	If loWind > mph Then  ' keep track of daily high and low temps
		loWind = mph
	End If
	If hiWind < mph Then
		hiWind = mph
	End If
	If pmph > ppmph Then
		ppmph = pmph
	End If

	'-------------- Save the data to log file, if greater than 4mph -------------------------------
	If mph > 2 or pmph > 3 Then
		Call LogIt (Now & "    " & Round(mph,1) & "   " & Round(pmph,2) & "    " & Round(loWind,1) & " - " & Round(hiWind,1) & " - " & Round(ppmph,1) & "    Avg: " & Round(avgWind,2))
	End If
		
	MSComm1.PortOpen = False
	Wscript.DisconnectObject MSComm1  
	Set MSComm1 = Nothing
	WScript.Sleep (2000)
	
	'--------------------- see if we have reached midnight ----------------------
	If flgDate <> Date Then
		' print high and lows for the day
		avgWind = ttlWind / samp
		
		Set objFSX = CreateObject("Scripting.FileSystemObject")
		Set objLogSum = objFSX.OpenTextFile(pSum & logSum, 8, True, 0)
		objLogSum.Writeline (Now & "  " & ACCnt & "  " & ACTtl & "  MaxPress: " & Round(PPMax, 3))
		objLogSum.Close
		Set objLogSum = Nothing
		Set objFSX  = Nothing
		flgDate = Date  ' set comparison date to current date
		ppmph = CSng(0)
		hiWind = CSng(0)
		loWind = CSng(90)
		avgWind = CSng(0)
		ttlWind = CSng(0)
		samp = CInt(0)
	End If	
	'-------------------------------------------------------------------------
	
	Wscript.Sleep (25500)	'// wait 1 minutes to run next scan
Loop 


Sub scomOpen ()
	On Error Resume Next
	
	Set MSComm1 = CreateObject("MSCOMMLib.MSComm")
	MSComm1.CommPort = 8   					' <------------------- COMM port number!
	MSComm1.Settings = "9600,n,8,1"
	MSComm1.InputLen = 0  					' read the entire buffer
	MSComm1.InputMode = comInputModeText   	' Text Only mode
	MSComm1.InBufferCount = 0  				' clear out the receive buffer
	MSComm1.RThreshold = 0 
	MSComm1.PortOpen = True
	WScript.Sleep (2000)
	
	If Err.Number <> 0 Then
		Call LogIt ("   COM " & MSComm1.CommPort & ": not opened!:  " & Err.Number)
		Err.Clear
		WScript.Sleep (10000)
		tryComClose ()
		Exit Sub
	End If
End Sub


Sub LogIt (LogString)
	OurDate = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "_"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objLogFile = objFSO.OpenTextFile(path & OurDate & logfile, 8, True, 0)
	'-------------- Save the data to log file -----------------------------
	objLogFile.Writeline (Now & LogString)
	objLogFile.Close
	 
	Wscript.DisconnectObject objFSO 
	Wscript.DisconnectObject objLogFile
	Set objLogFile = Nothing
	Set objFSO  = Nothing
End Sub


Sub tryComClose ()
	On Error Resume Next
	
	MSComm1.PortOpen = False
	Wscript.DisconnectObject MSComm1  
	Set MSComm1 = Nothing
	WScript.Sleep (500)
	
	If Err.Number <> 0 Then
		Call LogIt (Now & " - can't close open MSComm1 port")
		Err.Clear
		Exit Sub
	End If
End Sub
		