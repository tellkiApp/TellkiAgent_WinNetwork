'###################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								  		      ##
'##																													  		      ##
'## December, 2014																									  		      ##
'##																													  		      ##
'## Version 1.0																										  		      ##
'##																													  		      ##
'## DESCRIPTION: Monitor network interface traffic and errors																	  ##
'##																													  		      ##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "Network.vbs" <HOST> <METRIC_STATE> <USERNAME> <PASSWORD> <DOMAIN>         ##
'##																													  		      ##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "Network.vbs" "10.10.10.1" "1,1,1,0" "user" "pwd" "domain"	  	          ##
'##																													              ##
'## README:	<METRIC_STATE> is generated internally by Tellki and its only used by Tellki default monitors. 						  ##
'##         1 - metric is on ; 0 - metric is off					              												  ##
'## 																												              ##
'## 	    <USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this ##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   ##
'##			pass them to the script.																						      ##
'## 																												              ##
'###################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 5 Then 
	CALL ShowError(3, 0)
End If
'Set Culture - en-us
SetLocale(1033)

'METRIC_ID
Const NetIn = "223:Network traffic in:4"
Const NetOut = "224:Network traffic out:4"
Const NetInErrors = "96:Packets Received Errors:4"
Const NetOutErrors = "213:Packets Outbound Errors:4"


'INPUTS
Dim Host, MetricState, Username, Password, Domain
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
Username = WScript.Arguments(2)
Password = WScript.Arguments(3)
Domain = WScript.Arguments(4)


Dim arrMetrics
arrMetrics = Split(MetricState,",")
Dim objSWbemLocator, objSWbemServices, colItems,colItems2
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim Counter, objItem, objItem2, FullUserName,NetAdapterName
Counter = 0

	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)
	
	if Err.Number = 0 Then
		objSWbemServices.Security_.ImpersonationLevel = 3
		
		Set colItems2 = objSWbemServices.ExecQuery("select name from Win32_NetworkAdapter where NetConnectionStatus = '2' or  NetConnectionStatus = '9'",,16)
		If colItems2.Count <>  0 Then
			For Each objItem2 in colItems2
				' Replace "(" and ")" with "[" and "]"
				NetAdapterName = Replace(objItem2.Name,"(","[")
				NetAdapterName = Replace(NetAdapterName,")","]")
				Set colItems = objSWbemServices.ExecQuery("select name, BytesReceivedPersec,BytesSentPersec,PacketsReceivedErrors,PacketsOutboundErrors from Win32_PerfFormattedData_Tcpip_NetworkInterface  where name = '"& NetAdapterName &"'",,16) 
				If colItems.Count <> 0 Then
				For Each objItem in colItems
					'Received Traffic
					If arrMetrics(0)=1 Then _
					CALL Output(NetIn,FormatNumber((objItem.BytesReceivedPersec/1024)/1024),objItem.Name)
					'Transmitted Traffic
					If arrMetrics(1)=1 Then _
					CALL Output(NetOut,FormatNumber((objItem.BytesSentPersec/1024)/1024),objItem.Name)
					'PacketsReceivedErrors
					If arrMetrics(2)=1 Then _
					CALL Output(NetInErrors,FormatNumber(objItem.PacketsReceivedErrors),objItem.Name)
					'PacketsOutboundErrors
					If arrMetrics(3)=1 Then _
					CALL Output(NetOutErrors,FormatNumber(objItem.PacketsOutboundErrors),objItem.Name)
					Next
				Else
					'If there is no response in WMI query
					CALL ShowError(5, Host)
				End If
			Next
		Else
			'If there is no response in WMI query
			CALL ShowError(5, Host)
		End If
        If Err.number <> 0 Then
           CALL ShowError(5, Host)
           
            Err.Clear
        End If
	End If

If Err Then 
	CALL ShowError(1,0)
Else
	WScript.Quit(0)
End If

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Sub Output(MetricID, MetricValue, MetricObject)
	If MetricObject <> "" Then
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "|" & MetricObject & "|" 
		Else
			CALL ShowError(5, Host) 
		End If
	Else
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|"  & MetricValue & "|" 
		Else
			CALL ShowError(5, Host) 
		End If
	End If
End Sub


