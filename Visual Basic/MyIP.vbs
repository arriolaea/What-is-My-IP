' This Visual Basic script (VB Script) pulls local workstation/laptop network information and displays it in a message box.

' Copyright (C) 2021 Ernesto Arriola

' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.

' Explicitly declaring variables.
 Option Explicit
    DIM objHTTP, WshNetwork, strComputer, IPConfigSet, objWMIService, IPConfig, i, message, MyPubIP, BadIP
	DIM strMsg, CheckIPver

On Error Resume Next

' Pulls system public IP.
' Utilizes CloudFlare's icanhazip website which provides the publid IP.
    Set objHTTP = WScript.CreateObject("MSXML2.ServerXmlHttp")
    objHTTP.Open "GET", "http://icanhazip.com", False
    objHTTP.Send
	
	' Returns Public IP or error.
	MyPubIP = objHTTP.ResponseText
	BadIP = "Failed to get IP"

' Error checking on public IP pull.
	If err.number <> 0 then 
		MyPubIP = BadIP
	End If

' Network connector to pull system network info.
    Set WshNetwork = WScript.CreateObject("WScript.Network")

' Variables
	strComputer = "."
	strMsg = ""
	
' WMI connector to pull system service information.
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

' Sets object reference variables.
    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 

For Each IPConfig in IPConfigSet
	If Not IsNull(IPConfig.IPAddress) Then
		For i = LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
			If Not Instr(IPConfig.IPAddress(i), ":") > 0 Then
				strMsg = strMsg & IPConfig.IPAddress(i) & vbNewLine & vbTab & vbTab
			End If
			Next
	End If
Next

' Generates a message box displayed to the user with select feilds
message = Msgbox("Message to user, create your own." & vbNewLine & vbNewLine & vbNewLine & _
	"Domain:"				& vbTab & vbTab & WshNetwork.UserDomain				& vbNewLine & _
	"User Name:"			& vbTab & WshNetwork.UserName						& vbNewLine & _
	"Computer Name:"		& vbTab & WshNetwork.ComputerName		& vbNewLine & vbNewLine & _
	"Public IP Address: "	& vbTab & MyPubIP									& vbNewLine & _
	"Network IP(s): "		& vbTab & strMsg 			& vbNewLine & vbNewLine & vbNewLine & _
	"Servcie Desk: (123) 456-7890", vbInformation, "Message box title displayed for user, write your own.")
