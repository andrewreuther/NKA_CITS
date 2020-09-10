'NK System Info v5.0.0
'Processor Architecture
'Copyright (c) 2012 Nihon Kohden America, Inc.
'=============================================

strComputer = "."
 
On Error Resume Next
 
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
Set colSettings = objWMIService.ExecQuery ("Select * from Win32_Processor")
 
For Each objComputer in colSettings
 
     If objComputer.Architecture = 0 Then ArchitectureType = "32-bit"
 
     If objComputer.Architecture = 6 Then ArchitectureType = "Intel Itanium"
 
     If objComputer.Architecture = 9 Then ArchitectureType = "64-bit"
 
Next
 
 
 
Echo ArchitectureType