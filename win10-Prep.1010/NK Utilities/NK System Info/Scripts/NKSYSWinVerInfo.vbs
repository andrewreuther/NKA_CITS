'NK System Info v5.0.0
'Windows Version Information
'Copyright (c) 2012 Nihon Kohden America, Inc.
'=============================================

strComputer = "."
 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
 
 
Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
 
 
 
For Each objOperatingSystem in colOperatingSystems
 
    OSCaption = Trim(Replace(objOperatingSystem.Caption,"Microsoft ",""))
 
    OSCaption = Replace(OSCaption,"Microsoft","")
 
    OSCaption = Replace(OSCaption,"(R)","")
 
    OSCaption = Trim(Replace(OSCaption,",",""))
 
    Echo OSCaption
 
Next