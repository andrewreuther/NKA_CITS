'NK System Info v5.0.0
'Show Network Speed(s) (Active Only)
'Copyright (c) 2012 Nihon Kohden America, Inc.
'=============================================

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Set colFiles = objWMIService.ExecQuery ("SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled=True")

For Each objFile in colFiles

  Echo Round(objFile.Speed / 1000 / 1000, 2) & " Mb/s"

Next