MsgBox "Spooler Started",48

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Set colServices = objWMIService.ExecQuery("Select * from Win32_Service where DisplayName = 'Print Spooler'")

For Each objService in colServices

	WScript.Sleep(5000)
	objService.StartService

Next