Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c C:\sistema_prestamos\ejecutar_servidor.bat", 0, False
WScript.Sleep 2500
WshShell.Run "https://" & GetLanIp() & ":5000", 1, False

Function GetLanIp()
	Dim service, adapter, items
	Set service = GetObject("winmgmts:\\.\root\cimv2")
	Set items = service.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	For Each adapter In items
		If IsArray(adapter.IPAddress) Then
			For Each ip In adapter.IPAddress
				If InStr(ip, ".") > 0 And Left(ip, 3) <> "127" And Left(ip, 4) <> "169." Then
					GetLanIp = ip
					Exit Function
				End If
			Next
		End If
	Next
	GetLanIp = "127.0.0.1"
End Function