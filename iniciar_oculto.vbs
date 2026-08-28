Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c C:\sistema_prestamos\ejecutar_servidor.bat", 0, False
WScript.Sleep 2500
WshShell.Run "http://127.0.0.1:5000", 1, False