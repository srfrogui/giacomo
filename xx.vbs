Set WshShell = CreateObject("WScript.Shell")
' Executa o .bat com variável de ambiente SKIP_PAUSE=1
WshShell.Run "cmd.exe /c set SKIP_PAUSE=1 && xx.bat", 0, False
