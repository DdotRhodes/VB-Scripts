'==============Win 10 Fix====================
'
' Scripts by David Rhodes 12/11/16
'
'============================================

' Writing to the registry

Option Explicit
Dim objShell
Dim strSystem, strUIPI, strNewEntry1, strNewEntry2, strSystemHost, strUIPIHost

strSystem = "FilterAdministratorToken"
strUIPI = "@"

strSystemHost = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\"

strUIPIHost = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\UIPI\"

Set objShell = CreateObject("WScript.Shell")

strNewEntry1 = objShell.RegWrite(strSystemHost & strSystem, 1, "REG_DWORD")
strNewEntry2 = objShell.RegWrite(strUIPIHost & strUIPI, 1, "REG_DWORD")

' To change a differnt variable like DWord, you would do the following
'strNewEntry = objShell.RegWrite(strHostname & strNewEntry, 1, "REG_DWORD")