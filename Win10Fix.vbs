'==============Win 10 Fix====================
'
' Scripts by David Rhodes 12/13/16
'
'============================================


Option Explicit
Dim objShell
Dim strSystem, strUIPI, strNewEntry1, strNewEntry2, strSystemHost, strUIPIHost

'Registry input variables
strSystem = "FilterAdministratorToken"
strUIPI = "@"

'Registry paths
strSystemHost = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\"
strUIPIHost = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\UIPI\"

Set objShell = CreateObject("WScript.Shell")

'Changing the registry inputs
strNewEntry1 = objShell.RegWrite(strSystemHost & strSystem, 1, "REG_DWORD")
strNewEntry2 = objShell.RegWrite(strUIPIHost & strUIPI, 1, "REG_DWORD")
