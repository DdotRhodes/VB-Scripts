'==============Win UAC Fix====================
'
' Scripts by David Rhodes 12/19/16
'
'============================================


Option Explicit
Dim objShell
Dim strUAC, strNewEntry1, strNewEntry2, strSystemHost

'Registry input variables
strUAC = "ConsentPromptBehaviorAdmin"

'Registry paths
strSystemHost = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\"

Set objShell = CreateObject("WScript.Shell")

'Changing the registry inputs
strNewEntry1 = objShell.RegWrite(strSystemHost & strUAC, 5, "REG_DWORD")