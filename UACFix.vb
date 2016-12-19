'==============Win UAC Fix====================
'
' Scripts by David Rhodes 12/19/16
'
'============================================


Option Explicit
Dim objShell
Dim strUAC, strNewEntry1, strNewEntry2, strSystemHost

MsgBox "This application must be run as Administrator. If not run as Administrator, it will error out."

'Registry input variables
strUAC = "ConsentPromptBehaviorAdmin"

'Registry paths
strSystemHost = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\"

Set objShell = CreateObject("WScript.Shell")

'Changing UAC to level 2
strNewEntry1 = objShell.RegWrite(strSystemHost & strUAC, 5, "REG_DWORD")
