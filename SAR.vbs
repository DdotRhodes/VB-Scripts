'============Find and Replace================
'
' Scripts by David Rhodes 12/11/16
'
'============================================

' Find and replace text in file

'Function for file search
Function BrowseForFile()
    With CreateObject("WScript.Shell")
        Dim fso, tempFolder, tempName, path
         
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set tempFolder = fso.GetSpecialFolder(2)
        
        tempName = fso.GetTempName() & ".hta"
        path = "HKCU\Volatile Environment\MsgResp"
        
        With tempFolder.CreateTextFile(tempName)
            .Write "<input type=file name=f>" & _
            "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
            ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
            "close();</script>"
            .Close
            
        End With
        .Run tempFolder & "\" & tempName, 1, True
        BrowseForFile = .RegRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder & "\" & tempName
    End With
End Function


Dim objFSO, objFile, LogFile, Search, FoundText, strOld, strNew
LogFile = BrowseForFile
Const Forreading = 1
Const Forwriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(LogFile, Forreading)

strOld = InputBox("Input word to search for", "Search")

WshShell.AppActivate ("Search")

strNew = InputBox("Input word to replace with", "Replace")

Search = objFile.Readall
objFile.close
FoundText = Replace(Search, strOld, strNew)

Set objFile = objFSO.OpenTextFile(LogFile, Forwriting)
objFile.WriteLine FoundText
objFile.Close