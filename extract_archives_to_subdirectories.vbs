' Extract all .7z archives in the current directory to subdirectories named after the archive.
' Assumes 7-zip is installed to the default location.

' Get the password
pword = InputBox("Password? (blank if none)")
If pword <> "" Then
  pword = "-p" & pword & " "
End If

' Instantiate objects we'll need
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshell = WScript.CreateObject("WScript.Shell")

' Get contents of current directory
Set folder = fso.GetFolder(".")

' Iterate over files
For Each file In folder.Files

  fName = file.Name

  fNameMinusExtension = Replace(file.Name, "." & fso.GetExtensionName(file.Path), "")

  ' Only work with .7z files
  If InStr(fName, ".7z") <> 0 Then

    cs = "cmd /c c:\progra~1\7-zip\7z x " & pword & Chr(34) & "-o" & fNameMinusExtension & Chr(34) & " " & Chr(34) & fName & Chr(34)
    wshell.Run cs, 0, True

  End If ' 7z extension check

Next

MsgBox "Done"
