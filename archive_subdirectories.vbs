' 7z all subdirectories in the current directory, password-protected.
' Assumes 7-zip is installed in the default location.

' Get the password
pword = InputBox("Password? (blank for unsecured archive)")

' Instantiate objects we'll need
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshell = WScript.CreateObject("WScript.Shell")

' This is the string that will have the output at the end
sout = ""

' Get contents of current directory
Set folder = fso.GetFolder(".")

' Iterate over files
For Each subfolder In folder.SubFolders

  fName = subFolder.Name

  sCmd = "cmd /c c:\progra~1\7-zip\7z "
  If (pword <> "") Then
    sCmd = sCmd & "-p" & pword & " "
  End If
  sCmd = sCmd & "-r a " & Chr(34) & fName & Chr(34) & ".7z .\" & Chr(34) & fName & Chr(34) & "\*.* >> output.txt"
  wshell.Run sCmd, 0, True

Next

MsgBox "Done"
