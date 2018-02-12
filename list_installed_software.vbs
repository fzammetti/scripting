' **** NOTE: BITNESS MAY NOT BE ACCURATE! ****

' Keys where installed software is found.
Const HKLM = &H80000002 ' HKEY_LOCAL_MACHINE
Const KEY32BITAPPS = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
Const KEY64BITAPPS = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

' Class to store the data about each app
Class AppDesc
  Public sBitness
  Public sName
  Public sInstDate
  Public sVer
  Public sSize
End Class

' Open outcput file
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.OpenTextFile("Installed Software List.txt", 2, True)

' Open registry
Set objReg = GetObject("winmgmts://./root/default:StdRegProv")

' The array of AppDesc objects and the count of them.
Dim apps()
ReDim apps(0)
cnt = 0

' 32-bit apps.
KEY = KEY32BITAPPS
objReg.EnumKey HKLM, KEY, subKeys32
For Each subKey In subKeys32
  processEntry subKey, "x86"
Next

' 64-bit apps.
KEY = KEY64BITAPPS
objReg.EnumKey HKLM, KEY, subKeys64
For Each subKey In subKeys64
  processEntry subKey, "x64"
Next

' Sort apps.
sortArray

' Write to file and close it.
On Error Resume Next
outFile.Write "Installed software as of " & FormatDateTime(Now) & vbCrLf & vbCrLf
outFile.Write "**** NOTE: BITNESS MAY NOT BE ACCURATE! ****" & vbCrLf & vbCrLf
For i = 0 to UBound(apps)
  sName = apps(i).sName
  sBitness = apps(i).sBitness
  sVer = apps(i).sVer
  sInstDate = apps(i).sInstDate
  sSize = apps(i).sSize
  outFile.Write Trim(sName) & " " & sVer & "(" & sBitness & ", Installed: " & sInstDate & ", Size: " & sSize & ")" & VbCrLf
Next
outFile.Write VbCrLf & "Total: " & UBound(apps) & " (this may not be exact)" & vbCrLf
outFile.Close

' Report finished.
MsgBox "Done"




' Process a subKey entry (aka an installed app).
Function processEntry(subKey, bitness)

  ' Name.
  iRet = objReg.GetStringValue(HKLM, KEY & subKey, "DisplayName", sName)
  If iRet <> 0 Then
    objReg.GetStringValue HKLM, KEY & subKey, "QuietDisplayName", sName
  End If
  If sName = "" Or IsNull(sName) = True Then
    sName = "Unknown"
  End If
  If sName <> "Unknown" Then
    ' Grow array if there's at least one item added already.
    If cnt > 0 Then
      ReDim Preserve apps(UBound(apps) + 1)
    End If
    Set apps(cnt) = new AppDesc
    apps(cnt).sName = sName
    apps(cnt).sBitness = bitness
    ' Install Date.
    objReg.GetStringValue HKLM, KEY & subKey, "InstallDate", sInstDate
    If sInstDate = "" Or IsNull(sInstDate) = True Then
      sInstDate = "Unknown"
    Else
      sInstDate = Mid(sInstDate, 5, 2) & "/" & Mid(sInstDate, 7, 2) & "/" & Mid(sInstDate, 1, 4)
    End If
    apps(cnt).sInstDate = sInstDate
    ' Version.
    objReg.GetDWORDValue HKLM, KEY & subKey, "VersionMajor", sMajVer
    If sMajVer = "" Or IsNull(sMajVer) = True Then
      sMajVer = ""
    End If
    objReg.GetDWORDValue HKLM, KEY & subKey, "VersionMinor", sMinVer
    If sMinVer = "" Or isNull(sMinVer) = True Then
      sMinVer = ""
    End If
    If sMinVer = "" And sMajVer = "" Then
      sVer = ""
    Else
      sVer = "v" & sMajVer & "." & sMinVer & " "
    End If
    apps(cnt).sVer = sVer
    ' Estimated Size.
    objReg.GetDWORDValue HKLM, KEY & subKey, "EstimatedSize", iSize
    If iSize = "" Or IsNull(iSize) = True Then
      sSize = "Unknown"
    Else
      sSize = Round(iSize / 1024, 3) & "Mb)"
    End If
    apps(cnt).sSize = sSize
    ' Bump up counter.
    cnt = cnt + 1
  End If

End Function


' Sorts an array named apps.
Function sortArray()

  For a = UBound(apps) - 1 To 0 Step -1
    For i = 0 to a
      ' Get name of first app, if any.
      n1 = apps(i).sName
      n2 = apps(i + 1).sName
      ' Sort based on the names.
      If n1 > n2 Then
        Set tmp = apps(i + 1)
        Set apps(i + 1) = apps(i)
        Set apps(i) = tmp
      End If
    Next
  Next

End Function
