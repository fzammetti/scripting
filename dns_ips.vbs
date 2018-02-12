' Gets an IP address for a number of domains and writes them out to a file.  This is helpful if your ISP's DNS
' server goes down, this way you have the IPs to use (and if you schedule this to run once a night like I do then
' you're usually good to go when needed).

addressesToGet = Array( _
  "amazon.com", _
  "arstechnica.com", _
  "cnn.com", _
  "digg.com", _
  "engadget.com", _
  "etherient.com", _
  "facebook.com", _
  "google.com", _
  "linkedin.com", _
  "microsoft.com", _
  "nbcnews.com", _
  "netflix.com", _
  "reddit.com", _
  "slashdot.org", _
  "weather.com", _
  "xkcd.com", _
  "zammetti.com"
)

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.OpenTextFile("dns_ips.txt", 2, True)
For i = 0 to UBound(addressesToGet)
  ip = Trim(addressesToGet(i) & " = " & ParseOutIP(DNSLookup(addressesToGet(i))))
  outFile.write(ip & VbCrLf)
Next
outFile.Close


' -----------------------------------------------------------------------------
' Match a domain name to an IP address.
' -----------------------------------------------------------------------------
Function DNSLookup(sAlias)

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1

  Set oShell = CreateObject("WScript.Shell")
  Set oFSO = CreateObject("Scripting.FileSystemObject")

  sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
  sTempFile = sTemp & "\" & oFSO.GetTempName

  oShell.Run "%comspec% /c nslookup " & sAlias & ">" & sTempFile, 0, True

  Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsDefault)
  sResults = fFile.ReadAll
  fFile.Close
  oFSO.DeleteFile (sTempFile)

  DNSLookup = sResults

  Set oShell = Nothing
  Set oFSO = Nothing

End Function


' -----------------------------------------------------------------------------
' Parse out a single IP address from an nslookup result.
' -----------------------------------------------------------------------------
Function ParseOutIP(nsLookupResult)

  If InStr(nsLookupResult, "Addresses") = 0 Then
    ' A single address.
    addStrLoc = InstrRev(nsLookupResult, "Address")
    ipAddr = Mid(nsLookupResult, addStrLoc + 9)
    ParseOutIP = Trim(ipAddr)
  Else
    ' Multiple addresses.  We'll just take the first one.
    addStrLoc = InstrRev(nsLookupResult, "Address")
    ipAddr = Mid(nsLookupResult, addStrLoc + 11, 15)
    ParseOutIP = Trim(ipAddr)
  End If

End Function
