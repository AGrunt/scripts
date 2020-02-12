'Do not stop! Keep going!
On error resume next

'Fight!
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'Detect the dwelling of filthy root.
WSEEAD = WshShell.ExpandEnvironmentStrings("%appdata%") 'WshShellExpandEnvironmentStringsAppData XP

'Check files and kill them all.

If fso.fileexists(WSEEAD & "\1C\1CEStart\ibases.v8i") Then
   fso.DeleteFile WSEEAD & "\1C\1CEStart\ibases.v8i", True
   Set fcr = fso.CreateTextFile(WSEEAD & "\1C\1CEStart\ibases.v8i", True)
   fcr.close

   Else

   Set appdata1ccreate = fso.CreateFolder(WSEEAD & "\1C")
   Set appdata1ccreate = fso.CreateFolder(WSEEAD & "\1C\1CEStart\")
   Set fcr = fso.CreateTextFile(WSEEAD & "\1C\1CEStart\ibases.v8i", True)
   fcr.close

End If

'Create new bright brand new list of databases.

If fso.fileexists(WSEEAD & "\1C\1CEStart\1CEStart.cfg") Then
   fso.DeleteFile WSEEAD & "\1C\1CEStart\1CEStart.cfg", True
   Set fcr = fso.CreateTextFile(WSEEAD & "\1C\1CEStart\1CEStart.cfg", True)
   fcr.close

   Else

   Set fcr = fso.CreateTextFile(WSEEAD & "\1C\1CEStart\1CEStart.cfg", True)
   fcr.close

End If

'They are here! Database groups and users.
Set objGroup1 = GetObject("LDAP://CN=groupname,OU=groups,DC=domain,DC=local")
Set objGroup2 = GetObject("LDAP://CN=groupname1,OU=groups,DC=domain,DC=local")

'IT'S HE! Get a _username_
Set objSysInfo  = WScript.CreateObject("ADSystemInfo") 
strUserDN = objSysInfo.userName 
Set objUser = GetObject("LDAP://" & strUserDN) 

'Let us do it right. Set codepage for open.
Set ADODBStream = CreateObject("ADODB.Stream")
ADODBStream.Type = 2
ADODBStream.Charset = "UTF-8"
ADODBStream.Open()


'Chek belonging of user to a group and add block with database details in brand new database list file.
'<-------------
If objGroup1.IsMember(objUser.ADsPath) Then 
   Text =vbCrLf + "[DATABASE NAME]" + vbCrLf + "Connect=Srvr=""server.domain.local"";Ref=""database"";" + vbCrLf + "Folder=/" + vbCrLf + "External=0" + vbCrLf + "ClientConnectionSpeed=Normal" + vbCrLf + "App=Auto" + vbCrLf + "WA=1" + vbCrLf + "Version=8.3"
   ADODBStream.WriteText(Text)
End If

'Next database. Chek belonging of user to a group and add block with database details in brand new database list file.
'<-------------
If objGroup8.IsMember(objUser.ADsPath) Then 
   Text =vbCrLf + "[Казахстан]" + vbCrLf + "Connect=Srvr=""serv2.cl.local"";Ref=""tradekz"";" + vbCrLf + "Folder=/" + vbCrLf + "External=0" + vbCrLf + "ClientConnectionSpeed=Normal" + vbCrLf + "App=Auto" + vbCrLf + "WA=1" + vbCrLf + "Version="
   ADODBStream.WriteText(Text)
End If




'FINISH IT!
ADODBStream.SaveToFile WSEEAD & "\1C\1CEStart\ibases.v8i", 2
ADODBStream.Close() 

'Let us do it right. Set codepage for close. Add the last traits.
Set ADODBStream = CreateObject("ADODB.Stream")
ADODBStream.Type = 2
ADODBStream.Charset = "UTF-8"
ADODBStream.Open()
Text= "UseHWLicenses=0"
ADODBStream.WriteText(Text)
ADODBStream.SaveToFile WSEEAD & "\1C\1CEStart\1CEStart.cfg", 2
ADODBStream.Close()