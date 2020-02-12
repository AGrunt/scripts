On Error Resume Next

Set wshshell = CreateObject("WScript.Shell")
computername = WshShell.ExpandEnvironmentStrings("%computername%")

If Left(computername,2) <> "VM" then


strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colAccounts = objWMIService.ExecQuery ("Select * From Win32_Group Where LocalAccount = TRUE And SID = 'S-1-5-32-544'")

For Each objAccount in colAccounts
    AdmGrName = objAccount.Name
Next

Set GroupAdm = GetObject("WinNT://" & StrComputer & "/" & AdmGrName & ",group")
Set a = GroupAdm.members()
For Each b In a
    c = b.Name
    c = CStr(c)
    Select Case c                                                                  
                Case "Domain Admins"
		Case "Local Administrators"
                Case "Administrator"
                Case "domain_username"
                Case Else
                                GroupAdm.Remove b.ADsPath
    End Select
Next

end if