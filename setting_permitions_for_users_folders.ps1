# AD module
Import-Module ActiveDirectory

#home folders path
$Dir = "D:\users"
 
# OU with users
$Users = Get-ADUser -Filter * -SearchBase "OU=users,DC=domain,DC=local"
 
# Setting permissions
foreach ($User in $Users) {
    $User1 = $User.sAMAccountName
    $Path = $Dir + "\" + $User1
    $objUser = $User.sAMAccountName
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($ObjUser,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.AddAccessRule($Args)
    Set-Acl -path $Path -aclobject $ACL
    }