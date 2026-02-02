Connect-MgGraph -Scopes Application.Read.All, AppRoleAssignment.ReadWrite.All

### Define these variables here first ##
$ManagedIdentityName = "[Your Managed Identity Name]"
$permissions = "permission1","permission2"

$getPerms = (Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'").approles | Where-Object {$_.Value -in $permissions}
$ManagedIdentity = (Get-MgServicePrincipal -Filter "DisplayName eq '$ManagedIdentityName'")
$GraphID = (Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'").id

foreach ($perm in $getPerms){
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentity.Id -PrincipalId $ManagedIdentity.Id -ResourceId $GraphID -AppRoleId $perm.id
}

Disconnect-MGGraph
