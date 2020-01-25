$RootADConfigPartition = ([ADSI]'LDAP://RootDSE').configurationNamingContext

Write-host "Rood AD config partition" -BackgroundColor Yellow
$RootADConfigPartition

Write-Host "Creating ADSISearcher:" -BackgroundColor Yellow
Write-Host "LDAP://$RootADConfigPartition"

$ADConnection = [ADSISearcher]""
$ADConnection.SearchRoot = "LDAP://$RootADConfigPartition"

$ADConnection.Filter = "(objectCategory=CN=MS-Exch-Exchange-Server,CN=Schema,$RootADConfigPartition)"

write-host "Search root:" -BackgroundColor Yellow
$ADConnection | select -ExpandProperty SearchRoot


$allresults = $ADConnection.FindAll()

$Coll = @()
Foreach ($item in $allresults) {
    $coll+=[pscustomobject]@{
        Name = $Item.Properties['name']
        #ObjectCategory = $item.Properties['objectcategory']
        #ObjectClass = $item.Properties['objectClass']
        #Version1 = $item.Properties['msExchVersion']
        #VErsion2 = $item.Properties['msExchMinAdminVersion']
        Version = $Item.Properties['serialNumber']
        Roles = $Item.Properties['msExchCurrentServerRoles']

    }
}
$coll | fl
