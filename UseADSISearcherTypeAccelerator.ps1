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
$coll | ft


$Exchange2010Count = ($Coll | ? {$_.Version -match "14\."} | measure).count
$Exchange2013Count = ($Coll | ? {$_.Version -match "15\.0"}| measure).count
$Exchange2016Count = ($coll | ? {$_.Version -match "15\.1"}| measure).count

#NOTE: I am using (Object with filter | measure).count instead of (Object with filter).count directly
#because when there is only 1 object matching filter, for some reason the (object with filter).count
#return an empty object ... piping to " measure" which is an alias for Measure-Object, the count
#seems to be always right, even when there is only 1 object in the Object collection matching the filter.

Write-Host "There are $Exchange2010Count Exchange 2010 servers"
Write-Host "There are $Exchange2013Count Exchange 2013 servers"
Write-Host "There are $Exchange2016Count Exchange 2016 servers"
