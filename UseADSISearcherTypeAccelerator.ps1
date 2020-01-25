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
        Site = @("$($Item.Properties['msExchServerSite'] -Replace '^CN=|,.*$')")
        RolesNb = $Item.Properties['msExchCurrentServerRoles']
        RolesString = Switch ($Item.Properties['msExchCurrentServerRoles']){
                        2 {"MBX"}
                        38 {"CAS, HUB, MBX" -split ","}
                        16439 {"CAS, HUB, MBX"  -split ","}
                            }


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

Write-Host "Number of Exchange 2010 servers: $Exchange2010Count"
Write-Host "Number of Exchange 2013 servers: $Exchange2013Count"
Write-Host "Number of Exchange 2016 servers:           $Exchange2016Count"

<# msExchCurrentServerRoles values:
    The value here is issued from a bitwise value.
    Single role servers:
        MBX=2,
        CAS=4,
        UM=16,
        HT=32,
        Edge=64 
    multirole servers:
        CAS/HT=36,
        CAS/MBX/HT=38,
        CAS/UM=20,
        E2k13 MBX=54,
        E2K13 CAS=16385,
        E2k13 CAS/MBX=16439

#>