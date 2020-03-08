Function Get-SKExchangeServers {
    $RootADConfigPartition = ([ADSI]'LDAP://RootDSE').configurationNamingContext
    $ADConnection = [ADSISearcher]""
    $ADConnection.SearchRoot = "LDAP://$RootADConfigPartition"
    $ADConnection.Filter = "(objectClass=MSExchExchangeServer)"
    #NOTE: we can also use the filter: "(objectCategory=CN=MS-Exch-Exchange-Server,CN=Schema,$RootADConfigPartition)" that field requires a Distinguished Name and is a bit longer ...
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
                            2   {"MBX"}
                            4   {"CAS"}
                            16  {"UM"}
                            20  {"CAS, UM" -split ","}
                            32  {"HUB"}
                            36  {"CAS, HUB" -split ","}
                            38  {"CAS, HUB, MBX" -split ","}
                            54  {"MBX"}
                            64  {"Edge"}
                            16385   {"CAS"}
                            16439   {"CAS, HUB, MBX"  -split ","}
                                }
        }
    }
Return $Coll

}
