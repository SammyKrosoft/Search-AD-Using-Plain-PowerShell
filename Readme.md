# Intro

This repository is about using PowerShell as a quick way to search your current Active Directory about your Exchange servers without having to install the Exchange admin tools or even the Active Directory module.

By extension, you can search for users, computers, groups, and all sort of objects that Active Directory stores.

# Principle

We're using what we call "type accelerators", which are in PowerShell quick way to refer to .NET libraries, also called  .NET Framework classes.

The examples in this repository make use of the [ADSI] and [ADSISearcher] type accelerators, which are quick references to create references to the System.DirectoryServices.DirectoryEntry class and the System.DirectoryServices.DirectorySearcher class respectively. With these we can access objects in Active Directory from the Schema, the Domain, as well as the Application partitions.

You can get a list of the accelerators on your system by typing the below line:
```powershell
[psobject].Assembly.GetType(“System.Management.Automation.TypeAccelerators”)::get
```

# Other examples

You can find another example of the use of the [ADSI] and the [ADSISearcher] accelerators on my blog posts:
. https://blog.canadasam.ca/2020/01/23/How-To-Check-AD-Schema-Forest-Domain-Versions.html
. https://blog.canadasam.ca/searching-AD-using-powershell-adsisearcher-type-accelerator.html


# References and Thanks to links

Thanks to:

https://ingogegenwarth.wordpress.com/2015/01/14/get-exchange-server-via-ldap/

and:

https://ingogegenwarth.wordpress.com/2015/05/30/troubleshooting-exchange-with-logparserrca-logs/