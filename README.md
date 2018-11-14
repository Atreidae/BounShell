# BounShell

This software is *NOT* ready for production use.. its Beta and you totally run it at your own risk
**The final version will hopefully be delivered by the PS gallery. So none of this faffing about will be required**

As fo writing this readme it's now a PS Module.

Either drop it into your PS module folder. or add the folder *containing* the BounShell folder to your PSModulePath variable
 
Simplest way to do the latter is from and admin PS prompt run
**notepad $profile.AllUsersAllHosts**

And add the following line

**$env:PSModulePath = $env:PSModulePath + ";c:\github\\" #The bounshell folder lives in here**

Once PowerShell knows where to find the module.. you can start the tool with "Start-BounShell" from within the PowerShell ISE

I'll add more doco here soon.



##Initial Setup.
ATM this tool is mainly designed for the ISE and thus all of it's configuration takes place there.
Start up the ISE and run **Start-BounShell** this will load the module into memory and add the options to the Add-Ons menu
Now in the ISE Add-Ons menu navigate to BounShell > Settings...

Fill in the details for your tenants as appropriate hit **Save Config** and close the window.

You can now connect to any of the specified tenants by navigating to the Add-On's menu > BounShell > Tenant Name





## Troubleshooting Module installation and loading.
If you receive an error "The term 'start-bounshell' is not recognized as the name of a cmdlet, function, script file, or operable program. " when running "Start-BounShell" or "The specified module 'bounshell' was not loaded because no valid module file was found in any module directory" when you run "Import-Module BounShell"  check your module path is correctly defined in your current session by running $env:PSModulePath and validating that it does indeed contain the folder above the BounShell folder.

You can also check the contents of the all users profile here C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1
