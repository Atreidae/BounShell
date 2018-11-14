# BounShell

This software is *NOT* ready for production use.. at all

As fo writing this readme it's now a PS Module.

Either drop it into your PS module folder. or add the folder *containing* the BounShell folder to your PSModulePath variable
 
Simplest way to do the latter is from and admin PS prompt run
**notepad $profile.AllUsersAllHosts**

And add the following line

**$env:PSModulePath = $env:PSModulePath + ";c:\github\\" #The bounshell folder lives in here**

Once PowerShell knows where to find the module.. you can start the tool with "Start-BounShell" from within the PowerShell ISE

I'll add more doco here soon.


## Troubleshooting Module installation and loading.
If you receive an error "The term 'start-bounshell' is not recognized as the name of a cmdlet, function, script file, or operable program. " when running "Start-BounShell" or "The specified module 'bounshell' was not loaded because no valid module file was found in any module directory" when you run "Import-Module BounShell"  check your module path is correctly defined in your current session by running $env:PSModulePath and validating that it does indeed contain the folder above the BounShell folder.

You can also check the contents of the all users profile here C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1
