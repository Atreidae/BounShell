# BounShell

This software is *NOT* ready for production use.. at all

As fo writing this readme it's now a PS Module.

Either drop it into your PS module folder. or add the folder *containing* the BounShell folder to your PSModulePath variable
 
Simplest way to do the latter is from and admin PS prompt run
notepad $profile.AllUsersAllHosts

And add the following line

$env:PSModulePath = $env:PSModulePath + ";c:\github\" #The bounshell folder lives in here

Once PowerShell knows where to find the module.. you can start the tool with "Start-BounShell" from within the PowerShell ISE

I'll add more doco here soon.
