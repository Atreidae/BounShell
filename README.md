# BounShell

This software is *NOT* ready for production use.. its Beta and you totally run it at your own risk
**The final version will hopefully be delivered by the PS gallery. So none of this faffing about will be required**

As of writing this readme it's now a PS Module.

Either drop it into your PS module folder. or add the folder *containing* the BounShell folder to your PSModulePath variable
 
Simplest way to do the latter is from and admin PS prompt run
**notepad $profile.AllUsersAllHosts**

And add the following line

**$env:PSModulePath = $env:PSModulePath + ";c:\github\\" #The bounshell folder lives in here**

Once PowerShell knows where to find the module.. you can start the tool with "Start-BounShell" from within the PowerShell ISE

I'll add more doco here soon.



## Initial Setup.
ATM this tool is mainly designed for the ISE and thus all of it's configuration takes place there.
Start up the ISE and run **Start-BounShell** this will load the module into memory and add the options to the Add-Ons menu
Now in the ISE Add-Ons menu navigate to BounShell > Settings...

Fill in the details for your tenants as appropriate hit **Save Config** and close the window.

You can now connect to any of the specified tenants by navigating to the Add-On's menu > BounShell > Tenant Name

## Modern Auth Beta (Beta 0.6 and up)
I've implemented a basic form of dealing with Modern Auth. When you attempt to connect to a tenant with the **Modern Auth** flag set BounShell will invoke the modern auth credential window and then copy your username to the clipboard. Simply paste this into the modern auth window using **Ctrl+V** and BounShell will then copy your password into the clipboard.
Paste this again using **Ctrl+V** and BounShell will clean the password out of the clipboard for you.


## Troubleshooting Module installation and loading.
If you receive an error "The term 'start-bounshell' is not recognized as the name of a cmdlet, function, script file, or operable program. " when running "Start-BounShell" or "The specified module 'bounshell' was not loaded because no valid module file was found in any module directory" when you run "Import-Module BounShell"  check your module path is correctly defined in your current session by running $env:PSModulePath and validating that it does indeed contain the folder above the BounShell folder.

You can also check the contents of the all users profile here C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1


## Known Issues
High DPI scaling causes issues with the GUI elements for login credentials. I'm looking into ways to mediate this

Pasting too quickly with Modern Auth causes the username to not be pasted.

Modern Auth will prompt multiple times for login creds when connecting to a tenant. Looking at ways to cache a token

Teams Module throws "out-lineoutput : The method or operation is not implemented." This appears to be a compatability issue between the Teams Module and the ISE in module version 0.9.6, im still investigating this one.

## Fork me!
BounShell is free, open source and licensed under the MIT Licence. Feel free to view the source, fork it, raise issues and submit your improvements via pull requests. You can find on Github:
https://github.com/Atreidae/BounShell/
