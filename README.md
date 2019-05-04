# BounShell

## Installation
Installation of BounShell is simple.

### Method One
Use NuGet to automatically install it from the PowerShell gallery
`Install-Module -Name BounShell`


### Method two 
Grab the code from GitHub and install it in your modules folder: (C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules or C:\Program Files\WindowsPowerShell\Modules by default)

<a href="https://github.com/Atreidae/BounShell">https://github.com/Atreidae/BounShell</a>

### Install required modules
The current beta (0.6) doesnt automatially install the required modules (some of the code is there... but I'm still working on it) so you can install them manually with the following code

```PowerShell
#Install the modules
Install-Module -name MSOnline, MicrosoftTeams, Microsoft.Online.SharePoint.PowerShell, ExchangeOnlineShell, AADRM
#You should also run update-module to update anything thats already installed
Update-Module -name MSOnline, MicrosoftTeams, Microsoft.Online.SharePoint.PowerShell, ExchangeOnlineShell, AADRM
```

### Developing Code

Wanna try the bleeding edge and not have beta code sitting in your modules folder? Maybe develop your own fork?
Clone the dev branch from github and add the githib repo to your modules path.

```PowerShell
#Clone the repo
git home https://github.com/Atreidae/BounShell.git c:\github\Bounshell
#Open the all users profile
notepad $profile.AllUsersAllHosts
#And add the following line
$env:PSModulePath = $env:PSModulePath + ";c:\github\" 
#Save and restart PowerShell
```

## Initial Setup.
ATM this tool is mainly designed for the ISE and thus all of it's configuration takes place there.
Start up the ISE and run **Start-BounShell** this will load the module into memory and add the options to the Add-Ons menu
Now in the ISE Add-Ons menu navigate to BounShell > Settings...

If you dont want to use the ISE, you can run **Show-BsGuiElements** to show the same menu

Fill in the details for your tenants as appropriate hit **Save Config** and close the window.

You can now connect to any of the specified tenants by navigating to the Add-On's menu > BounShell > Tenant Name

Not using the ISE? Run **Connect-BsO365Tenant** and choose a tenant to connect to.

## Modern Auth Beta (Beta 0.6 and up)
I've implemented a basic form of dealing with Modern Auth. When you attempt to connect to a tenant with the **Modern Auth** flag set BounShell will invoke the modern auth credential window and then copy your username to the clipboard. Simply paste this into the modern auth window using **Ctrl+V** and BounShell will then copy your password into the clipboard.
Paste this again using **Ctrl+V** and BounShell will clean the password out of the clipboard for you.

## Known Issues
High DPI scaling causes issues with the GUI elements for login credentials. I'm looking into ways to mediate this

Pasting too quickly with Modern Auth causes the username to not be pasted.

Modern Auth will prompt multiple times for login creds when connecting to a tenant. Looking at ways to cache a token

## Fork me!
BounShell is free, open source and licensed under the MIT Licence. Feel free to view the source, fork it, raise issues and submit your improvements via pull requests. You can find on Github:
https://github.com/Atreidae/BounShell/
