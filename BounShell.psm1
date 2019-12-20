<#
    .SYNOPSIS

    This is a tool to help users manage multiple Office 365 tenants

    .DESCRIPTION

    Created by James Arber. www.UcMadScientist.com
    
    .NOTES

    Version                : 0.7
    Date                   : 24/08/2019 #todo
    Lync Version           : Tested against Skype4B 2015
    Author                 : James Arber
    Header stolen from     : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"
    Special Thanks to      : My Beta Testers. Greig Sheridan, Pat Richard and Justin O'Meara

    :v0.7.0: Public Beta 2
    New XAML based GUI
    New tenant specific settings options for things like Sharepoint URL and region settings
    New Changed settings tracking in GUI
    New Code Signing! (Thanks DigiCert)
    Tried and failed to multithread the GUI process
    New EasterEgg!
    Quick Shoutout To FoxDeploy for getting me back to my XAML GUI roots 
    
    :v0.6.6: Public Beta BugFix
    Added Reconnect flag to CSOnline session to resolve ISE crash
    Added new Gui Window (Not Enabled)

    :v0.6.5: Public Beta BugFix
    Updated CSOnlineSession Timers to delay ISE crash bug

    :v0.6.4: Public Beta Release
    PowerShell package management, (installs, updates and removes old versions of required modules)
    AutoUpdate's moves to package management
    Fixed AzureAD module deleting credentials out of unrelated variables
    Refractoring to support more than 10 tenants (internal testing only)
    Fixed the Autoupdate and Modern Auth checkboxes not saving to the config file
    Improved Modern Auth clipboard behaviour
    Fixed up alot of formatting and readibility issues with ISE Steroids
    
    :v0.6.3: Limited Beta Release
    Enabled AD Security and Compliance Center connection
    Enabled Azure AD connection
    Initial work for config file 0.3 to allow for more than 10 tenants
    Cleaned out a redundant function
    Lots of capitalization fixes
    Log file fixes

    :v0.6.2: Limited Beta Release
    Fixed PowerShell Nuget packaging
    Fixed bug in update code cause by move to SymVer (Thanks Greig)
    Fixed Typo's in configuration grid
    
    :v0.6.1: Limited Beta Release
    Moved to SymVer versioning 
    Pubished to PowerShell gallery
          
    :v0.6: Closed Beta Release
    Enabled Modern Auth Support
    Formating changes
    Broke up alot of my one-liners to make it easier for others to read/ understand the flow
    Updated error messages
    Better code comments
    Fixed an issue with the Compliance Portal code
    Added Module checker and installer based off Andrew Price's "Detect-MicrosoftTeams-Version" http://www.blogabout.cloud/2018/09/240/
    Now Gluten Free
    Finally stopped feature creep

    :v0.5: Closed Beta Release

    Disclaimer: Whilst I take considerable effort to ensure this script is error free and wont harm your enviroment.
    I have no way to test every possible senario it may be used in. I provide these scripts free
    to the Lync and Skype4B community AS IS without any warranty on it's appropriateness for use in
    your environment. I disclaim all implied warranties including,
    without limitation, any implied warranties of merchantability or of fitness for a particular
    purpose. The entire risk arising out of the use or performance of the sample scripts and
    documentation remains with you. In no event shall I be liable for any damages whatsoever
    (including, without limitation, damages for loss of business profits, business interruption,
    loss of business information, or other pecuniary loss) arising out of the use of or inability
    to use the script or documentation.

    Acknowledgements 	
    : Testing and Advice
    Greig Sheridan https://greiginsydney.com/about/ @greiginsydney

    : Auto Update Code
    Pat Richard https://ucunleashed.com @patrichard

    : Proxy Detection
    Michel de Rooij	http://eightwone.com

    : WPF GUI Handler
    Stephen Owen / Fox Deploy 
    https://foxdeploy.com/series/learning-gui-toolmaking-series/

    : MultiThreaded WPF
    Boe Prox / Learn-PowerShell
    https://learn-powershell.net/2012/10/14/powershell-and-wpf-writing-data-to-a-ui-from-a-different-runspace/

    : Everything you wanted to know about hashtables
    Kevin Marquette / PowerShellExplained
    https://powershellexplained.com/2016-11-06-powershell-hashtable-everything-you-wanted-to-know-about/

    : Code Signing Certificate
    DigiCert
    https://www.digicert.com/

    .LINK
    https://www.UcMadScientist.com/BounShell

    .KNOWN ISSUES
    Check https://github.com/Atreidae/BounShell/issues/

    .EXAMPLE
    Loads the Module
    PS C:\> Start-BounShell.ps1

#>

[CmdletBinding(DefaultParametersetName = 'Common')]
param
(
  [switch]$SkipUpdateCheck,
  [String]$ConfigFilePath = $null,
  [String]$LogFileLocation = $null,
  [float]$Tenant = $null

)

#region config
[Net.ServicePointManager]::SecurityProtocol = 'tls12, tls11, tls'
$StartTime                          = Get-Date
$VerbosePreference                  = 'SilentlyContinue' #TODO
[String]$ScriptVersion              = '0.7.0'
[string]$GithubRepo                 = 'BounShell'
[string]$GithubBranch               = 'devel' #todo
[string]$BlogPost                   = 'https://www.UcMadScientist.com/BounShell/' 

#Supported Modules

[String]$TestedTeamsModule          = 'MicrosoftTeams'
[String]$TestedTeamsModuleVer       = '1.0.0'
[String]$TestedExchangeModule       = 'ExchangeOnlineShell' #Using the community version without MFA support. Official version is a clickonce app
[String]$TestedExchangeModuleVer    = '2.0.3.2'
[String]$TestedMSOnlineModule       = 'MsOnline'
[String]$TestedMSOnlineModuleVer    = '1.1.183.17'
[String]$TestedSkype4BOModule       = 'SkypeOnlineConnector'
[String]$TestedSkype4BOModuleVer    = '7.0.0'
[String]$TestedSharepointModule     = 'Microsoft.Online.Sharepoint.PowerShell'
[String]$TestedSharepointModuleVer  = '16.0.8812.1200' 
[String]$TestedAzureADModule        = 'AzureAD' 
[String]$TestedAzureADModuleVer     = '2.0.2.16' 
[String]$TestedAzureADRMModule      = 'AADRM' 
[String]$TestedAzureADRMModuleVer   = '2.13.1.0' 
#[String]$TestedComplianceModule    =   #Not used. Uses New-PSSession
#[String]$TestedComplianceModuleVer =  

#Check to see if paths were specified, Otherwise set defaults
If (!$LogFileLocation) 
{
  $global:LogFileLocation = "$ENV:UserProfile\BounShell.log"
}

If (!$ConfigFilePath) 
{
  $global:ConfigFilePath = "$ENV:UserProfile\BounShell.xml"
}

#endregion config


Function Write-Log
{
  <#
      .SYNOPSIS
      Function to output messages to the console based on their severity and create log files

      .DESCRIPTION
      It's a logger.

      .PARAMETER Message
      The message to write

      .PARAMETER Path
      The location of the logfile.

      .PARAMETER Severity
      Sets the severity of the log message, Higher severities will call Write-Warning or Write-Error

      .PARAMETER Component
      Used to track the module or function that called "Write-Log" 

      .PARAMETER LogOnly
      Forces Write-Log to not display anything to the user

      .EXAMPLE
      Write-Log -Message 'This is a log message' -Severity 3 -component 'Example Component'
      Writes a log file message and displays a warning to the user

      .NOTES
      N/A

      .LINK
      http://www.UcMadScientist.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>
  [CmdletBinding()]
  PARAM
  (
    [String]$Message,
    [String]$Path = $global:LogFileLocation,
    [int]$Severity = 1,
    [string]$Component = 'Default',
    [switch]$LogOnly
  )
  $Date             = Get-Date -Format 'HH:mm:ss'
  $Date2            = Get-Date -Format 'MM-dd-yyyy'
  $MaxLogFileSizeMB = 10
  
  If(Test-Path -Path $Path)
  {
    if(((Get-ChildItem -Path $Path).length/1MB) -gt $MaxLogFileSizeMB) # Check the size of the log file and archive if over the limit.
    {
      $ArchLogfile = $Path.replace('.log', "_$(Get-Date -Format dd-MM-yyy_hh-mm-ss).lo_")
      Rename-Item -Path ren -NewName $Path -Path $ArchLogfile
    }
  }
         
  "$env:ComputerName date=$([char]34)$Date2$([char]34) time=$([char]34)$Date$([char]34) component=$([char]34)$Component$([char]34) type=$([char]34)$Severity$([char]34) Message=$([char]34)$Message$([char]34)"| Out-File -FilePath $Path -Append -NoClobber -Encoding default
  If (!$LogOnly) 
  {
    #If LogOnly is set, we dont want to write anything to the screen as we are capturing data that might look bad onscreen
      
      
    #If the log entry is just Verbose (1), output it to verbose
    if ($Severity -eq 1) 
    {
      "$Date $Message"| Write-Verbose
    }
      
    #If the log entry is just informational (2), output it to write-host
    if ($Severity -eq 2) 
    {
      "Info: $Date $Message"| Write-Host -ForegroundColor Green
    }
    #If the log entry has a severity of 3 assume it's a warning and write it to write-warning
    if ($Severity -eq 3) 
    {
      "$Date $Message"| Write-Warning
    }
    #If the log entry has a severity of 4 or higher, assume it's an error and display an error message (Note, critical errors are caught by throw statements so may not appear here)
    if ($Severity -ge 4) 
    {
      "$Date $Message"| Write-Error
    }
  }
}

Function Get-IEProxy
{
$function = 'Get-IEProxy'
  Write-Log -component $function -Message 'Checking for IE First Run' -severity 1
  if ((Get-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').Property -NotContains 'ProxyEnable')
  {
    $null = New-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -Name ProxyEnable -Value 0
  }
  

  Write-Log -component $function -Message 'Checking for Proxy' -severity 1
  If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0)
  {
    $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
    if ($proxies) 
    {
      if ($proxies -ilike '*=*')
      {
        return $proxies -replace '=', '://' -split (';') | Select-Object -First 1
      }
      
      Else 
      {
        return ('http://{0}' -f $proxies)
      }
    }
    
    Else 
    {
      return $null
    }
  }
  Else 
  {
    return $null
  }
}

Function Get-ScriptUpdate 
{
  $function = 'Get-ScriptUpdate'
  Write-Log -component $function -Message 'Checking for Script Update' -severity 1
  Write-Log -component $function -Message 'Checking for Proxy' -severity 1
  $ProxyURL = Get-IEProxy
  
  If ($ProxyURL)
  
  {
    Write-Log -component $function -Message "Using proxy address $ProxyURL" -severity 1
  }
  
  Else
  {
    Write-Log -component $function -Message 'No proxy setting detected, using direct connection' -severity 1
  }

  Write-Log -component $function -Message "Polling https://raw.githubusercontent.com/atreidae/$GithubRepo/$GithubBranch/version" -severity 1
  $GitHubScriptVersion = Invoke-WebRequest -Uri "https://raw.githubusercontent.com/atreidae/$GithubRepo/$GithubBranch/version" -TimeoutSec 10 -Proxy $ProxyURL -UseBasicParsing
  
  If ($GitHubScriptVersion.Content.length -eq 0) 
  {
    #Empty data, throw an error
    Write-Log -component $function -Message 'Error checking for new version. You can check manually using the url below' -severity 3
    Write-Log -component $function -Message $BlogPost -severity 3 
    Write-Log -component $function -Message 'Pausing for 5 seconds' -severity 1
    Start-Sleep -Seconds 5
  }
  else
  {
    #Process the returned data
    #Symver support!
    [string]$Symver = ($GitHubScriptVersion.Content)
    $splitgitver = $Symver.split('.') 
    $splitver = $ScriptVersion.split('.')
    $needsupdate = $false
    #Check for Major version

    if ([single]$splitgitver[0] -gt [single]$splitver[0])
    {
      $Needupdate = $true
      #New Major Build available, #Prompt user to download
      Write-Log -component $function -Message 'New Major Version Available' -severity 3
      $title = 'Update Available'
      $Message = 'a major update to this script is available, did you want to download it?'
    }

    if (([single]$splitgitver[1] -gt [single]$splitver[1]) -and ([single]$splitgitver[0] -eq [single]$splitver[0]))
    {
      $Needupdate = $true
      #New Major Build available, #Prompt user to download
      Write-Log -component $function -Message 'New Minor Version Available' -severity 3
      $title = 'Update Available'
      $Message = 'a minor update to this script is available, did you want to download it?'
    }

    if (([single]$splitgitver[2] -gt [single]$splitver[2]) -and ([single]$splitgitver[1] -gt [single]$splitver[1]) -and ([single]$splitgitver[0] -eq [single]$splitver[0]))
    {
      $Needupdate = $true
      #New Major Build available, #Prompt user to download
      Write-Log -component $function -Message 'New Bugfix Available' -severity 3
      $title = 'Update Available'
      $Message = 'a bugfix update to this script is available, did you want to download it?'
    }

    If($Needupdate)
    {
      $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
      'Update the installed PowerShell Module'

      $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
      'No thanks.'

      $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

      $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

      switch ($result)
      {
        0 
        {
          #User said yes
          Write-Log -component $function -Message 'User opted to download update' -severity 1
          #start $BlogPost
          Repair-BsInstalledModules -ModuleName 'BounShell' -Operation 'Update'
          Write-Log -component $function -Message 'Exiting Script' -severity 3
          Pause
          exit
        }
        #User said no
        1 
        {
          Write-Log -component $function -Message 'User opted to skip update' -severity 1
        }
      }
    }
    
    #We already have the lastest version
    Else
    {
      Write-Log -component $function -Message 'Script is upto date' -severity 1
    }
  }
}

Function Upgrade-BsConfigFile
{
  <#
      .SYNOPSIS
      Function to upgrade BounShell Config files

      .LINK
      http://www.UcMadScientist.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>
  $function = 'Upgrade-BsConfigFile'
  Write-Log -component $function
  


  Write-Log -component $function -Message "Found Config File Version $($global:Config.ConfigFileVersion)" -severity 2

  #Backup the current config
  Try
  {
    $ShrtDate = Get-Date -Format 'dd-MM-yy'
    $global:Config| Export-Clixml -Path "$ENV:UserProfile\BounShell-backup-$ShrtDate.xml"
    Write-Log -component $function -Message 'Backup File Saved' -severity 2
  }
  Catch 
  {
    Write-Log -component $function -Message 'Error writing Config Backup file' -severity 4
    Write-Log -component $function -Message "Sorry, something went wrong here and I couldn't backup your BounShell config. Please check permissions to create $ENV:UserProfile\BounShell-backup-$ShrtDate.xml" -severity 3
    Throw 'Bad File Operation, Abort Script'
  }
 
  If ($global:Config.ConfigFileVersion -lt '0.2')
  {
    Write-Log -component $function -Message 'Adding Version 0.2 changes' -severity 2
    #Config File Version 0.2 additions
    $global:Config.AutoUpdatesEnabled = $true
    $global:Config.ModernAuthClipboardEnabled = $true
    $global:Config.ModernAuthWarningAccepted = $false
    $global:Config.Tenant1.ConnectToAzureAD = $false
    $global:Config.Tenant1.ConnectToCompliance = $false
    $global:Config.Tenant2.ConnectToAzureAD = $false
    $global:Config.Tenant2.ConnectToCompliance = $false
    $global:Config.Tenant3.ConnectToAzureAD = $false
    $global:Config.Tenant3.ConnectToCompliance = $false
    $global:Config.Tenant4.ConnectToAzureAD = $false
    $global:Config.Tenant4.ConnectToCompliance = $false
    $global:Config.Tenant5.ConnectToAzureAD = $false
    $global:Config.Tenant5.ConnectToCompliance = $false
    $global:Config.Tenant6.ConnectToAzureAD = $false
    $global:Config.Tenant6.ConnectToCompliance = $false
    $global:Config.Tenant7.ConnectToAzureAD = $false
    $global:Config.Tenant7.ConnectToCompliance = $false
    $global:Config.Tenant8.ConnectToAzureAD = $false
    $global:Config.Tenant8.ConnectToCompliance = $false
    $global:Config.Tenant9.ConnectToAzureAD = $false
    $global:Config.Tenant9.ConnectToCompliance = $false
    $global:Config.Tenant10.ConnectToAzureAD = $false
    $global:Config.Tenant10.ConnectToCompliance = $false
    [Float]$global:Config.ConfigFileVersion = '0.2'
  } 
 
  #upgrade the config file to V3
  If ($global:Config.ConfigFileVersion -lt '0.3')
  {
    Write-Log -component $function -Message 'Adding Version 0.3 changes and migrating to hashtables' -severity 2
    #Config File Version 0.3 additions
    #Populate with Values
    
    #declare the new hastable
    $global:Config.Tenants = @{}
    
    #fill it with data
    for ($i = 1; $i -le 10; $i++)
    { $tenant = "Tenant$i"

      Write-host $tenant
      Write-host $global:config.$tenant.displayname

      
      $global:Config.Tenants[$i] = @{}
      $global:Config.Tenants[$i].DisplayName = ($global:config.$tenant.displayname)
      $global:Config.Tenants[$i].SignInAddress = ($global:config.$tenant.SignInAddress)
      $global:Config.Tenants[$i].Credential = ($global:config.$tenant.Credential)
      $global:Config.Tenants[$i].ModernAuth = ($global:config.$tenant.ModernAuth)
      $global:Config.Tenants[$i].ConnectToTeams = ($global:config.$tenant.ConnectToTeams)
      $global:Config.Tenants[$i].ConnectToSkype = ($global:config.$tenant.ConnectToSkype)
      $global:Config.Tenants[$i].ConnectToExchange = ($global:config.$tenant.ConnectToExchange)
      $global:Config.Tenants[$i].ConnectToAzureAD = ($global:config.$tenant.ConnectToAzureAD)
      $global:Config.Tenants[$i].ConnectToCompliance = ($global:config.$tenant.ConnectToCompliance)

    }
    
  } 
  
  

  #Write the XML File
  Try
  {
    $global:Config| Export-Clixml -Path "$ENV:UserProfile\BounShell.xml"
    Write-Log -component $function -Message "Config File Updated to Version $($global:Config.ConfigFileVersion)" -severity 2
  }
  Catch 
  {
    Write-Log -component $function -Message 'Error writing Config file' -severity 3
  }    
}

Function Read-BsConfigFile
{
  $function = 'Read-BsConfigFile'
  Write-Log -component $function
  Write-Log -component $function -Message "Reading Config file $($global:ConfigFilePath)" -severity 2
  If(!(Test-Path $global:ConfigFilePath)) 
  
  {
    #Cant locate test file, Throw error
    Write-Log -component $function -Message 'Could not locate config file!' -severity 3
    Write-Log -component $function -Message 'Error reading Config, Loading Defaults' -severity 3
    Import-BsDefaultConfig
  }
  Else
  {
    #Found the config file
    Write-Log -component $function -Message 'Found Config file in the specified folder' -severity 1
  }

  Write-Log -component $function -Message 'Pulling XML data' -severity 1
  $null = (Remove-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue )
  
  #Try
  #{
    #Load the Config
    $global:Config = @{}
    $global:Config = (Import-Clixml -Path $global:ConfigFilePath)
    Write-Log -component $function -Message 'Config File Read OK' -severity 2
    
    #Check the config file version
    If ($global:Config.ConfigFileVersion -lt 0.3)
    {
      Write-Log -component $function -Message 'Old Config File Detected. Upgrading config file' -severity 3
      Upgrade-BsConfigFile
    }
    



    #Config file is good.



    #Update the Gui options if we are loaded in the ISE
    If($PSISE) 
    {
      #Update the PS ISE Addon Menu
      Update-BsAddonMenu
    }
    
    #Populate with Values
    $null =  $Global:grid_Tenants.Rows.Clear()
    $null =  $Global:grid_Tenants.Rows.Add('1',$global:Config.Tenant1.DisplayName,$global:Config.Tenant1.SignInAddress,'****',$global:Config.Tenant1.ModernAuth,$global:Config.Tenant1.ConnectToTeams,$global:Config.Tenant1.ConnectToSkype,$global:Config.Tenant1.ConnectToExchange,$global:Config.Tenant1.ConnectToAzureAD,$global:Config.Tenant1.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('2',$global:Config.Tenant2.DisplayName,$global:Config.Tenant2.SignInAddress,'****',$global:Config.Tenant2.ModernAuth,$global:Config.Tenant2.ConnectToTeams,$global:Config.Tenant2.ConnectToSkype,$global:Config.Tenant2.ConnectToExchange,$global:Config.Tenant2.ConnectToAzureAD,$global:Config.Tenant2.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('3',$global:Config.Tenant3.DisplayName,$global:Config.Tenant3.SignInAddress,'****',$global:Config.Tenant3.ModernAuth,$global:Config.Tenant3.ConnectToTeams,$global:Config.Tenant3.ConnectToSkype,$global:Config.Tenant3.ConnectToExchange,$global:Config.Tenant3.ConnectToAzureAD,$global:Config.Tenant3.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('4',$global:Config.Tenant4.DisplayName,$global:Config.Tenant4.SignInAddress,'****',$global:Config.Tenant4.ModernAuth,$global:Config.Tenant4.ConnectToTeams,$global:Config.Tenant4.ConnectToSkype,$global:Config.Tenant4.ConnectToExchange,$global:Config.Tenant4.ConnectToAzureAD,$global:Config.Tenant4.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('5',$global:Config.Tenant5.DisplayName,$global:Config.Tenant5.SignInAddress,'****',$global:Config.Tenant5.ModernAuth,$global:Config.Tenant5.ConnectToTeams,$global:Config.Tenant5.ConnectToSkype,$global:Config.Tenant5.ConnectToExchange,$global:Config.Tenant5.ConnectToAzureAD,$global:Config.Tenant5.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('6',$global:Config.Tenant6.DisplayName,$global:Config.Tenant6.SignInAddress,'****',$global:Config.Tenant6.ModernAuth,$global:Config.Tenant6.ConnectToTeams,$global:Config.Tenant6.ConnectToSkype,$global:Config.Tenant6.ConnectToExchange,$global:Config.Tenant6.ConnectToAzureAD,$global:Config.Tenant6.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('7',$global:Config.Tenant7.DisplayName,$global:Config.Tenant7.SignInAddress,'****',$global:Config.Tenant7.ModernAuth,$global:Config.Tenant7.ConnectToTeams,$global:Config.Tenant7.ConnectToSkype,$global:Config.Tenant7.ConnectToExchange,$global:Config.Tenant7.ConnectToAzureAD,$global:Config.Tenant7.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('8',$global:Config.Tenant8.DisplayName,$global:Config.Tenant8.SignInAddress,'****',$global:Config.Tenant8.ModernAuth,$global:Config.Tenant8.ConnectToTeams,$global:Config.Tenant8.ConnectToSkype,$global:Config.Tenant8.ConnectToExchange,$global:Config.Tenant8.ConnectToAzureAD,$global:Config.Tenant8.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('9',$global:Config.Tenant9.DisplayName,$global:Config.Tenant9.SignInAddress,'****',$global:Config.Tenant9.ModernAuth,$global:Config.Tenant9.ConnectToTeams,$global:Config.Tenant9.ConnectToSkype,$global:Config.Tenant9.ConnectToExchange,$global:Config.Tenant9.ConnectToAzureAD,$global:Config.Tenant9.ConnectToCompliance)
    $null =  $Global:grid_Tenants.Rows.Add('10',$global:Config.Tenant10.DisplayName,$global:Config.Tenant10.SignInAddress,'****',$global:Config.Tenant10.ModernAuth,$global:Config.Tenant10.ConnectToTeams,$global:Config.Tenant10.ConnectToSkype,$global:Config.Tenant10.ConnectToExchange,$global:Config.Tenant10.ConnectToAzureAD,$global:Config.Tenant10.ConnectToCompliance)
    
    
    $Global:cbx_AutoUpdates.Checked = $Global:Config.AutoUpdatesEnabled
    $Global:cbx_ClipboardAuth.Checked = $Global:Config.ModernAuthClipboardEnabled
 # }
    
 # Catch
 # {
 #   #For some reason we ran into an issue updating variables, throw an error and revert to defaults
 #   Write-Log -component $function -Message 'Error reading Config or updating GUI, Loading Defaults' -severity 3
 #   Import-BsDefaultConfig
 # }
}

Function Write-BsConfigFile
{
  $function = 'Write-BsConfigFile'
  Write-Log -component $function -Message 'Writing Config file' -severity 2
  
  #Grab items from the GUI and stuff them into something useful

  $global:Config.Tenant1.DisplayName = $Global:grid_Tenants.Rows[0].Cells[1].Value
  $global:Config.Tenant1.SignInAddress = $Global:grid_Tenants.Rows[0].Cells[2].Value
  $global:Config.Tenant1.ModernAuth = $Global:grid_Tenants.Rows[0].Cells[4].Value
  $global:Config.Tenant1.ConnectToTeams = $Global:grid_Tenants.Rows[0].Cells[5].Value
  $global:Config.Tenant1.ConnectToSkype = $Global:grid_Tenants.Rows[0].Cells[6].Value
  $global:Config.Tenant1.ConnectToExchange = $Global:grid_Tenants.Rows[0].Cells[7].Value
  $global:Config.Tenant1.ConnectToAzureAD = $Global:grid_Tenants.Rows[0].Cells[8].Value
  $global:Config.Tenant1.ConnectToCompliance = $Global:grid_Tenants.Rows[0].Cells[9].Value
 

  $global:Config.Tenant2.DisplayName = $Global:grid_Tenants.Rows[1].Cells[1].Value
  $global:Config.Tenant2.SignInAddress = $Global:grid_Tenants.Rows[1].Cells[2].Value
  $global:Config.Tenant2.ModernAuth = $Global:grid_Tenants.Rows[1].Cells[4].Value
  $global:Config.Tenant2.ConnectToTeams = $Global:grid_Tenants.Rows[1].Cells[5].Value
  $global:Config.Tenant2.ConnectToSkype = $Global:grid_Tenants.Rows[1].Cells[6].Value
  $global:Config.Tenant2.ConnectToExchange = $Global:grid_Tenants.Rows[1].Cells[7].Value
  $global:Config.Tenant2.ConnectToAzureAD = $Global:grid_Tenants.Rows[1].Cells[8].Value
  $global:Config.Tenant2.ConnectToCompliance = $Global:grid_Tenants.Rows[1].Cells[9].Value


  $global:Config.Tenant3.DisplayName = $Global:grid_Tenants.Rows[2].Cells[1].Value
  $global:Config.Tenant3.SignInAddress = $Global:grid_Tenants.Rows[2].Cells[2].Value
  $global:Config.Tenant3.ModernAuth = $Global:grid_Tenants.Rows[2].Cells[4].Value
  $global:Config.Tenant3.ConnectToTeams = $Global:grid_Tenants.Rows[2].Cells[5].Value
  $global:Config.Tenant3.ConnectToSkype = $Global:grid_Tenants.Rows[2].Cells[6].Value
  $global:Config.Tenant3.ConnectToExchange = $Global:grid_Tenants.Rows[2].Cells[7].Value
  $global:Config.Tenant3.ConnectToAzureAD = $Global:grid_Tenants.Rows[2].Cells[8].Value
  $global:Config.Tenant3.ConnectToCompliance = $Global:grid_Tenants.Rows[2].Cells[9].Value

 
  $global:Config.Tenant4.DisplayName = $Global:grid_Tenants.Rows[3].Cells[1].Value
  $global:Config.Tenant4.SignInAddress = $Global:grid_Tenants.Rows[3].Cells[2].Value
  $global:Config.Tenant4.ModernAuth = $Global:grid_Tenants.Rows[3].Cells[4].Value
  $global:Config.Tenant4.ConnectToTeams = $Global:grid_Tenants.Rows[3].Cells[5].Value
  $global:Config.Tenant4.ConnectToSkype = $Global:grid_Tenants.Rows[3].Cells[6].Value
  $global:Config.Tenant4.ConnectToExchange = $Global:grid_Tenants.Rows[3].Cells[7].Value
  $global:Config.Tenant4.ConnectToAzureAD = $Global:grid_Tenants.Rows[3].Cells[8].Value
  $global:Config.Tenant4.ConnectToCompliance = $Global:grid_Tenants.Rows[3].Cells[9].Value

 
  $global:Config.Tenant5.DisplayName = $Global:grid_Tenants.Rows[4].Cells[1].Value
  $global:Config.Tenant5.SignInAddress = $Global:grid_Tenants.Rows[4].Cells[2].Value
  $global:Config.Tenant5.ModernAuth = $Global:grid_Tenants.Rows[4].Cells[4].Value
  $global:Config.Tenant5.ConnectToTeams = $Global:grid_Tenants.Rows[4].Cells[5].Value
  $global:Config.Tenant5.ConnectToSkype = $Global:grid_Tenants.Rows[4].Cells[6].Value
  $global:Config.Tenant5.ConnectToExchange = $Global:grid_Tenants.Rows[4].Cells[7].Value
  $global:Config.Tenant5.ConnectToAzureAD = $Global:grid_Tenants.Rows[4].Cells[8].Value
  $global:Config.Tenant5.ConnectToCompliance = $Global:grid_Tenants.Rows[4].Cells[9].Value

 
  $global:Config.Tenant6.DisplayName = $Global:grid_Tenants.Rows[5].Cells[1].Value
  $global:Config.Tenant6.SignInAddress = $Global:grid_Tenants.Rows[5].Cells[2].Value
  $global:Config.Tenant6.ModernAuth = $Global:grid_Tenants.Rows[5].Cells[4].Value
  $global:Config.Tenant6.ConnectToTeams = $Global:grid_Tenants.Rows[5].Cells[5].Value
  $global:Config.Tenant6.ConnectToSkype = $Global:grid_Tenants.Rows[5].Cells[6].Value
  $global:Config.Tenant6.ConnectToExchange = $Global:grid_Tenants.Rows[5].Cells[7].Value
  $global:Config.Tenant6.ConnectToAzureAD = $Global:grid_Tenants.Rows[5].Cells[8].Value
  $global:Config.Tenant6.ConnectToCompliance = $Global:grid_Tenants.Rows[5].Cells[9].Value

 
  $global:Config.Tenant7.DisplayName = $Global:grid_Tenants.Rows[6].Cells[1].Value
  $global:Config.Tenant7.SignInAddress = $Global:grid_Tenants.Rows[6].Cells[2].Value
  $global:Config.Tenant7.ModernAuth = $Global:grid_Tenants.Rows[6].Cells[4].Value
  $global:Config.Tenant7.ConnectToTeams = $Global:grid_Tenants.Rows[6].Cells[5].Value
  $global:Config.Tenant7.ConnectToSkype = $Global:grid_Tenants.Rows[6].Cells[6].Value
  $global:Config.Tenant7.ConnectToExchange = $Global:grid_Tenants.Rows[6].Cells[7].Value
  $global:Config.Tenant7.ConnectToAzureAD = $Global:grid_Tenants.Rows[6].Cells[8].Value
  $global:Config.Tenant7.ConnectToCompliance = $Global:grid_Tenants.Rows[6].Cells[9].Value

 
  $global:Config.Tenant8.DisplayName = $Global:grid_Tenants.Rows[7].Cells[1].Value
  $global:Config.Tenant8.SignInAddress = $Global:grid_Tenants.Rows[7].Cells[2].Value
  $global:Config.Tenant8.ModernAuth = $Global:grid_Tenants.Rows[7].Cells[4].Value
  $global:Config.Tenant8.ConnectToTeams = $Global:grid_Tenants.Rows[7].Cells[5].Value
  $global:Config.Tenant8.ConnectToSkype = $Global:grid_Tenants.Rows[7].Cells[6].Value
  $global:Config.Tenant8.ConnectToExchange = $Global:grid_Tenants.Rows[7].Cells[7].Value
  $global:Config.Tenant8.ConnectToAzureAD = $Global:grid_Tenants.Rows[7].Cells[8].Value
  $global:Config.Tenant8.ConnectToCompliance = $Global:grid_Tenants.Rows[7].Cells[9].Value

 
  $global:Config.Tenant9.DisplayName = $Global:grid_Tenants.Rows[8].Cells[1].Value
  $global:Config.Tenant9.SignInAddress = $Global:grid_Tenants.Rows[8].Cells[2].Value
  $global:Config.Tenant9.ModernAuth = $Global:grid_Tenants.Rows[8].Cells[4].Value
  $global:Config.Tenant9.ConnectToTeams = $Global:grid_Tenants.Rows[8].Cells[5].Value
  $global:Config.Tenant9.ConnectToSkype = $Global:grid_Tenants.Rows[8].Cells[6].Value
  $global:Config.Tenant9.ConnectToExchange = $Global:grid_Tenants.Rows[8].Cells[7].Value
  $global:Config.Tenant9.ConnectToAzureAD = $Global:grid_Tenants.Rows[8].Cells[8].Value
  $global:Config.Tenant9.ConnectToCompliance = $Global:grid_Tenants.Rows[8].Cells[9].Value


  $global:Config.Tenant10.DisplayName = $Global:grid_Tenants.Rows[9].Cells[1].Value
  $global:Config.Tenant10.SignInAddress = $Global:grid_Tenants.Rows[9].Cells[2].Value
  $global:Config.Tenant10.ModernAuth = $Global:grid_Tenants.Rows[9].Cells[4].Value
  $global:Config.Tenant10.ConnectToTeams = $Global:grid_Tenants.Rows[9].Cells[5].Value
  $global:Config.Tenant10.ConnectToSkype = $Global:grid_Tenants.Rows[9].Cells[6].Value
  $global:Config.Tenant10.ConnectToExchange = $Global:grid_Tenants.Rows[9].Cells[7].Value
  $global:Config.Tenant10.ConnectToAzureAD = $Global:grid_Tenants.Rows[9].Cells[8].Value
  $global:Config.Tenant10.ConnectToCompliance = $Global:grid_Tenants.Rows[9].Cells[9].Value

  #Encrypt passwords
  If ($Global:grid_Tenants.Rows[0].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant1.Credential = ($Global:grid_Tenants.Rows[0].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[1].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant2.Credential = ($Global:grid_Tenants.Rows[1].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[2].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant3.Credential = ($Global:grid_Tenants.Rows[2].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[3].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant4.Credential = ($Global:grid_Tenants.Rows[3].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[4].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant5.Credential = ($Global:grid_Tenants.Rows[4].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[5].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant6.Credential = ($Global:grid_Tenants.Rows[5].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[6].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant7.Credential = ($Global:grid_Tenants.Rows[6].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[7].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant8.Credential = ($Global:grid_Tenants.Rows[7].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[8].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant9.Credential = ($Global:grid_Tenants.Rows[8].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }
  If ($Global:grid_Tenants.Rows[9].Cells[3].Value -ne '****') 
  {
    $global:Config.Tenant10.Credential = ($Global:grid_Tenants.Rows[10].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)
  }

  #Clear the password fields
  $Global:grid_Tenants.Rows[0].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[1].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[2].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[3].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[4].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[5].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[6].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[7].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[8].Cells[3].Value = '****'
  $Global:grid_Tenants.Rows[9].Cells[3].Value = '****'

  #Add the Global Settings

  $global:Config.AutoUpdatesEnabled = $Global:cbx_AutoUpdates.Checked
  $global:Config.ModernAuthClipboardEnabled = $Global:cbx_ClipboardAuth.Checked
 

  #Write the XML File
  Try
  {
    $global:Config| Export-Clixml -Path "$ENV:UserProfile\BounShell.xml"
    Write-Log -component $function -Message 'Config File Saved' -severity 2
  }
  Catch 
  {
    Write-Log -component $function -Message 'Error writing Config file' -severity 3
  }
}

Function Import-BsDefaultConfig 
{

  #Set Variables to Defaults
  #Remove and re-create the Config HashTable
  $null = (Remove-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)
  $global:Config = @{}

  
  #Old Gui
  $null =  $Global:grid_Tenants.Rows.Clear()

  #Define the tenants hash table and populate with defaults
  $global:Config.Tenants = @{}

  #99 tenants baby!
  for ($i = 1; $i -lt 100; $i++)
  { 
   
    $global:Config.Tenants[$i] = @{}
    $global:Config.Tenants[$i].DisplayName = 'Undefined'
    $global:Config.Tenants[$i].SignInAddress = 'user1@fabrikam.com'
    $global:Config.Tenants[$i].Credential = '****'
    $global:Config.Tenants[$i].ModernAuth = $false
    $global:Config.Tenants[$i].ConnectToTeams = $false
    $global:Config.Tenants[$i].ConnectToSkype = $false
    $global:Config.Tenants[$i].ConnectToExchange = $false
    $global:Config.Tenants[$i].ConnectToAzureAD = $false
    $global:Config.Tenants[$i].ConnectToCompliance = $false
    $global:Config.Tenants[$i].ConnectToAzureAD = $false
    $global:Config.Tenants[$i].ConnectToCompliance = $false
    
    #Old Gui
    $null =  $Global:grid_Tenants.Rows.Add($i,'Undefined','user@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  }
  
  [Float]$global:Config.ConfigFileVersion = '0.3'
  [string]$global:Config.Description = 'BounShell Configuration file. See UcMadScientist.com/BounShell for more information'
  
  #Config File Version 0.2 additions
  $global:Config.AutoUpdatesEnabled = $true
  $global:Config.ModernAuthClipboardEnabled = $true
  $global:Config.ModernAuthWarningAccepted = $false
}

Function Invoke-BsNewTenantTab 
{
  <#
      .SYNOPSIS
      Function to Open new tab in ISE and call Connect-BsO365Tenant to connect to relevant services

      .DESCRIPTION
        
      .PARAMETER Tenant
      The tenant number to pass to Connect-BsO365Tenant

      .PARAMETER Tabname
      The name for the new tab

      .EXAMPLE
      Invoke-BsNewTenantTab -Tenant 1 -Tabname "Skype4badmin"
      Opens a new tab and connects to the Tenant stored in slot 1

      .NOTES
      N/A

      .LINK
      http://www.UcMadScientist.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>
  param
  (
    [Parameter(Mandatory)] [string]$Tabname,
    [Parameter(Mandatory)] [float]$Tenant
  )
  
  $function = 'Invoke-BsNewTenantTab'
  Write-Log -component $function -Message "Called Invoke-BsNewTenantTab to connect to Tenant $Tenant with a Tabname of $TabName" -severity 1 
  $TabExists = $False
  
  #Check to see if we already have a tab open with that name
  Write-Log -component $function -Message "Checking to see if we already have a tab open called $TabName" -severity 1 
  $OpenTabs = ($PSISE.PowerShellTabs.displayname)
  
  Write-Log -component $function -Message "Existing tabs $OpenTabs" -severity 1  
  
  If ($OpenTabs -Contains $Tabname)
  { 
    #We found an existing tab with that name, prompt the user
    
    $title = "Tab, $Tabname already exists"
    $Message = "There appears to already be a tab called $TabName, Did you want a new tab called 'Copy of $TabName' instead?"
    $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
    'Opens a new Tab'

    $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
    'Aborts the connection'

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

    switch ($result)
    {
      0 
      {
        #User said yes
        Write-Log -component $function -Message 'User opted to open a new tab' -severity 1
        $TabName = ("Copy of $TabName")
      }
      #User said no
      1 
      {
        Write-Log -component $function -Message 'User opted to abort connection' -severity 1
        Write-Log -component $function -Message 'Aborting connection...' -severity 3
        $TabExists = $True
      }
    }
  }   
  
  if(($Tabname -ne 'Undefined') -and ($TabExists -eq $False)) 
  {
    Try
    {
      #kick off a new tab and call it tabname
      Write-Log -component $function -Message 'Opening new ISE tab...' -severity 1 
      $TabNameTab = $PSISE.PowerShellTabs.Add()
      $TabNameTab.DisplayName = $Tabname
      
      #Wait for the tab to wake up
      Write-Log -component $function -Message 'Waiting for tab to become invokable' -severity 1 
      Do 
      {
        Start-Sleep -Milliseconds 100
      }
      While (!$TabNameTab.CanInvoke)
      
      #Kick off the connection
      Write-Log -component $function -Message "Invoking Command: Connect-BsO365Tenant -Tenant $Tenant" -severity 1
      $TabNameTab.Invoke("Connect-BsO365Tenant -Tenant $Tenant")
    }
    
    Catch
    {
      #Something went wrong opening a new tab, probably already a tab with that name open
      Write-Log -component $function -Message 'Failed to open new tab. Is there already a connection open to that tenant?' -severity 3
    }
  }
  
  if($Tabname -eq 'Undefined')
  {
    #Tabname is "undefined", user clicked a tenant thats not confgured yay
    Write-Log -component $function -Message "Sorry, I can't find a config for Tenant $Tenant" -severity 3
  }
}

Function Connect-BsO365Tenant
{
  <#
      .SYNOPSIS
      Connects to relevant Office365 services based on the configuration of the tenant we are passed

      .DESCRIPTION
        
      .PARAMETER Tenant
      The tenant number to connect to

      .EXAMPLE
      Connect-BsO365Tenant -Tenant 1
      Connects to the Tenant stored in slot 1 based on the current context

      .NOTES
      N/A

      .LINK
      http://www.UcMadScientist.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>

  [CmdletBinding()]
  PARAM
  (
    $Tenant
  )
  [string]$function = 'Connect-BsO365Tenant'
  [bool]$ModernAuth = $false
  [bool]$ConnectToTeams = $false
  [bool]$ConnectToSkype = $false
  [bool]$ConnectToExchange = $false
  [bool]$ConnectToSharepoint = $false
  [bool]$ConnectToAzureAD = $false
  [bool]$ConnectToCompliance = $false
  $ModernAuthPassword = ConvertTo-SecureString -String 'Foo' -AsPlainText -Force
  [string]$ModernAuthUsername
  

  #Check to see if we are running in the ISE
  If ($PSISE)
  {  
    #Clean up any stale sessions (we shouldn't have any, but whatever)
    Get-PSSession | Remove-PSSession
  }
  #load the gui stuff for configuration #todo, put a check here and only load it if it's not loaded.
  Import-BsGuiElements

  #Import the Config file so we have data  
  Read-BsConfigFile


  #check to see if a tenant was specified
  If ($Tenant.length -eq 0) 
  {
    Write-Log -Message 'Connect-BsO365Tenant called without a tenant, displaying menu' -severity 1
			
    #Menu code thanks to Greig.

    #Dodgy hack until I refactor the config code #TODO
    $Tenants = @()
    $Tenants += ($global:Config.Tenant1.DisplayName)
    $Tenants += ($global:Config.Tenant2.DisplayName)
    $Tenants += ($global:Config.Tenant3.DisplayName)
    $Tenants += ($global:Config.Tenant4.DisplayName)
    $Tenants += ($global:Config.Tenant5.DisplayName)
    $Tenants += ($global:Config.Tenant6.DisplayName)
    $Tenants += ($global:Config.Tenant7.DisplayName)
    $Tenants += ($global:Config.Tenant8.DisplayName)
    $Tenants += ($global:Config.Tenant9.DisplayName)
    $Tenants += ($global:Config.Tenant10.DisplayName)


    #First figure out the maximum width of the item's name (for the tabular menu):
    $width = 0
    foreach ($Tenant in ($Tenants)) 
    {
      if ($Tenant.Length -gt $width) 
      {
        $width = $Tenant.Length
      }
    }

    #Provide an on-screen menu of tenants for the user to choose from:
    $index = 1
    Write-Host -Object ''
    Write-Host -Object ('ID    '), ('Tenant Name'.Padright($width + 1), ' ')
    foreach ($Tenant in ($Tenants)) 
    {
      Write-Host -Object ($index.ToString()).PadRight(2, ' '), ' | ', ($Tenant.Padright($width + 1), ' ')
      $index++
    }
    $index--	#Undo that last increment
    Write-Host
    Write-Host -Object 'Choose the tenant you wish to use'
    $chosen = Read-Host -Prompt 'Or any other value to quit'
    Write-Log -Message "User input $chosen" -severity 1
    if ($chosen -notmatch '^\d$') 
    {
      Exit
    }
    if ([int]$chosen -lt 0) 
    {
      Exit
    }
    if ([int]$chosen -gt $index) 
    {
      Exit
    }
    $Tenant = $chosen
  }


  Write-Log -component $function -Message "Called to connect to Tenant $Tenant" -severity 1
  #Set the global Modern Auth flag
  if ($global:Config.ModernAuthClipboardEnabled -eq $true)
  {
    [bool]$NoPassword = $true
    Write-Log -component $function -Message "Modern Auth Password intergration enabled" -severity 2
  }
  Else
  {
    [bool]$NoPassword = $False
    Write-Log -component $function -Message "Modern Auth Password intergration disabled" -severity 2
  }

  #change config based on tenant
  #region tenantswitch
  switch ($Tenant)
  {
    1 #Tenant 1
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant1.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant1.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant1.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant1.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant1.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant1.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
     
      If (!$global:Config.Tenant1.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant1.SignInAddress, $global:Config.Tenant1.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant1.Credential
        $ModernAuthUsername = $global:Config.Tenant1.SignInAddress
        
      }
    }
    
    
    2 #Tenant 2
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant2.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant2.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant2.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant2.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant2.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant2.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant2.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant2.SignInAddress, $global:Config.Tenant2.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant2.Credential
        $ModernAuthUsername = $global:Config.Tenant2.SignInAddress
      }
    }
    
    
    3 #Tenant 3
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant3.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant3.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant3.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant3.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant3.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant3.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant3.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant3.SignInAddress, $global:Config.Tenant3.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant3.Credential
        $ModernAuthUsername = $global:Config.Tenant3.SignInAddress
      }
    }
    
    
    4 #Tenant 4
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant4.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant4.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant4.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant4.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant4.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant4.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant4.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant4.SignInAddress, $global:Config.Tenant4.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant4.Credential
        $ModernAuthUsername = $global:Config.Tenant4.SignInAddress
      }
    }
    
    
    5 #Tenant 5
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant5.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant5.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant5.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant5.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant5.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant5.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant5.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant5.SignInAddress, $global:Config.Tenant5.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant5.Credential
        $ModernAuthUsername = $global:Config.Tenant5.SignInAddress
      }
    }
    
    
    6  #Tenant 6
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant6.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant6.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant6.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant6.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant6.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant6.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant6.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant6.SignInAddress, $global:Config.Tenant6.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant6.Credential
        $ModernAuthUsername = $global:Config.Tenant6.SignInAddress
      }
    }
    
    
    7  #Tenant 7
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant7.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant7.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant7.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant7.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant7.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant7.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant7.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant7.SignInAddress, $global:Config.Tenant7.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant7.Credential
        $ModernAuthUsername = $global:Config.Tenant7.SignInAddress
      }
    }
    
    
    8  #Tenant 8
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant8.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant8.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant8.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant8.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant8.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant8.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant8.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant8.SignInAddress, $global:Config.Tenant8.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant8.Credential
        $ModernAuthUsername = $global:Config.Tenant8.SignInAddress
      }
    }
    
    
    9  #Tenant 9
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant9.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant9.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant9.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant9.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant9.ConnectToCompliance
      
      Write-Log -component $function -Message "Loading $($global:Config.Tenant9.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant9.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant9.SignInAddress, $global:Config.Tenant9.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant9.Credential
        $ModernAuthUsername = $global:Config.Tenant9.SignInAddress
      }
    }
    
    
    10 #Tenant 10
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant10.ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant10.ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant10.ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant10.ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant10.ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant10.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant10.ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant10.SignInAddress, $global:Config.Tenant10.Credential)
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Greig!
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant10.Credential
        $ModernAuthUsername = $global:Config.Tenant10.SignInAddress
      }
    }
    11 #Tenant 11 should never happen. This is refactoring for 0.3 config file with array support, it's here for testing.
    {
      #Set Connection flags
      [bool]$ConnectToTeams = $global:Config.Tenant[$Tenant].ConnectToTeams
      [bool]$ConnectToSkype = $global:Config.Tenant[$Tenant].ConnectToSkype
      [bool]$ConnectToExchange = $global:Config.Tenant[$Tenant].ConnectToExchange
      [bool]$ConnectToSharepoint = $false
      [bool]$ConnectToAzureAD = $global:Config.Tenant[$Tenant].ConnectToAzureAD
      [bool]$ConnectToCompliance = $global:Config.Tenant[$Tenant].ConnectToCompliance
      Write-Log -component $function -Message "Loading $($global:Config.Tenant[$Tenant].DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant[$Tenant].ModernAuth) 
      {
        #Not using modern auth
        $global:StoredPsCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($global:Config.Tenant[$Tenant].SignInAddress, $global:Config.Tenant[$Tenant].Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $true
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant[$Tenant].Credential
        $ModernAuthUsername = $global:Config.Tenant[$Tenant].SignInAddress
      }
    }
    
  }

  #endregion tenantswitch
  
  #Check to see if the Modern Auth flag has been set and use the appropriate connection method
  If ($ModernAuth) 
  {
    # We are using Modern Auth, Check to see if the user accepted the warning. If not, prompt them
    If ($global:Config.ModernAuthWarningAccepted -eq $false) 
    { 


      #We should only warn them if the feature is actually on.
      If ($global:Config.ModernAuthClipboardEnabled = $true)
      {
        Write-Log -Message "User hasn't accepted the Modern Auth disclaimer yet. Prompting them." -Severity 1 -Component $function
        Write-Host -Object 'Modern Auth Clipboard integration is currently enabled'
        Write-Host -Object 'Your Username and password will be placed into the clipboard to facilitate login'
        Write-Host -Object 'You can disable this feature in Add-ons > BounShell > Settings'
        Write-Host -Object 'More information on this is available at https://UcMadScientist.com/BounShell/'
        Write-Host -Object 'You will only be shown this warning once.'
        Write-Host -Object '.'
        Write-Host -Object '**************************************************************************************'
        Write-Host -Object '***** Whilst all care is taken, you are still responsible for your own security. *****'
        Write-Host -Object '** I cannot be held liable if your tenant is compromised, explodes or catches fire. **'
        Write-Host -Object "** To remove this warning and enable this feature, type 'I Accept' and press enter. **"
        Write-Host -Object '**************************************************************************************'
        Write-Host -Object 'Press Ctrl+C to abort this connection or'
        $disclaimer = (Read-Host -Prompt "Type 'I Accept' to continue")
        if ($disclaimer -eq 'I Accept')
        {
          Write-Log -Message 'User chose to accept' -Severity 3 -Component $function
          $global:Config.ModernAuthWarningAccepted = $true
          Write-BsConfigFile
        }
        Else
        {
          Write-Log -Message 'User did not accept' -Severity 3 -Component $function
          $global:Config.ModernAuthWarningAccepted = $false
          Throw 'User did not accept modern auth warning, aborting connection'
        }
      }
    }
    If ($global:Config.ModernAuthClipboardEnabled = $true)
    {     
      Write-Host -Object 'Your Username has been copied to the clipboard'
      Write-Host -Object 'You can paste it into the Modern Auth Window using CTRL+V to speed up the sign in process.'
      Write-Host -Object 'Then paste your password using CTRL+V again and sign in.'
      Write-Host -Object 'Upon pasting your password, BounShell will clear the clipboard and overwrite the memory just in case.'
    }
    Else
    {
      Write-Host -Object 'Your Username has been copied to the clipboard'
      Write-Host -Object 'You can paste it into the Modern Auth Window using CTRL+V to speed up the sign in process'
      Write-Host -Object 'You will need to enter your password manually. You can enable password support in Settings'
    }
    #As we are dealing with modern auth we need to convert the password back to an unsecure string do that here
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ModernAuthPassword)
    $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
    
    If ($ConnectToTeams) 
    {
      Try
      {
        # So now we need to kick off a new window that waits for the clipboard events
        Write-Log -Message 'Connecting to Microsoft Teams' -Severity 2 -Component $function
        #Create a script block with the expanded variables
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell converts the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process PowerShell $sb

        #Sleep for a few seconds to let the PowerShell window pop and fill the clipboard.
        Start-Sleep -Seconds 3
        
        #Now we can invoke the session
        Connect-MicrosoftTeams
      } 
      Catch 
      {
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Microsoft Teams' -Severity 3 -Component $function
      }
    }
    
    #Check for the Exchange connection flag
    If ($ConnectToExchange)
    {
      #Flag is set, connect to Exchange
      Try
      {
        #Exchange connection try block
        Write-Log -Message 'Connecting to Exchange Online' -Severity 2 -Component $function
        #So Now we need to kick off a new window that waits for the clipboard events
        #Create a script block with the expanded variables
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell converts the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process PowerShell $sb

        #Sleep for a few seconds to let the PowerShell window pop and fill the clipboard.
        Start-Sleep -Seconds 3

        #Now we invoke the session
        $O365Session = (Connect-ExchangeOnlineShell)
        #Write-Log -Message "Importing Session" -Severity 1 -Component $function
        #$VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        #Import-Module (Import-PSSession -Session $O365Session -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        #$VerbosePreference = "Continue" #Todo. fix for import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch 
      {
        #We had an issue connecting to Exchange
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function
        Write-Log -Message 'Error connecting to Exchange Online' -Severity 3 -Component $function
      }
    }

    #Check for the Skype4B connection flag
    If ($ConnectToSkype) 
    {
      #Flag is set, connect to Skype4B
      Try
      {
        #Skype connection try block
        Write-Log -Message 'Connecting to Skype4B Online' -Severity 2 -Component $function
        #So Now we need to kick off a new window that waits for the clipboard events
        #Create a script block with the expanded variables
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell converts the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process PowerShell $sb

        #Sleep for a few seconds to let the PowerShell window pop and fill the clipboard.
        Start-Sleep -Seconds 3

        #Now we invoke the session
        $S4BOSession = (New-CsOnlineSession)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $S4BOSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Enable-CsOnlineSessionForReconnection #Fix for ISE Lockup!
      } 
      Catch
      {
        #We had an issues connecting to Skype
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Skype4B Online' -Severity 3 -Component $function
      }
    }
    
    
    #Check for the Sharepoint connection flag
    If ($ConnectToSharepoint) 
    {
      #Flag is set, connect to Sharepoint
      Try
      {
        #Sharepoint connection try block
        Write-Log -Message 'Connecting to Sharepoint Online' -Severity 2 -Component $function
        
        #So Now we need to kick off a new window that waits for the clipboard events
        #Create a script block with the expanded variables
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell converts the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process PowerShell $sb

        #Sleep for a few seconds to let the PowerShell window pop and fill the clipboard.
        Start-Sleep -Seconds 3

        #Now we invoke the session

        $SharepointSession = (Connect-SPOService)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $SharepointSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to Sharepoint
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Sharepoint Online' -Severity 3 -Component $function
      }
    }
    
    #Check for the AzureAD connection flag
    If ($ConnectToAzureAD) 
    {
      #Flag is set, connect to AzureAD
      Try
      {
        #Azure AD try block
        Write-Log -Message 'Connecting to Azure AD' -Severity 2 -Component $function
        #So Now we need to kick off a new window that waits for the clipboard events
        #Create a script block with the expanded variables
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell converts the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process PowerShell $sb

        #Sleep for a few seconds to let the PowerShell window pop and fill the clipboard.
        Start-Sleep -Seconds 3

        #Now we invoke the session

        $AADSession = (Connect-AzureAD)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $AADSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to AzureAD
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Azure AD' -Severity 3 -Component $function
      }
    }
    
    #Check for the 365 Compliance Centre flag
    If ($ConnectToCompliance)
    {
      #Flag is set, connect to Compiance Centre
      Try
      {
        Write-Log -Message 'Connecting to Office 365 Security and Compliance Centre' -Severity 2 -Component $function

        #So Now we need to kick off a new window that waits for the clipboard events
        #Create a script block with the expanded variables
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell converts the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process PowerShell $sb

        #Sleep for a few seconds to let the PowerShell window pop and fill the clipboard.
        Start-Sleep -Seconds 3

        #Now we invoke the session
        
        $ComplianceSession = (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -AllowRedirection)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $ComplianceSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to the Compliance Centre
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Office 365 Security Compliance Centre' -Severity 3 -Component $function
      }
    }
  }
  
  



  #region NoModern
  #If the modern auth flag hasn't been set, we can simply connect to the services using secure credentials
  If (!$ModernAuth) 

  {
    #See if we got passed creds
    Write-Log -Message 'Checking for Office365 Credentials' -Severity 1 -Component $function
    If ($global:StoredPsCred -eq $null) 
    {
      #No credentials, prompt user for some
      Write-Log -Message 'No Office365 credentials Found, Prompting user for creds' -Severity 3 -Component $function
      $global:StoredPsCred = Get-Credential
    }
    Else
    {
      #Found creds, continue
      Write-Log -Message "Found Office365 Creds for Username: $($global:StoredPsCred.username)" -Severity 1 -Component $function
    }
    
    #Check for the Exchange connection flag
    If ($ConnectToExchange)
    {
      #Flag is set, connect to Exchange
      Try
      {
        $pscred = $global:StoredPsCred
        #Exchange connection try block
        Write-Log -Message 'Connecting to Exchange Online' -Severity 2 -Component $function
        $O365Session = (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.Office365.com/powershell-liveid/ -Credential $pscred -Authentication Basic -AllowRedirection )
        Write-Log -Message 'Importing Session' -Severity 1 -Component $function
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $O365Session -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch 
      {
        #We had an issue connecting to Exchange
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function
        Write-Log -Message 'Error connecting to Exchange Online' -Severity 3 -Component $function
      }
    }

    #Check for the Skype4B connection flag
    If ($ConnectToSkype) 
    {
      #Flag is set, connect to Skype4B
      Try
      {
        $pscred = $global:StoredPsCred
        #Skype connection try block
        Write-Log -Message 'Connecting to Skype4B Online' -Severity 2 -Component $function
        $S4BOSession = (New-CsOnlineSession -Credential $pscred)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $S4BOSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch
      {
        #We had an issue connecting to Skype
        $ErrorMessage = $_.Exception.Messag
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Skype4B Online' -Severity 3 -Component $function
      }
    }
    
    #Check for the Teams connection flag
    If ($ConnectToTeams)
    {
      #Flag is set, connect to Teams
      Try 
      {
        $pscred = $global:StoredPsCred
        #Teams connection try block
        Write-Log -Message 'Connecting to Microsoft Teams' -Severity 2 -Component $function
        $TeamsSession = (Connect-MicrosoftTeams -Credential $pscred)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        #No need to import the session. Import-Module (Import-PSSession -Session $TeamsSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch
      {
        #We had an issue connecting to Teams
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Microsoft Teams' -Severity 3 -Component $function
      }
    }
    
    #Check for the Sharepoint connection flag
    If ($ConnectToSharepoint) 
    {
      #Flag is set, connect to Sharepoint
      Try
      {
        $pscred = $global:StoredPsCred
        #Sharepoint connection try block
        Write-Log -Message 'Connecting to Sharepoint Online' -Severity 2 -Component $function
        $SharepointSession = (Connect-SPOService -Credential $pscred)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $SharepointSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to Sharepoint
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Sharepoint Online' -Severity 3 -Component $function
      }
    }

    #Check for the AzureAD connection flag
    If ($ConnectToAzureAD) 
    {
      #Flag is set, connect to AzureAD
      Try
      {
        $pscred = $global:StoredPsCred
        #Azure AD try block
        Write-Log -Message 'Connecting to Azure AD' -Severity 2 -Component $function
        $AADSession = (Connect-AzureAD -Credential $pscred)
        #$VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        #Import-Module (Import-PSSession -Session $AADSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        #$VerbosePreference = "Continue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to AzureAD
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Azure AD' -Severity 3 -Component $function
      }
    }
    
    #Check for the 365 Compliance Centre flag
    If ($ConnectToCompliance)
    {
      #Flag is set, connect to Compliance Centre
      Try
      {
        $pscred = $global:StoredPsCred
        Write-Log -Message 'Connecting to Office 365 Compliance Centre' -Severity 2 -Component $function
        $ComplianceSession = (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $pscred -Authentication Basic -AllowRedirection)
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $ComplianceSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to the Compliance Centre
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message $ErrorMessage -Severity 3 -Component $function 
        Write-Log -Message 'Error connecting to Office 365 Compliance Centre' -Severity 3 -Component $function
      }
    }
  }
  #endregion NoModern 
}

Function Update-BsAddonMenu 
{
  #Check to see if we are loaded, if we are then cleanup after ourselves
  if (($PSISE.CurrentPowerShellTab.AddOnsMenu.Submenus).displayname -eq '_BounShell') 
  {
    $null = $PSISE.CurrentPowerShellTab.AddOnsMenu.Submenus.remove($Global:isemenuitem)
  }

  #Create Initial Menu Object
  $null = ($Global:isemenuitem = ($PSISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('_BounShell',$null ,$null)))

  #Add the Settings Button

  $null = ($Global:isemenuitem.Submenus.add('_Settings...', {
        Show-BsGuiElements
  }, $null) )


  #Now add each Tenant

  #Need to put a 'for each' loop in here that adds Tenant 1 through 10
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant1.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant1.DisplayName -Tenant 1
  }, 'Ctrl+Alt+1'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant2.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant2.DisplayName -Tenant 2
  }, 'Ctrl+Alt+2'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant3.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant3.DisplayName -Tenant 3
  }, 'Ctrl+Alt+3'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant4.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant4.DisplayName -Tenant 4
  }, 'Ctrl+Alt+4'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant5.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant5.DisplayName -Tenant 5
  }, 'Ctrl+Alt+5'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant6.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant6.DisplayName -Tenant 6
  }, 'Ctrl+Alt+6'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant7.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant7.DisplayName -Tenant 7
  }, 'Ctrl+Alt+7'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant8.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant8.DisplayName -Tenant 8
  }, 'Ctrl+Alt+8'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant9.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant9.DisplayName -Tenant 9
  }, 'Ctrl+Alt+9'))
  $null = ($Global:isemenuitem.Submenus.add("$($global:Config.Tenant10.DisplayName)",{
        Invoke-BsNewTenantTab -Tabname $global:Config.Tenant10.DisplayName -Tenant 10
  }, 'Ctrl+Alt+0'))
}

Function Import-BsGuiElements 
{
  Write-Log -component $function -Message 'Old Winforms GUI called, this is bad.' -severity 3
  #First we need to import the Functions so they exist for the GUI items
  Import-BsGuiFunctions
  

  #region Gui
  $null = [System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
  $null = [System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
  $Global:SettingsForm = New-Object -TypeName System.Windows.Forms.Form
  [System.Windows.Forms.DataGridView]$Global:grid_Tenants = $null
  [System.Windows.Forms.Button]$Global:btn_CancelConfig = $null
  [System.Windows.Forms.Button]$Global:Btn_ReloadConfig = $null
  [System.Windows.Forms.Button]$Global:Btn_SaveConfig = $null
  [System.Windows.Forms.Button]$Global:Btn_Default = $null
  [System.Windows.Forms.DataGridViewTextBoxColumn]$Global:Tenant_ID = $null
  [System.Windows.Forms.DataGridViewTextBoxColumn]$Global:Tenant_DisplayName = $null
  [System.Windows.Forms.DataGridViewTextBoxColumn]$Global:Tenant_Email = $null
  [System.Windows.Forms.DataGridViewTextBoxColumn]$Global:Tenant_Credentials = $null
  [System.Windows.Forms.DataGridViewCheckBoxColumn]$Global:Tenant_ModernAuth = $null
  [System.Windows.Forms.DataGridViewCheckBoxColumn]$Global:Tenant_Teams = $null
  [System.Windows.Forms.DataGridViewCheckBoxColumn]$Global:Tenant_Skype = $null
  [System.Windows.Forms.DataGridViewCheckBoxColumn]$Global:Tenant_Exchange = $null
  [System.Windows.Forms.CheckBox]$Global:cbx_AutoUpdates = $null

  [System.Windows.Forms.DataGridViewCellStyle]$Global:dataGridViewCellStyle1 = (New-Object -TypeName System.Windows.Forms.DataGridViewCellStyle)
  [System.Windows.Forms.DataGridViewCellStyle]$Global:dataGridViewCellStyle2 = (New-Object -TypeName System.Windows.Forms.DataGridViewCellStyle)
  [System.Windows.Forms.DataGridViewCellStyle]$Global:dataGridViewCellStyle3 = (New-Object -TypeName System.Windows.Forms.DataGridViewCellStyle)
  $Global:btn_CancelConfig = (New-Object -TypeName System.Windows.Forms.Button)
  $Global:Btn_ReloadConfig = (New-Object -TypeName System.Windows.Forms.Button)
  $Global:Btn_SaveConfig = (New-Object -TypeName System.Windows.Forms.Button)
  $Global:cbx_AutoUpdates = (New-Object -TypeName System.Windows.Forms.CheckBox)
  $Global:grid_Tenants = (New-Object -TypeName System.Windows.Forms.DataGridView)
  $Global:Btn_Default = (New-Object -TypeName System.Windows.Forms.Button)
  $Global:Tenant_ID = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
  $Global:Tenant_DisplayName = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
  $Global:Tenant_Email = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
  $Global:Tenant_Credentials = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
  $Global:Tenant_ModernAuth = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
  $Global:Tenant_Teams = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
  $Global:Tenant_Skype = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
  $Global:Tenant_Exchange = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
  $Global:Tenant_AzureAD = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
  $Global:Tenant_Compliance = (New-Object -TypeName System.Windows.Forms.DataGridViewCheckBoxColumn)
  $Global:cbx_ClipboardAuth = (New-Object -TypeName System.Windows.Forms.CheckBox)
  $Global:cliplabel = (New-Object -TypeName System.Windows.Forms.LinkLabel)
  ([System.ComponentModel.ISupportInitialize]$Global:grid_Tenants).BeginInit()
  $Global:SettingsForm.SuspendLayout()
  #
  #btn_CancelConfig
  #
  $Global:btn_CancelConfig.BackColor = [System.Drawing.Color]::White
  $Global:btn_CancelConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:btn_CancelConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif', [System.Single]8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel, ([System.Byte][System.Byte]0)))
  $Global:btn_CancelConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:btn_CancelConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]936, [System.Int32]368))
  $Global:btn_CancelConfig.Name = [System.String]'btn_CancelConfig'
  $Global:btn_CancelConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]94, [System.Int32]23))
  $Global:btn_CancelConfig.TabIndex = [System.Int32]59
  $Global:btn_CancelConfig.Text = [System.String]'Cancel'
  $Global:btn_CancelConfig.UseVisualStyleBackColor = $false
  $Global:btn_CancelConfig.add_Click($Global:btn_CancelConfig_Click)
  #
  #Btn_ReloadConfig
  #
  $Global:Btn_ReloadConfig.BackColor = [System.Drawing.Color]::White
  $Global:Btn_ReloadConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:Btn_ReloadConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif', [System.Single]8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel, ([System.Byte][System.Byte]0)))
  $Global:Btn_ReloadConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:Btn_ReloadConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]704, [System.Int32]368))
  $Global:Btn_ReloadConfig.Name = [System.String]'Btn_ReloadConfig'
  $Global:Btn_ReloadConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110, [System.Int32]23))
  $Global:Btn_ReloadConfig.TabIndex = [System.Int32]58
  $Global:Btn_ReloadConfig.Text = [System.String]'Reload Config'
  $Global:Btn_ReloadConfig.UseVisualStyleBackColor = $true
  $Global:Btn_ReloadConfig.add_Click($Global:Btn_ConfigReload_Click)
  #
  #Btn_SaveConfig
  #
  $Global:Btn_SaveConfig.BackColor = [System.Drawing.Color]::White
  $Global:Btn_SaveConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:Btn_SaveConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif', [System.Single]8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel, ([System.Byte][System.Byte]0)))
  $Global:Btn_SaveConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:Btn_SaveConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]820, [System.Int32]368))
  $Global:Btn_SaveConfig.Name = [System.String]'Btn_SaveConfig'
  $Global:Btn_SaveConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110, [System.Int32]23))
  $Global:Btn_SaveConfig.TabIndex = [System.Int32]57
  $Global:Btn_SaveConfig.Text = [System.String]'Save Config'
  $Global:Btn_SaveConfig.UseVisualStyleBackColor = $false
  $Global:Btn_SaveConfig.add_Click($Global:Btn_SaveConfig_Click)
  #
  #cbx_AutoUpdates
  #
  $Global:cbx_AutoUpdates.AutoSize = $true
  $Global:cbx_AutoUpdates.Checked = $true
  $Global:cbx_AutoUpdates.CheckState = [System.Windows.Forms.CheckState]::Checked
  $Global:cbx_AutoUpdates.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:cbx_AutoUpdates.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27, [System.Int32]370))
  $Global:cbx_AutoUpdates.Name = [System.String]'cbx_AutoUpdates'
  $Global:cbx_AutoUpdates.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]183, [System.Int32]17))
  $Global:cbx_AutoUpdates.TabIndex = [System.Int32]75
  $Global:cbx_AutoUpdates.Text = [System.String]'Automatically Check for Updates'
  $Global:cbx_AutoUpdates.UseVisualStyleBackColor = $true
  $Global:cbx_AutoUpdates.add_CheckedChanged($Global:cbx_NoIntLCD_CheckedChanged_1)
  #
  #grid_Tenants
  #
  $Global:grid_Tenants.AllowUserToAddRows = $false
  $Global:grid_Tenants.AllowUserToDeleteRows = $false
  $Global:dataGridViewCellStyle1.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $Global:dataGridViewCellStyle1.BackColor = [System.Drawing.SystemColors]::Control
  $Global:dataGridViewCellStyle1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif', [System.Single]8.25, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel, ([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle1.ForeColor = [System.Drawing.SystemColors]::WindowText
  $Global:dataGridViewCellStyle1.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle1.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle1.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
  $Global:grid_Tenants.ColumnHeadersDefaultCellStyle = $Global:dataGridViewCellStyle1
  $Global:grid_Tenants.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
  $Global:grid_Tenants.Columns.AddRange($Global:Tenant_ID,$Global:Tenant_DisplayName,$Global:Tenant_Email,$Global:Tenant_Credentials,$Global:Tenant_ModernAuth,$Global:Tenant_Teams,$Global:Tenant_Skype,$Global:Tenant_Exchange,$Global:Tenant_AzureAD,$Global:Tenant_Compliance)
  $Global:dataGridViewCellStyle2.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $Global:dataGridViewCellStyle2.BackColor = [System.Drawing.SystemColors]::Window
  $Global:dataGridViewCellStyle2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif', [System.Single]8.25, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel, ([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle2.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:dataGridViewCellStyle2.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle2.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle2.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
  $Global:grid_Tenants.DefaultCellStyle = $Global:dataGridViewCellStyle2
  $Global:grid_Tenants.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12, [System.Int32]12))
  $Global:grid_Tenants.Name = [System.String]'grid_Tenants'
  $Global:dataGridViewCellStyle3.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $Global:dataGridViewCellStyle3.BackColor = [System.Drawing.SystemColors]::Control
  $Global:dataGridViewCellStyle3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif', [System.Single]8.25, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel, ([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle3.ForeColor = [System.Drawing.SystemColors]::WindowText
  $Global:dataGridViewCellStyle3.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle3.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle3.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
  $Global:grid_Tenants.RowHeadersDefaultCellStyle = $Global:dataGridViewCellStyle3
  $Global:grid_Tenants.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]1020, [System.Int32]336))
  $Global:grid_Tenants.TabIndex = [System.Int32]76
  $Global:grid_Tenants.add_CellContentClick($Global:grid_Tenants_CellContentClick)

  #
  #Btn_Default
  #
  $Global:Btn_Default.BackColor = [System.Drawing.Color]::White
  $Global:Btn_Default.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:Btn_Default.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif', [System.Single]8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel, ([System.Byte][System.Byte]0)))
  $Global:Btn_Default.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:Btn_Default.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]588, [System.Int32]368))
  $Global:Btn_Default.Name = [System.String]'Btn_Default'
  $Global:Btn_Default.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110, [System.Int32]23))
  $Global:Btn_Default.TabIndex = [System.Int32]77
  $Global:Btn_Default.Text = [System.String]'Reset to Default'
  $Global:Btn_Default.UseVisualStyleBackColor = $true
  $Global:Btn_Default.add_Click($Global:Btn_Default_Click)
  #
  #Tenant_ID
  #
  $Global:Tenant_ID.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
  $Global:Tenant_ID.Frozen = $true
  $Global:Tenant_ID.HeaderText = [System.String]'ID'
  $Global:Tenant_ID.Name = [System.String]'Tenant_ID'
  $Global:Tenant_ID.ReadOnly = $true
  $Global:Tenant_ID.Width = [System.Int32]43
  #
  #Tenant_DisplayName
  #
  $Global:Tenant_DisplayName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
  $Global:Tenant_DisplayName.Frozen = $true
  $Global:Tenant_DisplayName.HeaderText = [System.String]'Display Name'
  $Global:Tenant_DisplayName.Name = [System.String]'Tenant_DisplayName'
  $Global:Tenant_DisplayName.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
  $Global:Tenant_DisplayName.Width = [System.Int32]78
  #
  #Tenant_Email
  #
  $Global:Tenant_Email.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
  $Global:Tenant_Email.Frozen = $true
  $Global:Tenant_Email.HeaderText = [System.String]'Sign-In Address'
  $Global:Tenant_Email.Name = [System.String]'Tenant_Email'
  $Global:Tenant_Email.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
  $Global:Tenant_Email.Width = [System.Int32]78
  #
  #Tenant_Credentials
  #
  $Global:Tenant_Credentials.Frozen = $true
  $Global:Tenant_Credentials.HeaderText = [System.String]'Credentials'
  $Global:Tenant_Credentials.Name = [System.String]'Tenant_Credentials'
  #
  #Tenant_ModernAuth
  #
  $Global:Tenant_ModernAuth.HeaderText = [System.String]'Uses Modern Auth?'
  $Global:Tenant_ModernAuth.Name = [System.String]'Tenant_ModernAuth'
  #
  #Tenant_Teams
  #
  $Global:Tenant_Teams.HeaderText = [System.String]'Connect to Teams?'
  $Global:Tenant_Teams.Name = [System.String]'Tenant_Teams'
  #
  #Tenant_Skype
  #
  $Global:Tenant_Skype.HeaderText = [System.String]'Connect to Skype?'
  $Global:Tenant_Skype.Name = [System.String]'Tenant_Skype'
  #
  #Tenant_Exchange
  #
  $Global:Tenant_Exchange.HeaderText = [System.String]'Connect to Exchange?'
  $Global:Tenant_Exchange.Name = [System.String]'Tenant_Exchange'
  #
  #Tenant_AzureAD
  #
  $Global:Tenant_AzureAD.HeaderText = [System.String]'Connect to Azure AD?'
  $Global:Tenant_AzureAD.Name = [System.String]'Tenant_AzureAD'
  #
  #Tenant_Compliance
  #
  $Global:Tenant_Compliance.HeaderText = [System.String]'Connect to Compliance Centre?'
  $Global:Tenant_Compliance.Name = [System.String]'Tenant_Compliance'
  #
  #cbx_ClipboardAuth
  #
  $Global:cbx_ClipboardAuth.AutoSize = $true
  $Global:cbx_ClipboardAuth.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:cbx_ClipboardAuth.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]230, [System.Int32]370))
  $Global:cbx_ClipboardAuth.Name = [System.String]'cbx_ClipboardAuth'
  $Global:cbx_ClipboardAuth.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]228, [System.Int32]17))
  $Global:cbx_ClipboardAuth.TabIndex = [System.Int32]78
  $Global:cbx_ClipboardAuth.Text = [System.String]'Enable Modern Auth Clipboard Integration'
  $Global:cbx_ClipboardAuth.UseVisualStyleBackColor = $true
  #
  #cliplabel
  #
  $Global:cliplabel.AutoSize = $true
  $Global:cliplabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]460, [System.Int32]370))
  $Global:cliplabel.Name = [System.String]'cliplabel'
  $Global:cliplabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]53, [System.Int32]13))
  $Global:cliplabel.TabIndex = [System.Int32]79
  $Global:cliplabel.TabStop = $true
  $Global:cliplabel.Text = [System.String]'more info'
  $Global:cliplabel.add_Click($Global:cliplabel_click)
  #
  #Global:SettingsForm
  #
  $Global:SettingsForm.BackColor = [System.Drawing.Color]::White
  $Global:SettingsForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]1044, [System.Int32]404))
  $Global:SettingsForm.Controls.Add($Global:Btn_Default)
  $Global:SettingsForm.Controls.Add($Global:grid_Tenants)
  $Global:SettingsForm.Controls.Add($Global:cbx_AutoUpdates)
  $Global:SettingsForm.Controls.Add($Global:btn_CancelConfig)
  $Global:SettingsForm.Controls.Add($Global:Btn_ReloadConfig)
  $Global:SettingsForm.Controls.Add($Global:Btn_SaveConfig)
  $Global:SettingsForm.Controls.Add($Global:cliplabel)
  $Global:SettingsForm.Controls.Add($Global:cbx_ClipboardAuth)
  $Global:SettingsForm.Name = [System.String]'Global:SettingsForm'
  $Global:SettingsForm.add_Load($Global:SettingsForm_Load)
  ([System.ComponentModel.ISupportInitialize]$Global:grid_Tenants).EndInit()
  $Global:SettingsForm.ResumeLayout($false)
  $Global:SettingsForm.PerformLayout()
  Add-Member -InputObject $Global:SettingsForm -Name base -Value $base -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name grid_Tenants -Value $Global:grid_Tenants -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name btn_CancelConfig -Value $Global:btn_CancelConfig -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Btn_ReloadConfig -Value $Global:Btn_ReloadConfig -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Btn_SaveConfig -Value $Global:Btn_SaveConfig -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Btn_Default -Value $Global:Btn_Default -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_ID -Value $Global:Tenant_ID -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_DisplayName -Value $Global:Tenant_DisplayName -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_Email -Value $Global:Tenant_Email -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_Credentials -Value $Global:Tenant_Credentials -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_ModernAuth -Value $Global:Tenant_ModernAuth -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_Teams -Value $Global:Tenant_Teams -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_Skype -Value $Global:Tenant_Skype -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_Exchange -Value $Tenant_Exchange -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_AzureAD -Value $Global:Tenant_AzureAD -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_Compliance -Value $Global:Tenant_Compliance -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name cbx_ClipboardAuth -Value $Global:cbx_ClipboardAuth -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name cliplabel -Value $Global:cliplabel -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name cbx_AutoUpdates -Value $Global:cbx_AutoUpdates -MemberType NoteProperty
  #endregion Gui
}

Function Import-BsGuiFunctions 
{
  #Gui Cancel button
  $Global:btn_CancelConfig_Click = 
  {
    Read-BsConfigFile
    Hide-BsGuiElements
  }

  #Gui Save Config Button
  $Global:Btn_SaveConfig_Click = 
  {
    $Global:btn_CancelConfig.Text = [System.String]'Close'
    Write-BsConfigFile
    If ($PSISE)
    {
      Update-BsAddonMenu
    }
  }

  #Gui Set Defaults Button
  $Global:Btn_Default_Click = 
  {
    Import-BsDefaultConfig
    If ($PSISE)
    {
      Update-BsAddonMenu
    }
  }

  #Gui Button to Reload Config
  $Global:Btn_ConfigReload_Click = 
  {
    Read-BsConfigFile
    If ($PSISE)
    {
      Update-BsAddonMenu
    }
  }
  #Gui link object to open a browser for more info on the modern auth clipboard
  $Global:cliplabel_click = 
  {
    Start-Process -FilePath 'https://UcMadScientist.com/BounShell/'
  }
}

Function Show-BsGuiElements
{
  #Reset the cancel button
  $Global:btn_CancelConfig.Text = [System.String]'Cancel'
  $null = $Global:SettingsForm.ShowDialog()
}

Function Hide-BsGuiElements
{
  $null = $Global:SettingsForm.Hide()
  If ($global:Config.AutoUpdatesEnabled)
  { 
    #Check for the required modules
    Write-Log -component $function -Message 'Checking for required modules based on selections. This can take some time.' -severity 2
    #Teams Module Check
    if ($global:Config.Tenant1.ConnectToTeams -or $global:Config.Tenant2.ConnectToTeams -or $global:Config.Tenant3.ConnectToTeams -or $global:Config.Tenant4.ConnectToTeams -or $global:Config.Tenant5.ConnectToTeams -or $global:Config.Tenant6.ConnectToTeams -or $global:Config.Tenant7.ConnectToTeams -or $global:Config.Tenant8.ConnectToTeams -or $global:Config.Tenant9.ConnectToTeams -or $global:Config.Tenant10.ConnectToTeams)
    {
      Test-BsInstalledModules -ModuleName $TestedTeamsModule -ModuleVersion $TestedTeamsModuleVer
      
    }

    #Exchange Module Check
    if ($global:Config.Tenant1.ConnectToExchange -or $global:Config.Tenant2.ConnectToExchange -or $global:Config.Tenant3.ConnectToExchange -or $global:Config.Tenant4.ConnectToExchange -or $global:Config.Tenant5.ConnectToExchange -or $global:Config.Tenant6.ConnectToExchange -or $global:Config.Tenant7.ConnectToExchange -or $global:Config.Tenant8.ConnectToExchange -or $global:Config.Tenant9.ConnectToExchange -or $global:Config.Tenant10.ConnectToExchange)
    {
      Test-BsInstalledModules -ModuleName $TestedExchangeModule -ModuleVersion $TestedExchangeModuleVer
    }
    
    #MsOnline Module Check
    if ($global:Config.Tenant1.ConnectToAzureAD -or $global:Config.Tenant2.ConnectToAzureAD -or $global:Config.Tenant3.ConnectToAzureAD -or $global:Config.Tenant4.ConnectToAzureAD -or $global:Config.Tenant5.ConnectToAzureAD -or $global:Config.Tenant6.ConnectToAzureAD -or $global:Config.Tenant7.ConnectToAzureAD -or $global:Config.Tenant8.ConnectToAzureAD -or $global:Config.Tenant9.ConnectToAzureAD -or $global:Config.Tenant10.ConnectToAzureAD)
    {
      Test-BsInstalledModules -ModuleName $TestedMSOnlineModule -ModuleVersion $TestedMSOnlineModuleVer
    }
    
    #Skype4B Module Check
    if ($global:Config.Tenant1.ConnectToSkype -or $global:Config.Tenant2.ConnectToSkype -or $global:Config.Tenant3.ConnectToSkype -or $global:Config.Tenant4.ConnectToSkype -or $global:Config.Tenant5.ConnectToSkype -or $global:Config.Tenant6.ConnectToSkype -or $global:Config.Tenant7.ConnectToSkype -or $global:Config.Tenant8.ConnectToSkype -or $global:Config.Tenant9.ConnectToSkype -or $global:Config.Tenant10.ConnectToSkype)
    {
      Test-BsInstalledModules -ModuleName $TestedSkype4BOModule -ModuleVersion $TestedSkype4BOModuleVer
    }
    
    Write-Log -component $function -Message 'Module check complete' -severity 2
  }
}

Function Start-BounShell
{
  $function = 'Start-BounShell'
  #Allows us to separate all the "onetime" run objects in case we get dot sourced.
  Write-Log -component $function -Message "Script executed from $PSScriptRoot" -severity 1
  Write-Log -component $function -Message 'Loading BounShell...' -severity 2

  #Load the Gui Elements
  Import-BsGuiElements

  #check for config file then load the default

  #Check for and load the config file if present
  If(Test-Path $global:ConfigFilePath)
  {
    Write-Log -component $function -Message "Found $ConfigFilePath, loading..." -severity 1
    Read-BsConfigFile
  }

  Else
  {
    Write-Log -component $function -Message 'Could not locate config file, Using Defaults' -severity 3
    #If there is no config file. Load a default
    Import-BsDefaultConfig

    Write-Log -component $function -Message "As we didn't find a config file we will assume this is a first run." -severity 3
    Write-Log -component $function -Message 'Thus we will remind you that while all care is taken to store your credentials in a safe manner, we cannot be held responsible for any data breaches' -severity 3
    Write-Log -component $function -Message "If someone was to get a hold of your BounShell.xml AND your user profile private encryption key it's possible to reverse engineer stored credentials" -severity 3
    Write-Log -component $function -Message "Seriously, whilst the password store is encrypted, it's not perfect!" -severity 3
    Pause
  }

  #check for script update
  if ($SkipUpdateCheck -eq $false)
  {
    Get-ScriptUpdate
  } 

  #Check for Modules
  
  
  #Now Create the Objects in the ISE
  If($PSISE) 
  {
    Update-BsAddonMenu
  }
  
  
  Write-Log -component $function -Message 'BounShell Loaded' -severity 2
  
  #Check we are actually in the ISE
  If(!$PSISE) 
  {   
    Write-Log -component $function -Message 'Could not locate $PSISE Variable' -severity 1
    Write-Log -component $function -Message 'Launched BounShell without ISE Support, Keyboard Shortcuts will be unavailable' -severity 2
    Write-Host -Object ''
    Write-Log -component $function -Message 'To configure BounShell tenants run Show-BsGuiElements' -severity 2
    Write-Log -component $function -Message 'To connect to a tenant run Connect-BsO365Tenant'  -severity 2
    Return #Yes I know Return sucks, but I think it's better than Throw.
  }
}

Function Watch-BsCredentials
{
  [CmdletBinding()]
  PARAM
  (
    $ModernAuthUsername,
    $UnsecurePassword,
    [switch]$NoPassword
  )
  [string]$function = 'Watch-BsCredentials'
  If (!$ModernAuthUsername)
  {
    Write-Log -component $function -Message "This Cmdlet is for BoundShell's internal use only. Please use 'Start-Bounshell' to launch the tool" -severity 3
    Pause
    return
  }
  Write-Log -component $function -Message "Called to connect to $ModernAuthUsername" -severity 3
  # Load the API we need for Keypresses
  $signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
    public static extern short GetAsyncKeyState(int virtualKeyCode); 
'@

  # load signatures and make members available
  $API = Add-Type -MemberDefinition $signature -Name 'Keypress' -Namespace API -PassThru
    
  # define that we are waiting for 'v' keypress
  $waitFor = 'v'
  $ascii = [byte][char]$waitFor.ToUpper()

  Set-Clipboard -Value $ModernAuthUsername
  Write-Log -component $function -Message "$ModernAuthUsername placed into Clipboard" -severity 1
  Write-Log -component $function -Message "Press 'Ctrl+v' to paste the username $ModernAuthUsername in the modern auth window" -severity 2


  do 
  {
    Start-Sleep -Milliseconds 40
  }
  until ($API::GetAsyncKeyState($ascii) -eq -32767)
  
  Start-Sleep -Milliseconds 1000
  if ($NoPassword)
  {
    Set-Clipboard -Value $UnsecurePassword
    Write-Log -component $function -Message 'Password placed into Clipboard' -severity 1
    Write-Log -component $function -Message "Press 'Ctrl+v' to paste the password in the modern auth window" -severity 2

    do 
    {
      Start-Sleep -Milliseconds 40
    }
    until ($API::GetAsyncKeyState($ascii) -eq -32767)

    Set-Clipboard -Value 'Thanks for using BounShell'
  }
}

Function Test-BsInstalledModules 
{
  param
  (
    [Parameter(Mandatory)] [string]$ModuleName,
    [Parameter()] [string]$ModuleVersion
  )
  [string]$function = 'Test-BsInstalledModules'
  Write-Log -component $function -Message "Called to check $ModuleName" -severity 1
   
  $NeedsInstall = $false
  $needsupdate = $false
  $NeedsCleaup = $false
  $LatestModuleVersion = $null
  
  #Pull the module from the local machine
  $Module = Get-Module -Name $ModuleName -ListAvailable
  
  #If we have more than one version we need to clean up
  if($Module.count -gt 1)
  {
    $needscleanup = $true
    
    # Identify modules with multiple versions installed
    $MultiModules = $Module | Group-Object -Property name -NoElement | Where-Object -Property count -GT -Value 1

    $title = "Multiple copies of $ModuleName installed"
    $Message = "I've detected multiple installs of $ModuleName. Should I remove them and install the latest version?"
    $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
    'Cleans up the installed PowerShell Modules'

    $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
    'No thanks.'

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

    switch ($result)
    {
      0 
      {
        #User said yes
        Write-Log -component $function -Message 'User opted to cleanup modules' -severity 1
        #start $BlogPost
        Repair-BsInstalledModules -ModuleName $ModuleName -Operation 'Cleanup'
        Write-Log -component $function -Message 'Cleanup completed' -severity 1
      }
      #User said no
      1 
      {
        Write-Log -component $function -Message 'User opted to skip cleanup' -severity 1
        Write-Log -component $function -Message "Found multiple copies of $ModuleName" -severity 3
        Write-Log -component $function -Message "To correct this later run 'Repair-BsInstalledModules -Operation Cleanup -ModuleName $ModuleName'" -severity 3
      }
    }
  }
   
  #If we have one module, check it's up to date
  if($Module.count -eq 1)
  {
    Write-Log -component $function -Message "Found one copy of $ModuleName" -severity 1
  }
  

  #Module not installed
  if($Module.count -eq 0)
  {
    Write-Log -component $function -Message "$ModuleName Module not installed on local computer" -severity 3
    $NeedsInstall = $true
    $title = "$ModuleName not found"
    $Message = "The $ModuleName PowerShell module is not installed on this computer, it is required to connect to your requested services. Can I install it for you?"
    $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
    'Installs required PowerShell modules'

    $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
    'No thanks.'

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

    switch ($result)
    {
      0 
      {
        #User said yes
        Write-Log -component $function -Message 'User opted to install module' -severity 1
        #start $BlogPost
        Repair-BsInstalledModules -ModuleName $ModuleName -Operation 'Install'
        Write-Log -component $function -Message 'Install completed' -severity 1
      }
      #User said no
      1 
      {
        Write-Log -component $function -Message 'User opted to skip Install' -severity 1
        Write-Log -component $function -Message "$ModuleName not found" -severity 3
        Write-Log -component $function -Message "To correct this later run 'Repair-BsInstalledModules -Operation Install -ModuleName $ModuleName'" -severity 3
      }
    }
  }

  #Okay, we have checked if everything is installed, now lets check and report on the version
 
  Write-Log -component $function -Message "Checking for the latest version of $ModuleName in the PSGallery" -severity 2
  $gallery = $Module.where({
      $_.repositorysourcelocation
  })

  foreach ($Module in $gallery) 
  {
    #find the current version in the gallery
    Try 
    {
      $online = Find-Module -Name $Module.name -Repository PSGallery -ErrorAction Stop
    }
    Catch 
    {
      #Todo What the f**k?
      Write-Warning -Message ('Module {0} was not found in the PSGallery' -f $Module.name)
    }

    #compare versions
    if ($online.version -gt $Module.version) 
    {
      $needsupdate = $true
      Write-Log -component $function -Message "An updated version of the $ModuleName module is available in the PSGallery" -severity 2
      $title = "Update for $ModuleName found"
      $Message = "There is a newer version of $ModuleName on the PowerShell gallery, would you like me to install it?"
      $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
      'Updates the PowerShell module'

      $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
      'No thanks.'

      $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

      $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

      switch ($result)
      {
        0 
        {
          #User said yes
          Write-Log -component $function -Message 'User opted to update module' -severity 1
          #start $BlogPost
          Repair-BsInstalledModules -ModuleName $ModuleName -Operation 'Update'
          Write-Log -component $function -Message 'Update completed' -severity 1
        }
        #User said no
        1 
        {
          Write-Log -component $function -Message 'User opted to skip update' -severity 1
          Write-Log -component $function -Message "Found updated version of $ModuleName" -severity 3
          Write-Log -component $function -Message "To update this later run 'Repair-BsInstalledModules -Operation Update -ModuleName $ModuleName'" -severity 3
        }
      }
    }
     
    else 
    {
      Write-Log -component $function -Message "Your version of the $ModuleName module is up to date" -severity 2
    }
  }
  

  ## for the official Excahnge online module
  ## Connect-IPPSSession
  ##https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application
}

Function Repair-BsInstalledModules 
{
  param
  (
    [Parameter(Mandatory)] [string]$ModuleName,
    [Parameter(Mandatory)] [string]$Operation #Install, Cleaup, Update
  )
 
  [string]$function = 'Repair-BsInstalledModules'
  Write-Log -component $function -Message "Called to $Operation $ModuleName" -severity 1
  Write-Log -component $function -Message 'Checking for elevated session' -severity 1
  #Check we are running as admin, we dont want to install modules in the users context
  if (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')))
  {
    # throw 'Please Note: You are trying to run this script without elevated Administator Privileges. In order to run this script you will require PowerShell running in Administrator Mode'
    Write-Log -component $function -Message 'Not running as Administrator, invoking new session' -severity 2
    Write-Log -component $function -Message 'You must close all tabs and restart the ISE once this process completes!' -severity 3
    $newProcess = New-Object -TypeName System.Diagnostics.ProcessStartInfo -ArgumentList 'PowerShell'
   
    # Specify the current script path and name as a parameter
    
    #Old non multi session friendly version
    #$newProcess.Arguments = "Repair-BsInstalledModules -modulename $ModuleName -operation $Operation"
    $newProcess.Arguments = "Test-BsInstalledModules -modulename $ModuleName"
   
    # Indicate that the process should be elevated
    $newProcess.Verb = 'runas'
   
    # Start the new process
    $foo = [System.Diagnostics.Process]::Start($newProcess)
  }
  else 
  {
    Write-Log -component $function -Message "Running as Administrator, performing $Operation" -severity 2
  
    if($ModuleName -eq 'MSExchange')
    {
      Write-Log -component $function -Message 'The Exchange Online PowerShell Module is not in the PsGallery, it is installed via your browser as a ClickOnce app' -severity 2
      Write-Log -component $function -Message 'Attempting to install now, if this fails please visit' -severity 2
      Write-Log -component $function -Message 'https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application ' -severity 2
      Write-Log -component $function -Message 'using Internet Explorer' -severity 2
      Start-Process -FilePath 'Iexplore https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application'
    }
    else
    {
      Switch ($Operation)

      {
        'Install'
        { 
          $output = (Install-Module -name $ModuleName)
          Write-Log -component $function -Message 'Complete!' -severity 2
          Get-Module -Name $ModuleName -ListAvailable
          Start-Sleep -Seconds 5
        }
        'Update'
        { 
          $output = (Update-Module -name $ModuleName)
          Write-Log -component $function -Message 'Complete!' -severity 2
          Get-Module -Name $ModuleName -ListAvailable
          Start-Sleep -Seconds 5
        }
        'Cleanup'
        { 
          #Pull the currently installed modules
          $Modules = Get-Module -Name $ModuleName -ListAvailable
          # Identify modules with multiple versions installed
          $gallery = $Modules.where({
              $_.repositorysourcelocation
          })
        
          foreach ($mod in $gallery)
          {
            Write-Log -component $function -Message "Removing $($mod.version)" -severity 2
            Uninstall-Module -Name $ModuleName -RequiredVersion $mod.version     
          }
          Install-Module -name $ModuleName
        }
 
      }
    }
  }
}

Function Import-BsWPFGuiElements
{
  #WPF handler heavily based of FoxDeploy's work. Go check him out and say I said "Hi!"
  $function = 'Import-BsWPFGuiElements'
  Write-Log -component $function -Message "Called $function" -severity 1 
  Write-Log -component $function -Message "Loading XAML into variable" -severity 1 
  $inputXML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Settings_Gui_2"
        xmlns:Themes="clr-namespace:Microsoft.Windows.Themes;assembly=PresentationFramework.Aero2" x:Name="Settings"
        Title="BounShell Settings" Height="450" Width="1086" ResizeMode="NoResize">
    <Window.Resources>
        <SolidColorBrush x:Key="ListBorder" Color="#828790"/>
        <Style x:Key="ListViewStyle1" TargetType="{x:Type ListView}">
            <Setter Property="Background" Value="{DynamicResource {x:Static SystemColors.WindowBrushKey}}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ListBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="#FF042271"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="true"/>
            <Setter Property="ScrollViewer.PanningMode" Value="Both"/>
            <Setter Property="Stylus.IsFlicksEnabled" Value="False"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ListView}">
                        <Themes:ListBoxChrome x:Name="Bd" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Background="{TemplateBinding Background}" RenderMouseOver="{TemplateBinding IsMouseOver}" RenderFocused="{TemplateBinding IsKeyboardFocusWithin}" SnapsToDevicePixels="true"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Background" TargetName="Bd" Value="{DynamicResource {x:Static SystemColors.ControlBrushKey}}"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsGrouping" Value="true"/>
                                    <Condition Property="VirtualizingPanel.IsVirtualizingWhenGrouping" Value="false"/>
                                </MultiTrigger.Conditions>
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid x:Name="grd_Main">
        <ListView x:Name="lst_Tenant" HorizontalAlignment="Left" Margin="10,10,0,79.656" Width="298.223">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="ID"/>
                    <GridViewColumn Header="Display Name"/>
                    <GridViewColumn Header="Sign In Address"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button x:Name="btn_NewTenant" Content="New Connection" HorizontalAlignment="Left" Margin="10,0,0,45.657" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <Button x:Name="btn_DeleteTenant" Content="Delete Connection" HorizontalAlignment="Left" Margin="181.7,0,0,45.657" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <Button x:Name="btn_SaveTenant" Content="Save Connection" HorizontalAlignment="Left" Margin="181.7,0,0,10" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <Button x:Name="btn_RevertTenant" Content="Revert Connection" HorizontalAlignment="Left" Margin="10,0,0,10" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <TextBlock x:Name="txt_TenantID" Height="17" Margin="313.223,10,0,0" TextWrapping="Wrap" Text="Tenant ID" VerticalAlignment="Top" HorizontalAlignment="Left" Width="144.361" FontWeight="SemiBold"/>
        <TextBox x:Name="tbx_TenantID" Height="20" Margin="464,5,0,0" TextWrapping="Wrap" Text="TenantName" VerticalAlignment="Top" HorizontalAlignment="Left" Width="28.916" IsEnabled="False"/>
        <Button x:Name="btn_MoveTenantUp" Content="Move Up" Margin="497.916,5,516.561,0" Height="20" VerticalAlignment="Top"/>
        <Button x:Name="btn_MoveTenantDown" Content="Move Dn" Margin="0,5,448.038,0" Height="20" VerticalAlignment="Top" HorizontalAlignment="Right" Width="63.523"/>
        <TextBlock x:Name="txt_TenantShortCut" Height="17" Margin="378.365,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="80.635"><Run Text="Ctrl + Alt + "/><Run Text="0"/></TextBlock>
        <TextBlock x:Name="txt_TenantDisplayName" Height="19" Margin="313.223,32,0,0" TextWrapping="Wrap" Text="Display name" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"/>
        <TextBox x:Name="tbx_TenantDisplayName" Height="20" Margin="464,30,391.57,0" TextWrapping="Wrap" Text="Tenant Name" VerticalAlignment="Top"/>
        <TextBlock x:Name="txt_TenantSignInAddress" Height="19" Margin="313.223,56,0,0" TextWrapping="Wrap" Text="Sign In Address" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"/>
        <TextBox x:Name="tbx_TenantSignInAddress" Height="20" Margin="464,55,391.57,0" TextWrapping="Wrap" Text="Example@Tenant.com" VerticalAlignment="Top"/>
        <TextBlock x:Name="txt_TenantPassword" Height="19" Margin="313.223,80,0,0" TextWrapping="Wrap" Text="Password" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"/>
        <PasswordBox x:Name="tbx_TenantPassword" Height="19" Margin="464,78,391.57,0" VerticalAlignment="Top" Password="This is a fake password"/>
        <TextBlock x:Name="txt_TenantRootDomain" Height="19" Margin="313.223,104,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"><Run Text="O"/><Run Text="n"/><Run Text="M"/><Run Text="icrosoft "/><Run Text="sub"/><Run Text="domain"/></TextBlock>
        <TextBox x:Name="tbx_TenantRootDomain" Height="20" Margin="464,102,499.449,0" TextWrapping="Wrap" VerticalAlignment="Top" Text="exampletenant"/>
        <TextBlock x:Name="txt_OnMicrososft_com" Margin="0,102,392.131,0" TextWrapping="Wrap" Height="19" VerticalAlignment="Top" HorizontalAlignment="Right" Width="102.318"><Run Text="."/><Run Text="onmicrosoft.com"/></TextBlock>
        <CheckBox x:Name="cbx_TenantModernAuth" Content="Requires Modern Auth (Multi Factor)" Height="17" Margin="319.793,129,405,0" VerticalAlignment="Top"/>
        <TextBlock x:Name="txt_TenantDisplayNameChanged" Height="19" Margin="0,32,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Visibility="Hidden"/>
        <TextBlock x:Name="txt_TenantSignInAddressChanged" Height="19" Margin="0,58,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Visibility="Hidden"/>
        <TextBlock x:Name="txt_TenantPasswordChanged" Height="19" Margin="0,82,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Visibility="Hidden"/>
        <TextBlock x:Name="txt_TenantRootDomainChanged" Height="19" Margin="0,106,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Visibility="Hidden"/>
        <TextBlock x:Name="txt_TenantModernAuthChanged" Height="19" Margin="0,127,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Visibility="Hidden"/>
        <Separator Height="14" Margin="310.223,151,378.991,0" VerticalAlignment="Top"/>
        <Grid x:Name="grd_ConnectionOptions" Margin="310.223,165,378.991,107">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="211*"/>
                <ColumnDefinition Width="180*"/>
            </Grid.ColumnDefinitions>
            <CheckBox x:Name="cbx_ConnectToTeams" Content="Connect to Microsoft Teams PowerShell" Margin="10,84.96,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToSkype" Content="Connect to Skype for Business Online PowerShell" Margin="10,125.156,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToAzureAD" Content="Connect to Azure AD PowerShell" Margin="10,24.666,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToAzureCompliance" Margin="10,44.764,10,0" Content="Connect to Azure Compliance Centre" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToSharepointOnline" Content="Connect to SharePoint Online PowerShell" Margin="10,105.058,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToExchangeOnline" Content="Connect to Exchange Online PowerShell" Margin="10,64.862,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <TextBlock x:Name="txt_TenantServiceOptions" Height="19" Margin="2.579,0.666,57.421,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold"><Run Text="Tenant "/><Run Text="Service Connections"/></TextBlock>
            <TextBlock x:Name="txt_TenantConnectToAzureADChanged" Height="19" Margin="0,23.666,0,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Visibility="Hidden"/>
            <TextBlock x:Name="txt_TenantConnectToAzureComplianceChanged" Height="19" Margin="0,43.764,0,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Visibility="Hidden"/>
            <TextBlock x:Name="txt_TenantConnectToExchangeOnlineChanged" Margin="0,67.01,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" VerticalAlignment="Top" Visibility="Hidden"/>
            <TextBlock x:Name="txt_TenantConnectToSharePointChanged" Margin="0,0,0,24.942" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1" Visibility="Hidden"/>
            <TextBlock x:Name="txt_TenantConnectToSkypeOnlineChanged" Margin="0,0,0,5.844" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1" Visibility="Hidden"/>
            <TextBlock x:Name="txt_TenantConnectToTeamsChanged" Margin="0,0,0,42.892" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1" Visibility="Hidden"/>
        </Grid>
        <Grid x:Name="grd_BounShellOptions" Margin="310.223,0,378.991,10" Height="92" VerticalAlignment="Bottom">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="105*"/>
                <ColumnDefinition Width="94*"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_BounShellTitle" Height="19" Margin="2.5,1,27.639,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold" Grid.ColumnSpan="2" Text="BounShell Specific Options"/>
            <CheckBox x:Name="cbx_BounShellCheckForPowerShellUpdates" Content="Check for PowerShell module updates" Margin="10,45.098,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top" IsChecked="True"/>
            <CheckBox x:Name="cbx_BounShellCheckForUpdates" Content="Check for BounShell updates" Margin="10,25,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top" IsChecked="True"/>
            <CheckBox x:Name="cbx_BounShellEnableModernAuthClipboard" Content="Enable Modern Auth clipboard integration" Margin="10,65.196,-58.648,0" VerticalAlignment="Top" IsChecked="True"/>
            <TextBlock x:Name="txt_BounShellModernAuthLink" Margin="63.648,0,27.639,11.844" TextWrapping="Wrap" Grid.Column="1" Foreground="Blue" TextDecorations="Underline" VerticalAlignment="Bottom"><Run Text="Learn more."/><Run Text=".."/></TextBlock>
            <TextBlock x:Name="txt_BounShellCheckForPowerShellUpdatesChanged" Margin="0,24,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Top" Grid.Column="1" Visibility="Hidden"/>
            <TextBlock x:Name="txt_BounShellCheckForUpdatesChanged" Margin="0,0,0,25.754" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="19" VerticalAlignment="Bottom" Visibility="Hidden"/>
            <TextBlock x:Name="txt_BounShellEnableModernAuthClipboardChanged" Margin="0,0,0,5.656" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="19" VerticalAlignment="Bottom" Visibility="Hidden"/>
        </Grid>
        <Separator Margin="507.009,197,165.991,199" RenderTransformOrigin="0.5,0.5">
            <Separator.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="-90"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Separator.RenderTransform>
        </Separator>
        <Grid x:Name="grd_SkypeOptions" Margin="0,4.05,10,0" HorizontalAlignment="Right" Width="354.777" Height="136.336" VerticalAlignment="Top" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="105*"/>
                <ColumnDefinition Width="94*"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_SkypeTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFD447D4" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Skype for Business "/><Run Text="Online"/></TextBlock>
            <CheckBox x:Name="cbx_SkypeOverideDomain" Content="Admin Override Domain" Margin="10,0,5,10" Height="17" VerticalAlignment="Bottom"/>
            <CheckBox x:Name="cbx_SkypeAllowReconnect" Content="Allow Reconnect (Fixes ISE hard lockup)" Margin="10,45.335,10,0" IsChecked="True" Height="17" VerticalAlignment="Top" Grid.ColumnSpan="2"/>
            <CheckBox x:Name="cbx_SkypeOverideAdmin" Content="Override Discovery URI" Margin="10,0,10,32" Height="18.001" VerticalAlignment="Bottom"/>
            <TextBlock x:Name="txt_SkypeISEFixes" Height="19" Margin="10,25.335,55.639,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold"><Run Text="ISE "/><Run Text="Bug Fixes"/></TextBlock>
            <TextBlock x:Name="txt_SkypeConnectionFix" Margin="10,62.335,20.139,55.001" TextWrapping="Wrap" FontWeight="SemiBold" Grid.ColumnSpan="2"><Run Text="BPOS / renamed "/><Run Text="/ hybrid "/><Run Text="tenant"/><Run Text=" and "/><Run Text="delegated admin"/><Run Text=" "/><Run Text="workarounds"/></TextBlock>
            <TextBox x:Name="tbx_SkypeDiscoveryUri" Margin="0,0,15.009,31" TextWrapping="Wrap" Text="fabrikam.onmicrosoft.com" Grid.Column="1" Height="21.001" VerticalAlignment="Bottom"/>
            <TextBox x:Name="tbx_SkypeAdminDomain" Margin="0,0,15.009,10" TextWrapping="Wrap" Height="20" VerticalAlignment="Bottom" Text="fabrikam.com" Grid.Column="1"/>
            <TextBlock x:Name="txt_SkypeDiscoveryURIChanged" Margin="0,0,0,30" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Bottom" RenderTransformOrigin="0.001,-2.169" Visibility="Hidden"/>
            <TextBlock x:Name="txt_SkypeAdminDomainChanged" Margin="0,0,0,11" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1" Visibility="Hidden"/>
            <TextBlock x:Name="txt_SkypeAllowReconnectChanged" Margin="0,45.335,-0.009,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" RenderTransformOrigin="0.001,-2.169" Visibility="Hidden"/>
        </Grid>
        <Grid x:Name="grd_AzureADOptions" Margin="0,135,10,0" HorizontalAlignment="Right" Width="354.777" Height="52.336" VerticalAlignment="Top" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="105*"/>
                <ColumnDefinition Width="94*"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_AzureADTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFFF9623" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Azure AD"/></TextBlock>
            <CheckBox x:Name="cbx_AzureADEnviromentName" Content="Azure Environment Name" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_AzureADEnviromentName" Margin="0,25.335,15.009,0" TextWrapping="Wrap" Text="AzureCloud" Grid.Column="1" Height="21.001" VerticalAlignment="Top"/>
            <TextBlock x:Name="txt_AzureADEnviromentNameChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" Visibility="Hidden"/>
        </Grid>
        <Grid x:Name="grd_AzureComplianceOptions" Margin="0,182.336,10,184.328" HorizontalAlignment="Right" Width="354.777" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_AzureComplianceTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FF29FCFF" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Azure "/><Run Text="Compliance"/></TextBlock>
            <CheckBox x:Name="cbx_AzureComplianceURI" Content="Connection Uri" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_AzureComplianceURI" Margin="0,25.335,15.009,0" Text="https://ps.compliance.protection.outlook.com/powershell-liveid/" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_AzureComplianceURIChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" Visibility="Hidden"/>
        </Grid>
        <Grid x:Name="grd_ExchangeOnlineOptions" Margin="0,0,10,131.992" HorizontalAlignment="Right" Width="354.777" Height="52.336" VerticalAlignment="Bottom" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_ExchangeOnlineTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFFA9EDA" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Exchange Online"/><LineBreak/><Run/></TextBlock>
            <CheckBox x:Name="cbx_ExchangeOnlineURI" Content="Connection Uri" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_ExchangeOnlineURI" Margin="0,25.335,15.009,0" Text="https://outlook.Office365.com/powershell-liveid/" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_ExchangeOnlineURIChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" Visibility="Hidden"/>
        </Grid>
        <Grid x:Name="grd_TeamsOptions" Margin="0,0,10,79.656" HorizontalAlignment="Right" Width="354.777" Height="52.336" VerticalAlignment="Bottom" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_TeamsTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="White" Background="Black"><Run Text=" "/><Run Text="■"/><Run Text=" Microsoft Teams"/></TextBlock>
            <CheckBox x:Name="cbx_TeamsEnviroment" Content="Teams Environment Name" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_TeamsEnviroment" Margin="10,25.335,15.009,0" Text="TeamsGCCH" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_TeamsEnviromentChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" Visibility="Hidden"/>
        </Grid>
        <Grid x:Name="grd_SharePointOptions" Margin="0,0,10,11" HorizontalAlignment="Right" Width="354.777" Height="68.656" VerticalAlignment="Bottom" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_SharepointTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFF9FF81" Background="Black"><Run Text=" "/><Run Text="■"/><Run Text=" "/><Run Text="SharePoint Online"/></TextBlock>
            <CheckBox x:Name="cbx_SharepointURL" Content="URL" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_SharepointURL" Margin="10,25.335,15.009,0" Text="https://contoso-admin.sharepoint.com" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_SharepointURLChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" Visibility="Hidden"/>
            <CheckBox x:Name="cbx_SharepointRegion" Content="Region" Margin="10,47.336,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_SharepointRegion" Margin="10,45.336,15.009,0" Text="Default" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_SharepointRegionChanged" Margin="0,45.336,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" Visibility="Hidden"/>
        </Grid>
    </Grid>
</Window>


'@ 
  
  #Clean out stuff PS doesnt support
  $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
  
  #Load WPF and format the XAML as XML
  Write-Log -component $function -Message "Loading WPF and putting XAML into XML" -severity 1 
  [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
  [xml]$XAML = $inputXML
  $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  
  #Call the WPF instance to load the window
  Write-Log -component $function -Message "Calling WPF" -severity 1 
  try{$global:SettingsWpfWindow=[Windows.Markup.XamlReader]::Load( $reader )}
  catch [System.Management.Automation.MethodInvocationException] {
    Write-Log -component $function -Message "Error loading BounShell GUI, likley something wrong with the XAML code" -severity 3
    Write-Log -component $function -Message $error[0].Exception.Message -severity 2 
  }
  catch{#if it broke some other way 
    Write-Log -component $function -Message "Error loading BounShell GUI, something I didnt catch!" -severity 3
    Write-Log -component $function -Message $error[0].Exception.Message -severity 2 
  }
    
  #Okay, WPF loaded, now we need to be able to find them in PowerShell
  Write-Log -component $function -Message "Importing WPF objects into PS variables" -severity 1
  $xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Scope Global -Value $global:SettingsWpfWindow.FindName($_.Name)}
  
  #todo, remove this.
  Function Get-FormVariables
  {
    get-variable WPF*
  }
    
  Get-FormVariables
  
}

Function Import-BsWpfGuiFunctions 
{
  $function = 'Import-BsWPFGuiElements'
  Write-Log -component $function -Message "Called $function" -severity 1 
  
  #check to see if the form is actually loaded. This doesnt capture if the window is loaded and "closed" and is thus, shit
  Write-Log -component $function -Message "Checking to see if the form is already loaded" -severity 1 
  
  If (Test-Path variable:global:SettingsWpfWindow)
  {
    Write-Log -component $function -Message "GUI variable exists" -severity 1 
  }
  Else
  {
    Write-Log -component $function -Message "GUI variable missing, reloading XAML" -severity 3
    Import-BSWPFGuiElements
  }
  
  #Stop WPF from unloading the window when the user clicks close
  
  $global:SettingsWpfWindow.Add_Closing({
      $_.Cancel = $true
      Hide-BsWPFGuiElements
  })


  #region ServiceToggles
  
  #Azure AD Toggle
  $WPFcbx_ConnectToAzureAD.add_Click({
      If ($WPFcbx_ConnectoAzureAD.Ischecked)                                                                                               
        {$WPFgrd_AzureADOptions.IsEnabled = $True}
        Else
        {$WPFgrd_AzureADOptions.IsEnabled = $False}
      $WPFtxt_TenantConnectToAzureADChanged.Visibility = "Visible"
    
  })
  #Azure Compliance Toggle
  $WPFcbx_ConnectToAzureCompliance.add_Click({
      If ($WPFcbx_ConnectToAzureCompliance.Ischecked)
        {$WPFgrd_AzureComplianceOptions.IsEnabled = $True}
        Else
        {$WPFgrd_AzureComplianceOptions.IsEnabled = $False}
      $WPFtxt_TenantConnectToAzureComplianceChanged.Visibility = "Visible"
    
  })
  #Exchange Online Toggle
  $WPFcbx_ConnectToExchangeOnline.add_Click({
      If ($WPFcbx_ConnectToExchangeOnline.Ischecked)
        {$WPFgrd_ExchangeOnlineOptions.IsEnabled = $True}
        Else
        {$WPFgrd_ExchangeOnlineOptions.IsEnabled = $False}
      $WPFtxt_TenantConnectToExchangeOnlineChanged.Visibility = "Visible"
    
  })
  
  #Teams Toggle
  $WPFcbx_ConnectToTeams.add_Click({
      If ($WPFcbx_ConnectToTeams.Ischecked)
        {$WPFgrd_TeamsOptions.IsEnabled = $True}
        Else
        {$WPFgrd_TeamsOptions.IsEnabled = $False}
      $WPFtxt_TenantConnectToTeamsChanged.Visibility = "Visible"
    
  })
  
  #Sharepoint Toggle
  $WPFcbx_ConnectToSharepointOnline.add_Click({
      If ($WPFcbx_ConnectToSharepointOnline.Ischecked)
        {$WPFgrd_SharePointOptions.IsEnabled = $True}
        Else
        {$WPFgrd_SharePointOptions.IsEnabled = $False}
      $WPFtxt_TenantConnectToSharePointChanged.Visibility = "Visible"
    
  })
  
  #Skype Online Toggle
  $WPFcbx_ConnectToSkype.add_Click({
      If ($WPFcbx_ConnectToSkype.Ischecked)
        {$WPFgrd_SkypeOptions.IsEnabled = $True}
        Else
        {$WPFgrd_SkypeOptions.IsEnabled = $False}
      $WPFtxt_TenantConnectToSkypeOnlineChanged.Visibility = "Visible"
    
  })
  #endregion ServiceToggles 
}

Function Import-BsMultiThreadedWpfGuiElements
{
  $Global:syncHash = [hashtable]::Synchronized(@{})
  $newRunspace =[runspacefactory]::CreateRunspace()
  $newRunspace.ApartmentState = "STA"
  $newRunspace.ThreadOptions = "ReuseThread"
  $newRunspace.Open()
  $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)


  $global:psCmd = [PowerShell]::Create().AddScript({
      [xml]$xaml = @'

<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Settings_Gui_2"
        xmlns:Themes="clr-namespace:Microsoft.Windows.Themes;assembly=PresentationFramework.Aero2" x:Name="Settings"
        Title="BounShell Settings" Height="450" Width="1086" ResizeMode="NoResize">
    <Window.Resources>
        <SolidColorBrush x:Key="ListBorder" Color="#828790"/>
        <Style x:Key="ListViewStyle1" TargetType="{x:Type ListView}">
            <Setter Property="Background" Value="{DynamicResource {x:Static SystemColors.WindowBrushKey}}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ListBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="#FF042271"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="true"/>
            <Setter Property="ScrollViewer.PanningMode" Value="Both"/>
            <Setter Property="Stylus.IsFlicksEnabled" Value="False"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ListView}">
                        <Themes:ListBoxChrome x:Name="Bd" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Background="{TemplateBinding Background}" RenderMouseOver="{TemplateBinding IsMouseOver}" RenderFocused="{TemplateBinding IsKeyboardFocusWithin}" SnapsToDevicePixels="true"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Background" TargetName="Bd" Value="{DynamicResource {x:Static SystemColors.ControlBrushKey}}"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsGrouping" Value="true"/>
                                    <Condition Property="VirtualizingPanel.IsVirtualizingWhenGrouping" Value="false"/>
                                </MultiTrigger.Conditions>
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid x:Name="grd_Main">
        <ListView x:Name="lst_Tenant" HorizontalAlignment="Left" Margin="10,10,0,79.656" Width="298.223">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="ID"/>
                    <GridViewColumn Header="Display Name"/>
                    <GridViewColumn Header="Sign In Address"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button x:Name="btn_NewTenant" Content="New Connection" HorizontalAlignment="Left" Margin="10,0,0,45.657" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <Button x:Name="btn_DeleteTenant" Content="Delete Connection" HorizontalAlignment="Left" Margin="181.7,0,0,45.657" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <Button x:Name="btn_SaveTenant" Content="Save Connection" HorizontalAlignment="Left" Margin="181.7,0,0,10" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <Button x:Name="btn_RevertTenant" Content="Revert Connection" HorizontalAlignment="Left" Margin="10,0,0,10" Width="128.523" Height="28.999" VerticalAlignment="Bottom"/>
        <TextBlock x:Name="txt_TenantID" Height="17" Margin="313.223,10,0,0" TextWrapping="Wrap" Text="Tenant ID" VerticalAlignment="Top" HorizontalAlignment="Left" Width="144.361" FontWeight="SemiBold"/>
        <TextBox x:Name="tbx_TenantID" Height="20" Margin="464,5,0,0" TextWrapping="Wrap" Text="TenantName" VerticalAlignment="Top" HorizontalAlignment="Left" Width="28.916" IsEnabled="False"/>
        <Button x:Name="btn_MoveTenantUp" Content="Move Up" Margin="497.916,5,516.561,0" Height="20" VerticalAlignment="Top"/>
        <Button x:Name="btn_MoveTenantDown" Content="Move Dn" Margin="0,5,448.038,0" Height="20" VerticalAlignment="Top" HorizontalAlignment="Right" Width="63.523"/>
        <TextBlock x:Name="txt_TenantShortCut" Height="17" Margin="378.365,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="80.635"><Run Text="Ctrl + Alt + "/><Run Text="0"/></TextBlock>
        <TextBlock x:Name="txt_TenantDisplayName" Height="19" Margin="313.223,32,0,0" TextWrapping="Wrap" Text="Display name" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"/>
        <TextBox x:Name="tbx_TenantDisplayName" Height="20" Margin="464,30,391.57,0" TextWrapping="Wrap" Text="Tenant Name" VerticalAlignment="Top"/>
        <TextBlock x:Name="txt_TenantSignInAddress" Height="19" Margin="313.223,56,0,0" TextWrapping="Wrap" Text="Sign In Address" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"/>
        <TextBox x:Name="tbx_TenantSignInAddress" Height="20" Margin="464,55,391.57,0" TextWrapping="Wrap" Text="Example@Tenant.com" VerticalAlignment="Top"/>
        <TextBlock x:Name="txt_TenantPassword" Height="19" Margin="313.223,80,0,0" TextWrapping="Wrap" Text="Password" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"/>
        <PasswordBox x:Name="tbx_TenantPassword" Height="19" Margin="464,78,391.57,0" VerticalAlignment="Top" Password="This is a fake password"/>
        <TextBlock x:Name="txt_TenantRootDomain" Height="19" Margin="313.223,104,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold" HorizontalAlignment="Left" Width="150.777"><Run Text="O"/><Run Text="n"/><Run Text="M"/><Run Text="icrosoft "/><Run Text="sub"/><Run Text="domain"/></TextBlock>
        <TextBox x:Name="tbx_TenantRootDomain" Height="20" Margin="464,102,499.449,0" TextWrapping="Wrap" VerticalAlignment="Top" Text="exampletenant"/>
        <TextBlock x:Name="txt_OnMicrososft_com" Margin="0,102,392.131,0" TextWrapping="Wrap" Height="19" VerticalAlignment="Top" HorizontalAlignment="Right" Width="102.318"><Run Text="."/><Run Text="onmicrosoft.com"/></TextBlock>
        <CheckBox x:Name="cbx_TenantModernAuth" Content="Requires Modern Auth (Multi Factor)" Height="17" Margin="319.793,129,405,0" VerticalAlignment="Top"/>
        <TextBlock x:Name="txt_TenantDisplayNameChanged" Height="19" Margin="0,32,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red"/>
        <TextBlock x:Name="txt_TenantSignInAddressChanged" Height="19" Margin="0,58,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red"/>
        <TextBlock x:Name="txt_TenantPasswordChanged" Height="19" Margin="0,82,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red"/>
        <TextBlock x:Name="txt_TenantRootDomainChanged" Height="19" Margin="0,106,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red"/>
        <TextBlock x:Name="txt_TenantModernAuthChanged" Height="19" Margin="0,127,378.991,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red"/>
        <Separator Height="14" Margin="310.223,151,378.991,0" VerticalAlignment="Top"/>
        <Grid x:Name="grd_ConnectionOptions" Margin="310.223,165,378.991,107">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="211*"/>
                <ColumnDefinition Width="180*"/>
            </Grid.ColumnDefinitions>
            <CheckBox x:Name="cbx_ConnectToTeams" Content="Connect to Microsoft Teams PowerShell" Margin="10,84.96,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToSkype" Content="Connect to Skype for Business Online PowerShell" Margin="10,125.156,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToAzureAD" Content="Connect to Azure AD PowerShell" Margin="10,24.666,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToAzureCompliance" Margin="10,44.764,10,0" Content="Connect to Azure Compliance Centre" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToSharepointOnline" Content="Connect to SharePoint Online PowerShell" Margin="10,105.058,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_ConnectToExchangeOnline" Content="Connect to Exchange Online PowerShell" Margin="10,64.862,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top"/>
            <TextBlock x:Name="txt_TenantServiceOptions" Height="19" Margin="2.579,0.666,57.421,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold"><Run Text="Tenant "/><Run Text="Service Connections"/></TextBlock>
            <TextBlock x:Name="txt_TenantConnectToAzureADChanged" Height="19" Margin="0,23.666,0,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1"/>
            <TextBlock x:Name="txt_TenantConnectToAzureComplianceChanged" Height="19" Margin="0,43.764,0,0" TextWrapping="Wrap" Text="*" VerticalAlignment="Top" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1"/>
            <TextBlock x:Name="txt_TenantConnectToExchangeOnlineChanged" Margin="0,67.01,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" VerticalAlignment="Top"/>
            <TextBlock x:Name="txt_TenantConnectToSharePointChanged" Margin="0,0,0,24.942" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1"/>
            <TextBlock x:Name="txt_TenantConnectToSkypeOnlineChanged" Margin="0,0,0,5.844" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1"/>
            <TextBlock x:Name="txt_TenantConnectToTeamsChanged" Margin="0,0,0,42.892" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1"/>
        </Grid>
        <Grid x:Name="grd_BounShellOptions" Margin="310.223,0,378.991,10" Height="92" VerticalAlignment="Bottom">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="105*"/>
                <ColumnDefinition Width="94*"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_BounShellTitle" Height="19" Margin="2.5,1,27.639,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold" Grid.ColumnSpan="2" Text="BounShell Specific Options"/>
            <CheckBox x:Name="cbx_BounShellCheckForPowerShellUpdates" Content="Check for PowerShell module updates" Margin="10,45.098,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top" IsChecked="True"/>
            <CheckBox x:Name="cbx_BounShellCheckForUpdates" Content="Check for BounShell updates" Margin="10,25,10,0" Grid.ColumnSpan="2" VerticalAlignment="Top" IsChecked="True"/>
            <CheckBox x:Name="cbx_BounShellEnableModernAuthClipboard" Content="Enable Modern Auth clipboard integration" Margin="10,65.196,-58.648,0" VerticalAlignment="Top" IsChecked="True"/>
            <TextBlock x:Name="txt_BounShellModernAuthLink" Margin="63.648,0,27.639,11.844" TextWrapping="Wrap" Grid.Column="1" Foreground="Blue" TextDecorations="Underline" VerticalAlignment="Bottom"><Run Text="Learn more."/><Run Text=".."/></TextBlock>
            <TextBlock x:Name="txt_BounShellCheckForPowerShellUpdatesChanged" Margin="0,24,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Top" Grid.Column="1"/>
            <TextBlock x:Name="txt_BounShellCheckForUpdatesChanged" Margin="0,0,0,25.754" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="19" VerticalAlignment="Bottom"/>
            <TextBlock x:Name="txt_BounShellEnableModernAuthClipboardChanged" Margin="0,0,0,5.656" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="19" VerticalAlignment="Bottom"/>
        </Grid>
        <Separator Margin="507.009,197,165.991,199" RenderTransformOrigin="0.5,0.5">
            <Separator.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="-90"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Separator.RenderTransform>
        </Separator>
        <Grid x:Name="grd_SkypeOptions" Margin="0,4.05,10,0" HorizontalAlignment="Right" Width="354.777" Height="136.336" VerticalAlignment="Top">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="105*"/>
                <ColumnDefinition Width="94*"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_SkypeTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFD447D4" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Skype for Business "/><Run Text="Online"/></TextBlock>
            <CheckBox x:Name="cbx_SkypeOverideDomain" Content="Admin Override Domain" Margin="10,0,5,10" Height="17" VerticalAlignment="Bottom"/>
            <CheckBox x:Name="cbx_SkypeAllowReconnect" Content="Allow Reconnect (Fixes ISE hard lockup)" Margin="10,45.335,10,0" IsChecked="True" Height="17" VerticalAlignment="Top" Grid.ColumnSpan="2"/>
            <CheckBox x:Name="cbx_SkypeOverideAdmin" Content="Override Discovery URI" Margin="10,0,10,32" Height="18.001" VerticalAlignment="Bottom"/>
            <TextBlock x:Name="txt_SkypeISEFixes" Height="19" Margin="10,25.335,55.639,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="SemiBold"><Run Text="ISE "/><Run Text="Bug Fixes"/></TextBlock>
            <TextBlock x:Name="txt_SkypeConnectionFix" Margin="10,62.335,20.139,55.001" TextWrapping="Wrap" FontWeight="SemiBold" Grid.ColumnSpan="2"><Run Text="BPOS / renamed "/><Run Text="/ hybrid "/><Run Text="tenant"/><Run Text=" and "/><Run Text="delegated admin"/><Run Text=" "/><Run Text="workarounds"/></TextBlock>
            <TextBox x:Name="tbx_SkypeDiscoveryUri" Margin="0,0,15.009,31" TextWrapping="Wrap" Text="fabrikam.onmicrosoft.com" Grid.Column="1" Height="21.001" VerticalAlignment="Bottom"/>
            <TextBox x:Name="tbx_SkypeAdminDomain" Margin="0,0,15.009,10" TextWrapping="Wrap" Height="20" VerticalAlignment="Bottom" Text="fabrikam.com" Grid.Column="1"/>
            <TextBlock x:Name="txt_SkypeDiscoveryURIChanged" Margin="0,0,0,30" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Bottom" RenderTransformOrigin="0.001,-2.169"/>
            <TextBlock x:Name="txt_SkypeAdminDomainChanged" Margin="0,0,0,11" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Height="19" VerticalAlignment="Bottom" Grid.Column="1"/>
            <TextBlock x:Name="txt_SkypeAllowReconnectChanged" Margin="0,45.335,-0.009,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top" RenderTransformOrigin="0.001,-2.169"/>
        </Grid>
        <Grid x:Name="grd_AzureADOptions" Margin="0,135,10,0" HorizontalAlignment="Right" Width="354.777" Height="52.336" VerticalAlignment="Top">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="105*"/>
                <ColumnDefinition Width="94*"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_AzureADTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFFF9623" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Azure AD"/></TextBlock>
            <CheckBox x:Name="cbx_AzureADEnviromentName" Content="Azure Environment Name" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_AzureADEnviromentName" Margin="0,25.335,15.009,0" TextWrapping="Wrap" Text="AzureCloud" Grid.Column="1" Height="21.001" VerticalAlignment="Top"/>
            <TextBlock x:Name="txt_AzureADEnviromentNameChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top"/>
        </Grid>
        <Grid x:Name="grd_AzureComplianceOptions" Margin="0,182.336,10,184.328" HorizontalAlignment="Right" Width="354.777">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_AzureComplianceTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FF29FCFF" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Azure "/><Run Text="Compliance"/></TextBlock>
            <CheckBox x:Name="cbx_AzureComplianceURI" Content="Connection Uri" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_AzureComplianceURI" Margin="0,25.335,15.009,0" Text="https://ps.compliance.protection.outlook.com/powershell-liveid/" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_AzureComplianceURIChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top"/>
        </Grid>
        <Grid x:Name="grd_ExchangeOnlineOptions" Margin="0,0,10,131.992" HorizontalAlignment="Right" Width="354.777" Height="52.336" VerticalAlignment="Bottom">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_ExchangeOnlineTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFFA9EDA" Background="Black"><Run Text=" "/><Run Text="■ "/><Run Text="Exchange Online"/><LineBreak/><Run/></TextBlock>
            <CheckBox x:Name="cbx_ExchangeOnlineURI" Content="Connection Uri" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_ExchangeOnlineURI" Margin="0,25.335,15.009,0" Text="https://outlook.Office365.com/powershell-liveid/" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_ExchangeOnlineURIChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top"/>
        </Grid>
        <Grid x:Name="grd_TeamsOptions" Margin="0,0,10,79.656" HorizontalAlignment="Right" Width="354.777" Height="52.336" VerticalAlignment="Bottom">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_TeamsTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="White" Background="Black"><Run Text=" "/><Run Text="■"/><Run Text=" Microsoft Teams"/></TextBlock>
            <CheckBox x:Name="cbx_TeamsEnviroment" Content="Teams Environment Name" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_TeamsEnviroment" Margin="10,25.335,15.009,0" Text="TeamsGCCH" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_TeamsEnviromentChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top"/>
        </Grid>
        <Grid x:Name="grd_SharePointOptions" Margin="0,0,10,11" HorizontalAlignment="Right" Width="354.777" Height="68.656" VerticalAlignment="Bottom">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="txt_SharepointTitle" Height="19" Margin="10,6.335,15.009,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" Grid.ColumnSpan="2" Foreground="#FFF9FF81" Background="Black"><Run Text=" "/><Run Text="■"/><Run Text=" "/><Run Text="SharePoint Online"/></TextBlock>
            <CheckBox x:Name="cbx_SharepointURL" Content="URL" Margin="10,27.335,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_SharepointURL" Margin="10,25.335,15.009,0" Text="https://contoso-admin.sharepoint.com" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_SharepointURLChanged" Margin="0,25.335,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top"/>
            <CheckBox x:Name="cbx_SharepointRegion" Content="Region" Margin="10,47.336,10,0" Height="18.001" VerticalAlignment="Top"/>
            <TextBox x:Name="tbx_SharepointRegion" Margin="10,45.336,15.009,0" Text="Default" Grid.Column="1" Height="21.001" VerticalAlignment="Top" MaxLines="1"/>
            <TextBlock x:Name="txt_SharepointRegionChanged" Margin="0,45.336,0,0" TextWrapping="Wrap" Text="*" HorizontalAlignment="Right" Width="10.009" FontWeight="Bold" Foreground="Red" Grid.Column="1" Height="20.001" VerticalAlignment="Top"/>
        </Grid>
    </Grid>
</Window>


'@

      $reader=(New-Object System.Xml.XmlNodeReader $xaml)
      $Global:syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
    
      [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
      [xml]$XAML = $xaml
      $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $Global:syncHash.Add($_.Name,$Global:syncHash.Window.FindName($_.Name) )

      }

      $Script:JobCleanup = [hashtable]::Synchronized(@{})
      $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

      #region Background runspace to clean up jobs
      $jobCleanup.Flag = $True
      $newRunspace =[runspacefactory]::CreateRunspace()
      $newRunspace.ApartmentState = "STA"
      $newRunspace.ThreadOptions = "ReuseThread"          
      $newRunspace.Open()        
      $newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)     
      $newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) 
      $jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
          #Routine to handle completed runspaces
          Do {    
            Foreach($runspace in $jobs) {            
              If ($runspace.Runspace.isCompleted) {
                [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                $runspace.powershell.dispose()
                $runspace.Runspace = $null
                $runspace.powershell = $null               
              } 
            }
            #Clean out unused runspace jobs
            $temphash = $jobs.clone()
            $temphash | Where {
              $_.runspace -eq $Null
            } | ForEach {
              $jobs.remove($_)
            }        
            Start-Sleep -Seconds 1     
          } while ($jobCleanup.Flag)
      })
      $jobCleanup.PowerShell.Runspace = $newRunspace
      $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
      #endregion Background runspace to clean up job
      

      #region Window Close 
      $Global:syncHash.Window.Add_Closed({
          Write-Verbose 'Halt runspace cleanup job processing'
          $jobCleanup.Flag = $False

          #Stop all runspaces
          $jobCleanup.PowerShell.Dispose()      
      })
      
      $Global:syncHash.Window.ShowDialog() | Out-Null
      $Global:syncHash.Error = $Error
  })
  $global:psCmd.Runspace = $newRunspace
 
}
Function Import-BsMultiThreadedWpfGuiFunctions 
{
  $Global:cbx_ConnectToTeams_Click = 
  {
    Read-BsConfigFile
    Hide-BsGuiElements
  }

  $Global:syncHash.cbx_ConnectToTeams.add_checked({
      Write-output "foo"
      Update-BsGuiWindow -Control cbx_ConnectToAzureAD -Property IsChecked -Value $true
  })

   
}

Function Show-BsMultiThreadWPFGuiElements
{
  #$global:WpfForm.ShowDialog() | out-null
  $data = $global:psCmd.BeginInvoke()
  
  Update-BsGuiWindow -Control cbx_ConnectToAzureAD -Property IsChecked -Value $false
  Update-BsGuiWindow -Control cbx_ConnectToAzureAD -Property IsChecked -Value $false
  
  $global:synchash.tbx_TenantRootDomain.ToString()
  $global:synchash.tbx_TenantRootDomain.ToString()
  
  
}

Function Show-BsWPFGuiElements
{
  #Thread will stick here until the user closes the window. Sucks but it works
  Write-Log -component $function -Message "Opening WPF Settings Window"
  $global:SettingsWpfWindow.ShowDialog() | out-null
  #Write-Log -component $function -Message "WPF Window closed, reloading elements"
  #I'm going to be honest, as I'm not Multithreading we are doing some cheating to make things "Work" after the user closes the WPF window.
  #Import-BsWpfGuiElements
  #Import-BsWpfGuiFunctions
  
}
Function Hide-BsWPFGuiElements
{
  Write-Log -component $function -Message "User closed WPF window, Checking for updates"
  $global:SettingsWpfWindow.Hide()
  
  If ($global:Config.AutoUpdatesEnabled)
  { 
    #Check for the required modules
    Write-Log -component $function -Message 'Checking for required modules based on selections, this can take some time.' -severity 2
    #Teams Module Check
    if ($global:Config.Tenant1.ConnectToTeams -or $global:Config.Tenant2.ConnectToTeams -or $global:Config.Tenant3.ConnectToTeams -or $global:Config.Tenant4.ConnectToTeams -or $global:Config.Tenant5.ConnectToTeams -or $global:Config.Tenant6.ConnectToTeams -or $global:Config.Tenant7.ConnectToTeams -or $global:Config.Tenant8.ConnectToTeams -or $global:Config.Tenant9.ConnectToTeams -or $global:Config.Tenant10.ConnectToTeams)
    {
      Test-BsInstalledModules -ModuleName $TestedTeamsModule -ModuleVersion $TestedTeamsModuleVer
    }

    #Exchange Module Check
    if ($global:Config.Tenant1.ConnectToExchange -or $global:Config.Tenant2.ConnectToExchange -or $global:Config.Tenant3.ConnectToExchange -or $global:Config.Tenant4.ConnectToExchange -or $global:Config.Tenant5.ConnectToExchange -or $global:Config.Tenant6.ConnectToExchange -or $global:Config.Tenant7.ConnectToExchange -or $global:Config.Tenant8.ConnectToExchange -or $global:Config.Tenant9.ConnectToExchange -or $global:Config.Tenant10.ConnectToExchange)
    {
      Test-BsInstalledModules -ModuleName $TestedExchangeModule -ModuleVersion $TestedExchangeModuleVer
    }
    
    #MsOnline Module Check
    if ($global:Config.Tenant1.ConnectToAzureAD -or $global:Config.Tenant2.ConnectToAzureAD -or $global:Config.Tenant3.ConnectToAzureAD -or $global:Config.Tenant4.ConnectToAzureAD -or $global:Config.Tenant5.ConnectToAzureAD -or $global:Config.Tenant6.ConnectToAzureAD -or $global:Config.Tenant7.ConnectToAzureAD -or $global:Config.Tenant8.ConnectToAzureAD -or $global:Config.Tenant9.ConnectToAzureAD -or $global:Config.Tenant10.ConnectToAzureAD)
    {
      Test-BsInstalledModules -ModuleName $TestedMSOnlineModule -ModuleVersion $TestedMSOnlineModuleVer
    }
    
    #Skype4B Module Check
    if ($global:Config.Tenant1.ConnectToSkype -or $global:Config.Tenant2.ConnectToSkype -or $global:Config.Tenant3.ConnectToSkype -or $global:Config.Tenant4.ConnectToSkype -or $global:Config.Tenant5.ConnectToSkype -or $global:Config.Tenant6.ConnectToSkype -or $global:Config.Tenant7.ConnectToSkype -or $global:Config.Tenant8.ConnectToSkype -or $global:Config.Tenant9.ConnectToSkype -or $global:Config.Tenant10.ConnectToSkype)
    {
      Test-BsInstalledModules -ModuleName $TestedSkype4BOModule -ModuleVersion $TestedSkype4BOModuleVer
    }
    
    Write-Log -component $function -Message 'Module check complete' -severity 2
  }
  
 }
 

Function Update-BsGuiWindow {
  Param (
    $Control,
    $Property,
    $Value,
    [switch]$AppendContent
  )

  # This is kind of a hack, there may be a better way to do this
  If ($Property -eq "Close") {
    $syncHash.Window.Dispatcher.invoke([action]{$syncHash.Window.Close()},"Normal")
    Return
  }

  # This updates the control based on the parameters passed to the function
  $syncHash.$Control.Dispatcher.Invoke([action]{
      # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
      If ($PSBoundParameters['AppendContent']) {
        $syncHash.$Control.AppendText($Value)
      } Else {
        $syncHash.$Control.$Property = $Value
      }
  }, "Normal")
}       
#now we export the relevant stuff

Export-ModuleMember -Function Read-BsConfigFile
Export-ModuleMember -Function Write-BsConfigFile
Export-ModuleMember -Function Import-BsDefaultConfig
Export-ModuleMember -Function Invoke-BsNewTenantTab
Export-ModuleMember -Function Connect-BsO365Tenant
Export-ModuleMember -Function Update-BsAddonMenu
Export-ModuleMember -Function Import-BsGuiElements
Export-ModuleMember -Function Import-BsGuiFunctions
Export-ModuleMember -Function Show-BsGuiElements
Export-ModuleMember -Function Hide-BsGuiElements
Export-ModuleMember -Function Start-BounShell
Export-ModuleMember -Function Watch-BsCredentials
Export-ModuleMember -Function Test-BsInstalledModules
Export-ModuleMember -Function Repair-BsInstalledModules
Export-ModuleMember -Function Import-BsWpfGuiElements
Export-ModuleMember -Function Import-BsWpfGuiFunctions
Export-ModuleMember -Function Show-BsWpfGuiElements
Export-ModuleMember -Function Hide-BsWpfGuiElements
Export-ModuleMember -Function Update-BsGuiWindow




# SIG # Begin signature block
  # MIINFwYJKoZIhvcNAQcCoIINCDCCDQQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
  # gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
  # AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU8iX7pOjCHQvrnHHD5ZO8lruX
  # Q3ygggpZMIIFITCCBAmgAwIBAgIQDoW3bt/ALpa0ONbdsxRpGjANBgkqhkiG9w0B
  # AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
  # VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
  # c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE5MDkyNTAwMDAwMFoXDTIyMDky
  # ODEyMDAwMFowXjELMAkGA1UEBhMCQVUxETAPBgNVBAgTCFZpY3RvcmlhMRAwDgYD
  # VQQHEwdCZXJ3aWNrMRQwEgYDVQQKEwtKYW1lcyBBcmJlcjEUMBIGA1UEAxMLSmFt
  # ZXMgQXJiZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD1k72Yov0l
  # oWfUocvwkcKMQ8XpX6rQX6sPS1PkkN3wu2w75MsF4KvpaBCP2+rS+St7q+6Mk3l+
  # l1m0hd0J18Y05jbC1TmYvYLkdpsVhVxhjQtNBYGYFao3rpcQU4apx1F/PaPhCkIe
  # oL2mBU8VFj4aR4ZuplGAY+jlJqDXNlMlEmqkHlDkoJt6AYqJlu8ObaMmNQWChemM
  # HoxLJpGmsU+13PTaZoxLg4qDrGD5AfG9AuRUCrZ+gtMaX1xT8OAia56e3pZeV8oz
  # Pjbyl6mVRO8hwAWgK9eaeQ2pSgxa1KjazrkqyDrd8SvcWCQjj0vVrHeHv+hkaeWR
  # hn2Fk5ViasXxAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg
  # +S32ZXUOWDAdBgNVHQ4EFgQUUqun5eFswlws49GKcNHr5AaDtXowDgYDVR0PAQH/
  # BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0
  # dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWg
  # M6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcx
  # LmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRw
  # czovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEE
  # eDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYB
  # BQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJB
  # c3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3
  # DQEBCwUAA4IBAQBeD1O9ZawxpsQCZLJMLCBmpK8VQNfuljVT9LCcSKKjSsGDwMtq
  # m8IzqON27XMnbsEVqFvGuk83jpleLFEx6/UivDcOZkisQ7nKM4VsFDEiw7DbfnID
  # 0awTzjjo4pn4Vp1NdOClfNvpfcroZkq8IaBn1TXCyCXir3amMVUM/gR+6mVrAB4y
  # T23T2jWJNPzyFCOPuj8cCNQdBWBM4Lzt11LU59swV4Par2FuQQMia1jNpAbunT/9
  # bXZuAEmVWD1ra6cp6+9APNnZFk4UvIqj3yVrsKrJukn+uNApOgngnkBLaFy3VlFj
  # F1jt5QXptaMVWfBtxweCXvRHN6Aju4bquUD8MIIFMDCCBBigAwIBAgIQBAkYG1/V
  # u2Z1U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UE
  # ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYD
  # VQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAw
  # WhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
  # cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdp
  # Q2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG
  # 9w0BAQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/
  # 5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH
  # 03sjlOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxK
  # hwjfDPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr
  # /mzLfnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi
  # 6CxR93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCC
  # AckwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAww
  # CgYIKwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8v
  # b2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRp
  # Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6
  # MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
  # ZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
  # aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1s
  # AAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMw
  # CgYIYIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1Ud
  # IwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+
  # 7A1aJLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbR
  # knUPUbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7
  # uq+1UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7
  # qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPa
  # s7CM1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR
  # 6mhsRDKyZqHnGKSaZFHvMYICKDCCAiQCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTAT
  # BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
  # MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
  # QQIQDoW3bt/ALpa0ONbdsxRpGjAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEK
  # MAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
  # AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUEWbtLrZ3F4Qe7rrY
  # lxGRebzpSmYwDQYJKoZIhvcNAQEBBQAEggEAyLS0/jzh/uAQIqYBJJC67+eefySd
  # zI+X9dlmip8Lg5HZbNPwz5fwacOmgkjD1tF3l5LtHqaWCMSkuwlGGWm7cBXAoTjs
  # jqGL00funtH1IySigfz3aUIukstdshQQyTrY4QVZrBWqxzOPVtKUY2dS56odm/rK
  # HQcRY/IrY0HmxVye/OjNSKL8U77sdokQlFMUptgrLcv8DiWFEkSzZHMoyZ536n0a
  # kBpEWRPhaxlpfmo45GG+OslVPF8tchg+ZutdAXKuKY/JmYcWKWCPj6c8Lh7XUAO5
  # z1cBJ9iwoKvn71FIhf3T7yGE9MGR5YjQdSgK9gMNIuXDI0fnIO7KUzr2zQ==
# SIG # End signature block
