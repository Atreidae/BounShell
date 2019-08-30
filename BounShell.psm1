<#
    .SYNOPSIS

    This is a tool to help users manage multiple office 365 tenants

    .DESCRIPTION

    Created by James Arber. www.UcMadScientist.com
    
    .NOTES

    Version                : 0.6.6
    Date                   : 24/08/2019
    Lync Version           : Tested against Skype4B 2015
    Author                 : James Arber
    Header stolen from     : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"
    Special Thanks to      : My Beta Testers. Greig Sheridan, Pat Richard and Justin O'Meara
    
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
    to the Lync and Skype4B community AS IS without any warranty on its appropriateness for use in
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
[String]$ScriptVersion              = '0.6.6'
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
[String]$TestedSharepointModule     = 'Microsoft.Online.Sharepoint.PowerShell' #Not used yet
[String]$TestedSharepointModuleVer  = '16.0.8812.1200' #Not used yet
[String]$TestedAzureADModule        = 'AzureAD' #Not used yet
[String]$TestedAzureADModuleVer     = '2.0.2.16' #Not used yet
[String]$TestedAzureADRMModule      = 'AADRM' #Not used yet
[String]$TestedAzureADRMModuleVer   = '2.13.1.0' #Not used yet
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
    #If the log entry has a severity of 3 assume its a warning and write it to write-warning
    if ($Severity -eq 3) 
    {
      "$Date $Message"| Write-Warning
    }
    #If the log entry has a severity of 4 or higher, assume its an error and display an error message (Note, critical errors are caught by throw statements so may not appear here)
    if ($Severity -ge 4) 
    {
      "$Date $Message"| Write-Error
    }
  }
}

Function Get-IEProxy
{
  Write-Log -component $function -Message 'Checking for Proxy' -severity 2
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
  $GitHubScriptVersion = Invoke-WebRequest -Uri "https://raw.githubusercontent.com/atreidae/$GithubRepo/$GithubBranch/version" -TimeoutSec 10 -Proxy $ProxyURL #todo change back to master!
  
  If ($GitHubScriptVersion.Content.length -eq 0) 
  {
    #Empty data, throw an error
    Write-Log -component $function -Message 'Error checking for new version. You can check manualy here' -severity 3
    Write-Log -component $function -Message $BlogPost -severity 1 #Todo Update URL
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

    if ([single]$splitgitver[1] -gt [single]$splitver[1])
    {
      $Needupdate = $true
      #New Major Build available, #Prompt user to download
      Write-Log -component $function -Message 'New Minor Version Available' -severity 3
      $title = 'Update Available'
      $Message = 'a minor update to this script is available, did you want to download it?'
    }

    if ([single]$splitgitver[2] -gt [single]$splitver[2])
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
      'Updates the installed PowerShell Module'

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
    
    #We alreday have the lastest version
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
    Write-Log -component $function -Message "Sorry, something went wrong here and I couldnt backup your BounShell config. Please check permissions to create $ENV:UserProfile\BounShell-backup-$ShrtDate.xml" -severity 3
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
  
  Try
  {
    #Load the Config
    $global:Config = @{}
    $global:Config = (Import-Clixml -Path $global:ConfigFilePath)
    Write-Log -component $function -Message 'Config File Read OK' -severity 2
    
    #Check the config file version
    If ($global:Config.ConfigFileVersion -lt 0.2)
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
  }
    
  Catch
  {
    #For some reason we ran into an issue updating variables, throw and error and revert to defaults
    Write-Log -component $function -Message 'Error reading Config or updating GUI, Loading Defaults' -severity 3
    Import-BsDefaultConfig
  }
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
  #Remove and re-create the Config Array
  $null = (Remove-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)
  $global:Config = @{}
  #Populate with Defaults
  $null =  $Global:grid_Tenants.Rows.Clear()
  
  $global:Config.Tenant1 = @{}
  $global:Config.Tenant1.DisplayName = 'Undefined'
  $global:Config.Tenant1.SignInAddress = 'user1@fabrikam.com'
  $global:Config.Tenant1.Credential = '****'
  $global:Config.Tenant1.ModernAuth = $false
  $global:Config.Tenant1.ConnectToTeams = $false
  $global:Config.Tenant1.ConnectToSkype = $false
  $global:Config.Tenant1.ConnectToExchange = $false
  $global:Config.Tenant1.ConnectToAzureAD = $false
  $global:Config.Tenant1.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('1','Undefined','user1@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  
  $global:Config.Tenant2 = @{}
  $global:Config.Tenant2.DisplayName = 'Undefined'
  $global:Config.Tenant2.SignInAddress = 'user2@fabrikam.com'
  $global:Config.Tenant2.Credential = '****'
  $global:Config.Tenant2.ModernAuth = $false
  $global:Config.Tenant2.ConnectToTeams = $false
  $global:Config.Tenant2.ConnectToSkype = $false
  $global:Config.Tenant2.ConnectToExchange = $false
  $global:Config.Tenant2.ConnectToAzureAD = $false
  $global:Config.Tenant2.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('2','Undefined','user2@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  $global:Config.Tenant3 = @{}
  $global:Config.Tenant3.DisplayName = 'Undefined'
  $global:Config.Tenant3.SignInAddress = 'user3@fabrikam.com'
  $global:Config.Tenant3.Credential = '****'
  $global:Config.Tenant3.ModernAuth = $false
  $global:Config.Tenant3.ConnectToTeams = $false
  $global:Config.Tenant3.ConnectToSkype = $false
  $global:Config.Tenant3.ConnectToExchange = $false
  $global:Config.Tenant3.ConnectToAzureAD = $false
  $global:Config.Tenant3.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('3','Undefined','user3@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  $global:Config.Tenant4 = @{}
  $global:Config.Tenant4.DisplayName = 'Undefined'
  $global:Config.Tenant4.SignInAddress = 'user4@fabrikam.com'
  $global:Config.Tenant4.Credential = '****'
  $global:Config.Tenant4.ModernAuth = $false
  $global:Config.Tenant4.ConnectToTeams = $false
  $global:Config.Tenant4.ConnectToSkype = $false
  $global:Config.Tenant4.ConnectToExchange = $false
  $global:Config.Tenant4.ConnectToAzureAD = $false
  $global:Config.Tenant4.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('4','Undefined','user4@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  $global:Config.Tenant5 = @{}
  $global:Config.Tenant5.DisplayName = 'Undefined'
  $global:Config.Tenant5.SignInAddress = 'user5@fabrikam.com'
  $global:Config.Tenant5.Credential = '****'
  $global:Config.Tenant5.ModernAuth = $false
  $global:Config.Tenant5.ConnectToTeams = $false
  $global:Config.Tenant5.ConnectToSkype = $false
  $global:Config.Tenant5.ConnectToExchange = $false
  $global:Config.Tenant5.ConnectToAzureAD = $false
  $global:Config.Tenant5.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('5','Undefined','user5@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  $global:Config.Tenant6 = @{}
  $global:Config.Tenant6.DisplayName = 'Undefined'
  $global:Config.Tenant6.SignInAddress = 'user6@fabrikam.com'
  $global:Config.Tenant6.Credential = '****'
  $global:Config.Tenant6.ModernAuth = $false
  $global:Config.Tenant6.ConnectToTeams = $false
  $global:Config.Tenant6.ConnectToSkype = $false
  $global:Config.Tenant6.ConnectToExchange = $false
  $global:Config.Tenant6.ConnectToAzureAD = $false
  $global:Config.Tenant6.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('6','Undefined','user6@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  $global:Config.Tenant7 = @{}
  $global:Config.Tenant7.DisplayName = 'Undefined'
  $global:Config.Tenant7.SignInAddress = 'user@fabrikam.com'
  $global:Config.Tenant7.Credential = '****'
  $global:Config.Tenant7.ModernAuth = $false
  $global:Config.Tenant7.ConnectToTeams = $false
  $global:Config.Tenant7.ConnectToSkype = $false
  $global:Config.Tenant7.ConnectToExchange = $false
  $global:Config.Tenant7.ConnectToAzureAD = $false
  $global:Config.Tenant7.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('7','Undefined','user7@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  $global:Config.Tenant8 = @{}
  $global:Config.Tenant8.DisplayName = 'Undefined'
  $global:Config.Tenant8.SignInAddress = 'user8@fabrikam.com'
  $global:Config.Tenant8.Credential = '****'
  $global:Config.Tenant8.ModernAuth = $false
  $global:Config.Tenant8.ConnectToTeams = $false
  $global:Config.Tenant8.ConnectToSkype = $false
  $global:Config.Tenant8.ConnectToExchange = $false
  $global:Config.Tenant8.ConnectToAzureAD = $false
  $global:Config.Tenant8.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('8','Undefined','user8@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
    
  $global:Config.Tenant9 = @{}
  $global:Config.Tenant9.DisplayName = 'Undefined'
  $global:Config.Tenant9.SignInAddress = 'user@fabrikam.com'
  $global:Config.Tenant9.Credential = '****'
  $global:Config.Tenant9.ModernAuth = $false
  $global:Config.Tenant9.ConnectToTeams = $false
  $global:Config.Tenant9.ConnectToSkype = $false
  $global:Config.Tenant9.ConnectToExchange = $false
  $global:Config.Tenant9.ConnectToAzureAD = $false
  $global:Config.Tenant9.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('9','Undefined','user9@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  $global:Config.Tenant10 = @{}
  $global:Config.Tenant10.DisplayName = 'Undefined'
  $global:Config.Tenant10.SignInAddress = 'user@fabrikam.com'
  $global:Config.Tenant10.Credential = '****'
  $global:Config.Tenant10.ModernAuth = $false
  $global:Config.Tenant10.ConnectToTeams = $false
  $global:Config.Tenant10.ConnectToSkype = $false
  $global:Config.Tenant10.ConnectToExchange = $false
  $global:Config.Tenant10.ConnectToAzureAD = $false
  $global:Config.Tenant10.ConnectToCompliance = $false
  $null =  $Global:grid_Tenants.Rows.Add('10','Undefined','user10@fabrikam.com','****',$false,$false,$false,$false,$false,$false)
  
  [Float]$global:Config.ConfigFileVersion = '0.2'
  [string]$global:Config.Description = 'BounShell Configuration file. See UcMadScientist.com/BounShell for more information'
  
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
      Write-Log -component $function -Message 'Failed to open new tab, Is there already a connection open to that tenant?' -severity 3
    }
  }
  
  if($Tabname -eq 'Undefined')
  {
    #Tabname is "undefined", user clicked a tenant thats not confgured yay
    Write-Log -component $function -Message "Sorry, I cant find a config for Tenant $Tenant" -severity 3
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
    #Clean up any stale sessions (we shouldnt have any, but whatever)
    Get-PSSession | Remove-PSSession
  }
  #load the gui stuff for configuration #todo, put  check here and only load it if its not loaded.
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


    #First figure out the maximum width of the items name (for the tabular menu):
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
        ($global:StoredPsCred).Password.MakeReadOnly() #Thanks for spotting this Grieg!
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
    11 #Tenant 11 should never happen. this is refactoring for 0.3 config file with array support, its here for testing.
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
  
  #check to see if the Modern Auth flag has been set and use the appropriate connection method
  If ($ModernAuth) 
  {
    # We are using Modern Auth, Check to see if the user accepted the warning. if not. Prompt them
    If ($global:Config.ModernAuthWarningAccepted -eq $false) 
    { 


      #We should only warn them if the feature is actually on.
      If ($global:Config.ModernAuthClipboardEnabled = $true)
      {
        Write-Log -Message 'User hasnt accepted the Modern Auth disclaimer, prompt' -Severity 1 -Component $function
        Write-Host -Object 'Modern Auth Clipboard intergration is currently enabled'
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
      Write-Host -Object 'You can paste it into the Modern Auth Window using CTRL+V to speed up the sign in process,'
      Write-Host -Object 'then paste your password using CTRL+V and sign in'
      Write-Host -Object 'Upon pasting this, BounShell will clear the clipboard and overwrite the memory just incase.'
    }
    Else
    {
      Write-Host -Object 'Your Username has been copied to the clipboard'
      Write-Host -Object 'You can paste it into the Modern Auth Window using CTRL+V to speed up the sign in process'
      Write-Host -Object 'You will need to enter your password manually. You can enable password support in settings'
    }
    #As we are dealing with modern auth we need to convert the password back to an insecure string do that here
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ModernAuthPassword)
    $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
    
    If ($ConnectToTeams) 
    {
      Try
      {
        # So Now we need to kick off a new window that waits for the clipboard events
        Write-Log -Message 'Connecting to Microsoft Teams' -Severity 2 -Component $function
        #Create a script block with the expanded variables
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell convers the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process powershell $sb

        #Sleep for a few seconds to let the powershell window pop and fill the clipboard.
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
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell convers the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process powershell $sb

        #Sleep for a few seconds to let the powershell window pop and fill the clipboard.
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
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell convers the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process powershell $sb

        #Sleep for a few seconds to let the powershell window pop and fill the clipboard.
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
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell convers the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process powershell $sb

        #Sleep for a few seconds to let the powershell window pop and fill the clipboard.
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
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell convers the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process powershell $sb

        #Sleep for a few seconds to let the powershell window pop and fill the clipboard.
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
        if ($global:Config.ModernAuthClipboardEnabled -eq $true) #workaround a bug where PowerShell convers the bool to a string and cant convert back
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword -NoPassword"
        }
        Else
        {
          [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        }
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-Process powershell $sb

        #Sleep for a few seconds to let the powershell window pop and fill the clipboard.
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
  #If the modern auth flag hasnt been set, we can simply connect to the services using secure credentials
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
        $O365Session = (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $pscred -Authentication Basic -AllowRedirection )
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
        $VerbosePreference = 'SilentlyContinue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $S4BOSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = 'Continue' #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch
      {
        #We had an issues connecting to Skype
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
      #Flag is set, connect to Compiance Centre
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
  #Check to see if we are loaded, if we are cleanup after ourselves
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

  #Need to put a for each code in here that adds Tenant 1 through 10
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
  $Global:cbx_AutoUpdates.Text = [System.String]'Automatically Check For Updates'
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
  $Global:Tenant_Email.HeaderText = [System.String]'Sign In Address'
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
  $Global:cbx_ClipboardAuth.Text = [System.String]'Enable Modern Auth Clipboard Intergration'
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
  $Global:cliplabel.Text = [System.String]'more info.'
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

Function Start-BounShell
{
  $function = 'Start-BounShell'
  #Allows us to seperate all the "onetime" run objects incase we get dot sourced.
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

    Write-Log -component $function -Message 'As we didnt find a config file we will assume this is a first run.' -severity 3
    Write-Log -component $function -Message 'Thus we will remind you that while all care is taken to store your credentials in a safe manner, we cannot be held responsible for any data breaches' -severity 3
    Write-Log -component $function -Message 'If someone was to get a hold of your BounShell.xml AND your user profile private encryption key its possible to reverse engineer stored credentials' -severity 3
    Write-Log -component $function -Message 'Seriously, Whilst the password store is encrypted, its not perfect!' -severity 3
    Pause
  }

  #check for script update
  if ($SkipUpdateCheck -eq $false)
  {
    Get-ScriptUpdate
  } #todo enable update checking

  #Check for Modules
  #Test-ManagementTools #todo fix
  
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
    Write-Log -component $function -Message 'Launched BounShell without ISE Support, Keyboard Shotcuts will be unavailable' -severity 2
    Write-Host -Object ''
    Write-Log -component $function -Message 'To configure BounShell tenants run Show-BsGuiElements' -severity 2
    Write-Log -component $function -Message 'To connect to a tenant run Connect-BsO365Tenant'  -severity 2
    Return #Yes I know Return sucks, but I think its better than Throw.
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
    [Parameter(Mandatory)] [string]$ModuleVersion
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
    $Message = "I've detected multiple installs of $ModuleName, Should I remove them and install the latest version?"
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
   
  #If we have one module, check its up to date
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
    $Message = "$ModuleName is not installed on this computer, it is required to connect to services. Can I install it for you?"
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
      #Todo What the fuck?
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
    # throw 'Please Note: You are trying to run this script without evalated Administator Priviliges. In order to run this script you will required Powershell running in Administrator Mode'
    Write-Log -component $function -Message 'Not Running as administrator, invoking new session' -severity 2
    $newProcess = New-Object -TypeName System.Diagnostics.ProcessStartInfo -ArgumentList 'PowerShell'
   
    # Specify the current script path and name as a parameter
    $newProcess.Arguments = "Repair-BsInstalledModules -modulename $ModuleName -operation $Operation"
   
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


Function Import-BsGui2Elements
{

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

