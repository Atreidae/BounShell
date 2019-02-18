<#
    .SYNOPSIS

    This is a tool to help users manage multiple office 365 tenants

    .DESCRIPTION

    Created by James Arber. www.UcMadScientist.com
    
    .NOTES

    Version                : 0.6
    Date                   : 10/02/2019
    Lync Version           : Tested against Skype4B 2015
    Author                 : James Arber
    Header stolen from     : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"
    Special Thanks to      : My Beta Testers. Greig Sheridan, Pat Richard and Justin O'Meara

    v0.6: Beta Release
    Enabled Modern Auth Support
    Formating changes
    Broke up alot of my one-liners to make it easier for others to read/ understand the flow
    Updated error messages
    Better code comments
    Fixed an issue with the Compliance Portal code
    Now Gluten Free

    :v0.5: Beta Release

    .LINK
    https://www.UcMadScientist.com

    .KNOWN ISSUES
    Beta, Buggy as all get out.

    .EXAMPLE
    Loads the Module
    PS C:\> Start-BounShell.ps1

#>

[CmdletBinding(DefaultParametersetName="Common")]
param
(
  [Parameter(Mandatory=$false)] [switch]$SkipUpdateCheck,
  [Parameter(Mandatory=$false)] [String]$ConfigFilePath = $null,
  [Parameter(Mandatory=$false)] [String]$LogFileLocation = $null,
  [Parameter(Mandatory=$false)] [float]$Tenant = $null

)

#region config
[Net.ServicePointManager]::SecurityProtocol = 'tls12, tls11, tls'
$StartTime                  =  Get-Date
$VerbosePreference          =  "SilentlyContinue" #TODO
[float]$ScriptVersion       =  '0.6'
[string]$GithubRepo         =  'BounShell' ##todo
[string]$GithubBranch       =  'devel' #todo
[string]$BlogPost           =  'https://www.UcMadScientist.com/BounShell/' #todo

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
         
  "$env:ComputerName date=$([char]34)$Date2$([char]34) time=$([char]34)$Date$([char]34) component=$([char]34)$component$([char]34) type=$([char]34)$severity$([char]34) Message=$([char]34)$Message$([char]34)"| Out-File -FilePath $Path -Append -NoClobber -Encoding default
  If (!$LogOnly) 
  {
    #If LogOnly is set, we dont want to write anything to the screen as we are capturing data that might look bad onscreen
      
      
    #If the log entry is just Verbose (1), output it to verbose
    if ($severity -eq 1) 
    {
      "$Date $Message"| Write-Verbose
    }
      
    #If the log entry is just informational (2), output it to write-host
    if ($severity -eq 2) 
    {
      "Info: $Date $Message"| Write-Host -ForegroundColor Green
    }
    #If the log entry has a severity of 3 assume its a warning and write it to write-warning
    if ($severity -eq 3) 
    {
      "$Date $Message"| Write-Warning
    }
    #If the log entry has a severity of 4 or higher, assume its an error and display an error message (Note, critical errors are caught by throw statements so may not appear here)
    if ($severity -ge 4) 
    {
      "$Date $Message"| Write-Error
    }
  }
}

Function Get-IEProxy
{
  Write-Host "Info: Checking for proxy settings" -ForegroundColor Green
  If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0)
  {
    $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
    if ($proxies) 
    {
      if ($proxies -ilike "*=*")
      {
        return $proxies -replace "=", "://" -split (';') | Select-Object -First 1
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
  $Function = 'Get-ScriptUpdate'
  Write-Log -component $function -Message "Checking for Script Update" -severity 1
  Write-Log -component $function -Message "Checking for Proxy" -severity 1
  $ProxyURL = Get-IEProxy
  
  If ($ProxyURL)
  
  {
    Write-Log -component $function -Message "Using proxy address $ProxyURL" -severity 1
  }
  
  Else
  {
    Write-Log -component $function -Message "No proxy setting detected, using direct connection" -severity 1
  }

  $GitHubScriptVersion = Invoke-WebRequest https://raw.githubusercontent.com/atreidae/BounShell/devel/version -TimeoutSec 10 -Proxy $ProxyURL #todo change back to master!
  
  If ($GitHubScriptVersion.Content.length -eq 0) 
  {
    #Empty data, throw an error
    Write-Log -component $function -Message "Error checking for new version. You can check manualy here" -severity 3
    Write-Log -component $function -Message $BlogPost -severity 1 #Todo Update URL
    Write-Log -component $function -Message "Pausing for 5 seconds" -severity 1
    start-sleep 5
  }
  else
  {
    #Process the returned data
    if ([single]$GitHubScriptVersion.Content -gt [single]$ScriptVersion)
    {
      #New Version available, #Prompt user to download
      Write-Log -component $function -Message "New Version Available" -severity 3
      $title = "Update Available"
      $message = "an update to this script is available, did you want to download it?"

      $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
      "Launches a browser window with the update"

      $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
      "No thanks."

      $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

      $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

      switch ($result)
      {
        0 {
          #User said yes
          Write-Log -component $function -Message "User opted to download update" -severity 1
          Start $BlogPost
          Write-Log -component $function -Message "Exiting Script" -severity 3
          Exit
        }
        #User said no
        1 {Write-Log -component $function -Message "User opted to skip update" -severity 1
        }
      }
    }
    
    #We alreday have the lastest version
    Else
    {
      Write-Log -component $function -Message "Script is upto date" -severity 1
    }
  }
}

Function Read-BsConfigFile
{
  $Function= 'Read-BsConfigFile'
  Write-Log -component $function
  Write-Log -component $function -Message "Reading Config file $($global:ConfigFilePath)" -severity 2
  If(!(Test-Path $global:ConfigFilePath)) 
  
  {
    #Cant locate test file, Throw error
    Write-Log -component $function -Message "Could not locate config file!" -severity 3
    Write-Log -component $function -Message "Error reading Config, Loading Defaults" -severity 3
    Import-BsDefaultConfig
  }
  Else
  {
    #Found the config file
    Write-Log -component $function -Message "Found Config file in the specified folder" -severity 1
  }

  Write-Log -component $function -Message "Pulling XML data" -severity 1
  [Void](Remove-Variable -Name Config -Scope Script -ErrorAction SilentlyContinue )
  Try
  {
    #Load the Config
    $global:Config=@{}
    $global:Config = (Import-CliXml -Path $global:ConfigFilePath)
    Write-Log -component $function -Message "Config File Read OK" -severity 2
    

    #Update the Gui options if we are loaded in the ISE
    If($PSISE) 
    {

      #Update the PS ISE Addon Menu
      Update-BsAddonMenu

      #Populate with Values
      [void] $Global:grid_Tenants.Rows.Clear()
      [void] $Global:grid_Tenants.Rows.Add("1",$global:Config.Tenant1.DisplayName,$global:Config.Tenant1.SignInAddress,"****",$global:Config.Tenant1.ModernAuth,$global:Config.Tenant1.ConnectToTeams,$global:Config.Tenant1.ConnectToSkype,$global:Config.Tenant1.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("2",$global:Config.Tenant2.DisplayName,$global:Config.Tenant2.SignInAddress,"****",$global:Config.Tenant2.ModernAuth,$global:Config.Tenant2.ConnectToTeams,$global:Config.Tenant2.ConnectToSkype,$global:Config.Tenant2.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("3",$global:Config.Tenant3.DisplayName,$global:Config.Tenant3.SignInAddress,"****",$global:Config.Tenant3.ModernAuth,$global:Config.Tenant3.ConnectToTeams,$global:Config.Tenant3.ConnectToSkype,$global:Config.Tenant3.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("4",$global:Config.Tenant4.DisplayName,$global:Config.Tenant4.SignInAddress,"****",$global:Config.Tenant4.ModernAuth,$global:Config.Tenant4.ConnectToTeams,$global:Config.Tenant4.ConnectToSkype,$global:Config.Tenant4.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("5",$global:Config.Tenant5.DisplayName,$global:Config.Tenant5.SignInAddress,"****",$global:Config.Tenant5.ModernAuth,$global:Config.Tenant5.ConnectToTeams,$global:Config.Tenant5.ConnectToSkype,$global:Config.Tenant5.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("6",$global:Config.Tenant6.DisplayName,$global:Config.Tenant6.SignInAddress,"****",$global:Config.Tenant6.ModernAuth,$global:Config.Tenant6.ConnectToTeams,$global:Config.Tenant6.ConnectToSkype,$global:Config.Tenant6.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("7",$global:Config.Tenant7.DisplayName,$global:Config.Tenant7.SignInAddress,"****",$global:Config.Tenant7.ModernAuth,$global:Config.Tenant7.ConnectToTeams,$global:Config.Tenant7.ConnectToSkype,$global:Config.Tenant7.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("8",$global:Config.Tenant8.DisplayName,$global:Config.Tenant8.SignInAddress,"****",$global:Config.Tenant8.ModernAuth,$global:Config.Tenant8.ConnectToTeams,$global:Config.Tenant8.ConnectToSkype,$global:Config.Tenant8.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("9",$global:Config.Tenant9.DisplayName,$global:Config.Tenant9.SignInAddress,"****",$global:Config.Tenant9.ModernAuth,$global:Config.Tenant9.ConnectToTeams,$global:Config.Tenant9.ConnectToSkype,$global:Config.Tenant9.ConnectToExchange)
      [void] $Global:grid_Tenants.Rows.Add("10",$global:Config.Tenant10.DisplayName,$global:Config.Tenant10.SignInAddress,"****",$global:Config.Tenant10.ModernAuth,$global:Config.Tenant10.ConnectToTeams,$global:Config.Tenant10.ConnectToSkype,$global:Config.Tenant10.ConnectToExchange)
    }
  }
    
  Catch
  {
    #For some reason we ran into an issue updating variables, throw and error and revert to defaults
    Write-Log -component $function -Message "Error reading Config or updating GUI, Loading Defaults" -severity 3
    Import-BsDefaultConfig
  }
}

Function Write-BsConfigFile
{
  $function = 'Write-BsConfigFile'
  Write-Log -component $function -Message "Writing Config file" -severity 2
  
  #Grab items from the GUI and stuff them into something useful

  $global:Config.Tenant1.DisplayName = $Global:grid_Tenants.Rows[0].Cells[1].Value
  $global:Config.Tenant1.SignInAddress = $Global:grid_Tenants.Rows[0].Cells[2].Value
  $global:Config.Tenant1.ModernAuth = $Global:grid_Tenants.Rows[0].Cells[4].Value
  $global:Config.Tenant1.ConnectToTeams = $Global:grid_Tenants.Rows[0].Cells[5].Value
  $global:Config.Tenant1.ConnectToSkype = $Global:grid_Tenants.Rows[0].Cells[6].Value
  $global:Config.Tenant1.ConnectToExchange = $Global:grid_Tenants.Rows[0].Cells[7].Value
 

  $global:Config.Tenant2.DisplayName = $Global:grid_Tenants.Rows[1].Cells[1].Value
  $global:Config.Tenant2.SignInAddress = $Global:grid_Tenants.Rows[1].Cells[2].Value
  $global:Config.Tenant2.ModernAuth = $Global:grid_Tenants.Rows[1].Cells[4].Value
  $global:Config.Tenant2.ConnectToTeams = $Global:grid_Tenants.Rows[1].Cells[5].Value
  $global:Config.Tenant2.ConnectToSkype = $Global:grid_Tenants.Rows[1].Cells[6].Value
  $global:Config.Tenant2.ConnectToExchange = $Global:grid_Tenants.Rows[1].Cells[7].Value


  $global:Config.Tenant3.DisplayName = $Global:grid_Tenants.Rows[2].Cells[1].Value
  $global:Config.Tenant3.SignInAddress = $Global:grid_Tenants.Rows[2].Cells[2].Value
  $global:Config.Tenant3.ModernAuth = $Global:grid_Tenants.Rows[2].Cells[4].Value
  $global:Config.Tenant3.ConnectToTeams = $Global:grid_Tenants.Rows[2].Cells[5].Value
  $global:Config.Tenant3.ConnectToSkype = $Global:grid_Tenants.Rows[2].Cells[6].Value
  $global:Config.Tenant3.ConnectToExchange = $Global:grid_Tenants.Rows[2].Cells[7].Value

 
  $global:Config.Tenant4.DisplayName = $Global:grid_Tenants.Rows[3].Cells[1].Value
  $global:Config.Tenant4.SignInAddress = $Global:grid_Tenants.Rows[3].Cells[2].Value
  $global:Config.Tenant4.ModernAuth = $Global:grid_Tenants.Rows[3].Cells[4].Value
  $global:Config.Tenant4.ConnectToTeams = $Global:grid_Tenants.Rows[3].Cells[5].Value
  $global:Config.Tenant4.ConnectToSkype = $Global:grid_Tenants.Rows[3].Cells[6].Value
  $global:Config.Tenant4.ConnectToExchange = $Global:grid_Tenants.Rows[3].Cells[7].Value

 
  $global:Config.Tenant5.DisplayName = $Global:grid_Tenants.Rows[4].Cells[1].Value
  $global:Config.Tenant5.SignInAddress = $Global:grid_Tenants.Rows[4].Cells[2].Value
  $global:Config.Tenant5.ModernAuth = $Global:grid_Tenants.Rows[4].Cells[4].Value
  $global:Config.Tenant5.ConnectToTeams = $Global:grid_Tenants.Rows[4].Cells[5].Value
  $global:Config.Tenant5.ConnectToSkype = $Global:grid_Tenants.Rows[4].Cells[6].Value
  $global:Config.Tenant5.ConnectToExchange = $Global:grid_Tenants.Rows[4].Cells[7].Value

 
  $global:Config.Tenant6.DisplayName = $Global:grid_Tenants.Rows[5].Cells[1].Value
  $global:Config.Tenant6.SignInAddress = $Global:grid_Tenants.Rows[5].Cells[2].Value
  $global:Config.Tenant6.ModernAuth = $Global:grid_Tenants.Rows[5].Cells[4].Value
  $global:Config.Tenant6.ConnectToTeams = $Global:grid_Tenants.Rows[5].Cells[5].Value
  $global:Config.Tenant6.ConnectToSkype = $Global:grid_Tenants.Rows[5].Cells[6].Value
  $global:Config.Tenant6.ConnectToExchange = $Global:grid_Tenants.Rows[5].Cells[7].Value

 
  $global:Config.Tenant7.DisplayName = $Global:grid_Tenants.Rows[6].Cells[1].Value
  $global:Config.Tenant7.SignInAddress = $Global:grid_Tenants.Rows[6].Cells[2].Value
  $global:Config.Tenant7.ModernAuth = $Global:grid_Tenants.Rows[6].Cells[4].Value
  $global:Config.Tenant7.ConnectToTeams = $Global:grid_Tenants.Rows[6].Cells[5].Value
  $global:Config.Tenant7.ConnectToSkype = $Global:grid_Tenants.Rows[6].Cells[6].Value
  $global:Config.Tenant7.ConnectToExchange = $Global:grid_Tenants.Rows[6].Cells[7].Value

 
  $global:Config.Tenant8.DisplayName = $Global:grid_Tenants.Rows[7].Cells[1].Value
  $global:Config.Tenant8.SignInAddress = $Global:grid_Tenants.Rows[7].Cells[2].Value
  $global:Config.Tenant8.ModernAuth = $Global:grid_Tenants.Rows[7].Cells[4].Value
  $global:Config.Tenant8.ConnectToTeams = $Global:grid_Tenants.Rows[7].Cells[5].Value
  $global:Config.Tenant8.ConnectToSkype = $Global:grid_Tenants.Rows[7].Cells[6].Value
  $global:Config.Tenant8.ConnectToExchange = $Global:grid_Tenants.Rows[7].Cells[7].Value

 
  $global:Config.Tenant9.DisplayName = $Global:grid_Tenants.Rows[8].Cells[1].Value
  $global:Config.Tenant9.SignInAddress = $Global:grid_Tenants.Rows[8].Cells[2].Value
  $global:Config.Tenant9.ModernAuth = $Global:grid_Tenants.Rows[8].Cells[4].Value
  $global:Config.Tenant9.ConnectToTeams = $Global:grid_Tenants.Rows[8].Cells[5].Value
  $global:Config.Tenant9.ConnectToSkype = $Global:grid_Tenants.Rows[8].Cells[6].Value
  $global:Config.Tenant9.ConnectToExchange = $Global:grid_Tenants.Rows[8].Cells[7].Value


  $global:Config.Tenant10.DisplayName = $Global:grid_Tenants.Rows[9].Cells[1].Value
  $global:Config.Tenant10.SignInAddress = $Global:grid_Tenants.Rows[9].Cells[2].Value
  $global:Config.Tenant10.ModernAuth = $Global:grid_Tenants.Rows[9].Cells[4].Value
  $global:Config.Tenant10.ConnectToTeams = $Global:grid_Tenants.Rows[9].Cells[5].Value
  $global:Config.Tenant10.ConnectToSkype = $Global:grid_Tenants.Rows[9].Cells[6].Value
  $global:Config.Tenant10.ConnectToExchange = $Global:grid_Tenants.Rows[9].Cells[7].Value

  #Encrypt passwords
  If ($Global:grid_Tenants.Rows[0].Cells[3].Value -ne "****") {$global:Config.Tenant1.Credential = ($Global:grid_Tenants.Rows[0].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[1].Cells[3].Value -ne "****") {$global:Config.Tenant2.Credential = ($Global:grid_Tenants.Rows[1].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[2].Cells[3].Value -ne "****") {$global:Config.Tenant3.Credential = ($Global:grid_Tenants.Rows[2].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[3].Cells[3].Value -ne "****") {$global:Config.Tenant4.Credential = ($Global:grid_Tenants.Rows[3].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[4].Cells[3].Value -ne "****") {$global:Config.Tenant5.Credential = ($Global:grid_Tenants.Rows[4].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[5].Cells[3].Value -ne "****") {$global:Config.Tenant6.Credential = ($Global:grid_Tenants.Rows[5].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[6].Cells[3].Value -ne "****") {$global:Config.Tenant7.Credential = ($Global:grid_Tenants.Rows[6].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[7].Cells[3].Value -ne "****") {$global:Config.Tenant8.Credential = ($Global:grid_Tenants.Rows[7].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[8].Cells[3].Value -ne "****") {$global:Config.Tenant9.Credential = ($Global:grid_Tenants.Rows[8].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[9].Cells[3].Value -ne "****") {$global:Config.Tenant10.Credential = ($Global:grid_Tenants.Rows[10].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}

  #Clear the password fields
  $Global:grid_Tenants.Rows[0].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[1].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[2].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[3].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[4].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[5].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[6].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[7].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[8].Cells[3].Value = "****"
  $Global:grid_Tenants.Rows[9].Cells[3].Value = "****"


  #Write the XML File
  Try
  {
    $global:Config| Export-CliXml -Path "$ENV:UserProfile\BounShell.xml"
    Write-Log -component $function -Message "Config File Saved" -severity 2
  }
  Catch 
  {
    Write-Log -component $function -Message "Error writing Config file" -severity 3
  }
}

Function Import-BsDefaultConfig 
{
  #Set Variables to Defaults
  #Remove and re-create the Config Array
  [Void](Remove-Variable -Name Config -Scope Script -ErrorAction SilentlyContinue)
  $global:Config=@{}
  #Populate with Defaults
  [void] $Global:grid_Tenants.Rows.Clear()
  
  $global:Config.Tenant1 =@{}
  $global:Config.Tenant1.DisplayName = "Undefined"
  $global:Config.Tenant1.SignInAddress = "user1@fabrikam.com"
  $global:Config.Tenant1.Credential = "****"
  $global:Config.Tenant1.ModernAuth = $false
  $global:Config.Tenant1.ConnectToTeams = $false
  $global:Config.Tenant1.ConnectToSkype = $false
  $global:Config.Tenant1.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("1",'Undefined','user1@fabrikam.com',"****",$False,$false,$false,$false)
  
  
  $global:Config.Tenant2 =@{}
  $global:Config.Tenant2.DisplayName = "Undefined"
  $global:Config.Tenant2.SignInAddress = "user2@fabrikam.com"
  $global:Config.Tenant2.Credential = "****"
  $global:Config.Tenant2.ModernAuth = $false
  $global:Config.Tenant2.ConnectToTeams = $false
  $global:Config.Tenant2.ConnectToSkype = $false
  $global:Config.Tenant2.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("2",'Undefined','user2@fabrikam.com',"****",$False,$false,$false,$false)

  $global:Config.Tenant3 =@{}
  $global:Config.Tenant3.DisplayName = "Undefined"
  $global:Config.Tenant3.SignInAddress = "user3@fabrikam.com"
  $global:Config.Tenant3.Credential = "****"
  $global:Config.Tenant3.ModernAuth = $false
  $global:Config.Tenant3.ConnectToTeams = $false
  $global:Config.Tenant3.ConnectToSkype = $false
  $global:Config.Tenant3.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("3",'Undefined','user3@fabrikam.com',"****",$False,$false,$false,$false)

  $global:Config.Tenant4 =@{}
  $global:Config.Tenant4.DisplayName = "Undefined"
  $global:Config.Tenant4.SignInAddress = "user4@fabrikam.com"
  $global:Config.Tenant4.Credential = "****"
  $global:Config.Tenant4.ModernAuth = $false
  $global:Config.Tenant4.ConnectToTeams = $false
  $global:Config.Tenant4.ConnectToSkype = $false
  $global:Config.Tenant4.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("4",'Undefined','user4@fabrikam.com',"****",$False,$false,$false,$false)

  $global:Config.Tenant5 =@{}
  $global:Config.Tenant5.DisplayName = "Undefined"
  $global:Config.Tenant5.SignInAddress = "user5@fabrikam.com"
  $global:Config.Tenant5.Credential = "****"
  $global:Config.Tenant5.ModernAuth = $false
  $global:Config.Tenant5.ConnectToTeams = $false
  $global:Config.Tenant5.ConnectToSkype = $false
  $global:Config.Tenant5.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("5",'Undefined','user5@fabrikam.com',"****",$False,$false,$false,$false)

  $global:Config.Tenant6 =@{}
  $global:Config.Tenant6.DisplayName = "Undefined"
  $global:Config.Tenant6.SignInAddress = "user6@fabrikam.com"
  $global:Config.Tenant6.Credential = "****"
  $global:Config.Tenant6.ModernAuth = $false
  $global:Config.Tenant6.ConnectToTeams = $false
  $global:Config.Tenant6.ConnectToSkype = $false
  $global:Config.Tenant6.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("6",'Undefined','user6@fabrikam.com',"****",$False,$false,$false,$false)

  $global:Config.Tenant7 =@{}
  $global:Config.Tenant7.DisplayName = "Undefined"
  $global:Config.Tenant7.SignInAddress = "user@fabrikam.com"
  $global:Config.Tenant7.Credential = "****"
  $global:Config.Tenant7.ModernAuth = $false
  $global:Config.Tenant7.ConnectToTeams = $false
  $global:Config.Tenant7.ConnectToSkype = $false
  $global:Config.Tenant7.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("7",'Undefined','user7@fabrikam.com',"****",$False,$false,$false,$false)

  $global:Config.Tenant8 =@{}
  $global:Config.Tenant8.DisplayName = "Undefined"
  $global:Config.Tenant8.SignInAddress = "user8@fabrikam.com"
  $global:Config.Tenant8.Credential = "****"
  $global:Config.Tenant8.ModernAuth = $false
  $global:Config.Tenant8.ConnectToTeams = $false
  $global:Config.Tenant8.ConnectToSkype = $false
  $global:Config.Tenant8.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("8",'Undefined','user8@fabrikam.com',"****",$False,$false,$false,$false)
    
  $global:Config.Tenant9 =@{}
  $global:Config.Tenant9.DisplayName = "Undefined"
  $global:Config.Tenant9.SignInAddress = "user@fabrikam.com"
  $global:Config.Tenant9.Credential = "****"
  $global:Config.Tenant9.ModernAuth = $false
  $global:Config.Tenant9.ConnectToTeams = $false
  $global:Config.Tenant9.ConnectToSkype = $false
  $global:Config.Tenant9.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("9",'Undefined','user9@fabrikam.com',"****",$False,$false,$false,$false)
  
  $global:Config.Tenant10 =@{}
  $global:Config.Tenant10.DisplayName = "Undefined"
  $global:Config.Tenant10.SignInAddress = "user@fabrikam.com"
  $global:Config.Tenant10.Credential = "****"
  $global:Config.Tenant10.ModernAuth = $false
  $global:Config.Tenant10.ConnectToTeams = $false
  $global:Config.Tenant10.ConnectToSkype = $false
  $global:Config.Tenant10.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("10",'Undefined','user10@fabrikam.com',"****",$False,$false,$false,$false)
  
  [Float]$global:Config.ConfigFileVersion = "0.1"
  [string]$global:Config.Description = "BounShell Configuration file. See Skype4BAdmin.com for more information"
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
    [Parameter(Mandatory=$true)] [string]$Tabname,
    [Parameter(Mandatory=$true)] [float]$Tenant
  )
  
  $Function= 'Invoke-BsNewTenantTab'
  Write-Log -component $function -Message "Called Invoke-BsNewTenantTab to connect to Tenant $tenant with a Tabname of $tabname" -severity 1 
  if($tabname -ne 'Undefined') 
  {
    Try
    {
      #kick off a new tab and call it tabname
      Write-Log -component $function -Message "Opening new ISE tab..." -severity 1 
      $TabNameTab=$psISE.PowerShellTabs.Add()
      $TabNameTab.DisplayName = $Tabname
      
      #Wait for the tab to wake up
      Write-Log -component $function -Message "Waiting for tab to become invokable" -severity 1 
      Do 
      {sleep -m 100}
      While (!$TabNameTab.CanInvoke)
      
      #Kick off the connection
      Write-Log -component $function -Message "Invoking Command: Connect-BsO365Tenant -Tenant $Tenant" -severity 1
      $TabNameTab.Invoke("Connect-BsO365Tenant -Tenant $Tenant")
      
    }
    
    Catch
    {
      #Something went wrong opening a new tab, probably already a tab with that name open
      Write-Log -component $function -Message "Failed to open new tab, Is there already a connection open to that tenant?" -severity 3
    }
  }
  
  Else
  {
    #Tabname is "undefined", user clicked a tenant thats not confgured yey
    Write-Log -component $function -Message "Sorry, I cant find a config for Tenant $tenant" -severity 3
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
  [string]$Function = 'Connect-BsO365Tenant'
  [bool]$ModernAuth = $false
  [bool]$ConnectToTeams = $false
  [bool]$ConnectToSkype = $false
  [bool]$ConnectToExchange = $false
  [bool]$ConnectToSharepoint = $false
  [bool]$ConnectToAAD = $false
  [bool]$ConnectToCompliance = $false
  $ModernAuthPassword = ConvertTo-SecureString "Foo" -asplaintext -force
  [string]$ModernAuthUsername
  
  Write-Log -component $Function -Message "Called to connect to Tenant $tenant" -severity 1
  
  #Check to see if we are running in the ISE
  If ($PSISE)
  {
    #load the relevant stuff for the ISE enviroment{
    Import-BsGuiElements
    #Clean up any stale sessions (we shouldnt have any, but whatever)
    Get-PSSession | Remove-PSSession
  }

  #Import the Config file so we have data  
  Read-BsConfigFile
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant1.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant1.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant1.SignInAddress,$global:Config.Tenant1.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant2.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant2.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant2.SignInAddress,$global:Config.Tenant2.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant3.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant3.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant3.SignInAddress,$global:Config.Tenant3.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant4.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant4.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant4.SignInAddress,$global:Config.Tenant4.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant5.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant5.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant5.SignInAddress,$global:Config.Tenant5.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant6.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant6.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant6.SignInAddress,$global:Config.Tenant6.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant7.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant7.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant7.SignInAddress,$global:Config.Tenant7.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant8.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant8.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant8.SignInAddress,$global:Config.Tenant8.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant9.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant9.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant9.SignInAddress,$global:Config.Tenant9.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
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
      Write-Log -component $function -Message "Loading $($global:Config.Tenant10.DisplayName) Settings" -severity 2
      #Check to see if the tenant is configured for modern auth
      If (!$global:Config.Tenant10.ModernAuth) 
      {
        #Not using modern auth
        $global:pscred = New-Object System.Management.Automation.PSCredential($global:Config.Tenant10.SignInAddress,$global:Config.Tenant10.Credential)
      }
      Else
      {
        #Using modern auth
        $ModernAuth = $True
        #Convert the config into something we can work with later
        $ModernAuthPassword = $global:Config.Tenant10.Credential
        $ModernAuthUsername = $global:Config.Tenant10.SignInAddress
      }
    }
    
    
  }

  #endregion tenantswitch

  #check to see if the Modern Auth flag has been set and use the appropriate connection method
  If ($ModernAuth) 
  {
    Write-host "Modern Auth is a Beta feature...." #Todo.
    Write-host "Your Username will be copied to the clipboard, Paste it into the Modern Auth Window"
    Write-host "Once Ctrl+V has been pressed BounShell will copy your password into the clipboard"
    Write-host "Upon pasting this, BounShell will clear the clipboard and overwrite the memory just incase."

    #As we are dealing with modern auth we need to convert the password back to an insecure string do that here
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ModernAuthPassword)
    $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
     

    If ($ConnectToTeams) {
      Try
      {
        # So Now we need to kick off a new window that waits for the clipboard events
        Write-Log -Message "Connecting to Microsoft Teams" -Severity 2 -Component $Function
        #Create a script block with the expanded variables
        [String]$cmd = "Watch-BsCredentials -ModernAuthUsername $ModernAuthUsername -UnsecurePassword $UnsecurePassword"
        [ScriptBlock]$sb = [ScriptBlock]::Create($cmd) 
        
        #and now call it
        Start-process powershell $sb
        
        #Now we can invoke the session
        #$TeamsSession = (Connect-MicrosoftTeams)
        Connect-MicrosoftTeams
      } 
      Catch {
        $ErrorMessage = $_.Exception.Message
        Write-log -Message $ErrorMessage -Severity 3 -Component $Function 
        Write-log -Message 'Error connecting to Microsoft Teams' -Severity 3 -Component $Function
      }
    }
  }
  



  #region NoModern
  #If the modern auth flag hasnt been set, we can simply connect to the services using secure credentials
  If (!$ModernAuth) 
  {
    #See if we got passed creds
    Write-Log -Message 'Checking for Office365 Credentials' -Severity 1 -Component $Function
    If ($pscred -eq $null) 
    {
      #No credentials, prompt user for some
      Write-Log -Message 'No Office365 credentials Found, Prompting user for creds' -Severity 3 -Component $Function
      $psCred = Get-Credential
    }
    Else
    {
      #Found creds, continue
      Write-Log -Message "Found Office365 Creds for Username: $($pscred.username)" -Severity 1 -Component $Function
    }
    
    #Check for the Exchange connection flag
    If ($ConnectToExchange)
    {
      #Flag is set, connect to Exchange
      Try
      {
        #Exchange connection try block
        Write-Log -Message "Connecting to Exchange Online" -Severity 2 -Component $Function
        $O365Session = (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $pscred -Authentication Basic -AllowRedirection )
        Write-Log -Message "Importing Session" -Severity 1 -Component $Function
        $VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $O365Session -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = "Continue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch 
      {
        #We had an issue connecting to Exchange
        $ErrorMessage = $_.Exception.Message
        Write-log -Message $ErrorMessage -Severity 3 -Component $Function
        Write-log -Message 'Error connecting to Exchange Online' -Severity 3 -Component $Function
      }
    }

    #Check for the Skype4B connection flag
    If ($ConnectToSkype) 
    {
      #Flag is set, connect to Skype4B
      Try
      {
        #Skype connection try block
        Write-Log -Message "Connecting to Skype4B Online" -Severity 2 -Component $Function
        $S4BOSession = (New-CsOnlineSession -Credential $pscred)
        $VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $S4BOSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = "Continue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch
      {
        #We had an issues connecting to Skype
        $ErrorMessage = $_.Exception.Messag
        Write-log -Message $ErrorMessage -Severity 3 -Component $Function 
        Write-log -Message 'Error connecting to Skype4B Online' -Severity 3 -Component $Function
      }
    }
    
    #Check for the Teams connection flag
    If ($ConnectToTeams)
    {
      #Flag is set, connect to Teams
      Try 
      {
        #Teams connection try block
        Write-Log -Message "Connecting to Microsoft Teams" -Severity 2 -Component $Function
        $TeamsSession = (Connect-MicrosoftTeams -Credential $pscred)
        $VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        #No need to import the session. Import-Module (Import-PSSession -Session $TeamsSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = "Continue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      } 
      Catch
      {
        #We had an issue connecting to Teams
        $ErrorMessage = $_.Exception.Message
        Write-log -Message $ErrorMessage -Severity 3 -Component $Function 
        Write-log -Message 'Error connecting to Microsoft Teams' -Severity 3 -Component $Function
      }
    }
    
    #Check for the Sharepoint connection flag
    If ($ConnectToSharepoint) 
    {
      #Flag is set, connect to Sharepoint
      Try
      {
        #Sharepoint connection try block
        Write-Log -Message "Connecting to Sharepoint Online" -Severity 2 -Component $Function
        $SharepointSession = (Connect-SPOService -Credential $pscred)
        $VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $SharepointSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = "Continue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to Sharepoint
        $ErrorMessage = $_.Exception.Message
        Write-log -Message $ErrorMessage -Severity 3 -Component $Function 
        Write-log -Message 'Error connecting to Sharepoint Online' -Severity 3 -Component $Function
      }
    }

    #Check for the AzureAD connection flag
    If ($ConnectToAAD) 
    {
      #Flag is set, connect to AzureAD
      Try
      {
        #Azure AD try block
        Write-Log -Message "Connecting to Azure AD" -Severity 2 -Component $Function
        $AADSession = (Connect-AzureAD -Credential $pscred)
        $VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $AADSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = "Continue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to AzureAD
        $ErrorMessage = $_.Exception.Message
        Write-log -Message $ErrorMessage -Severity 3 -Component $Function 
        Write-log -Message 'Error connecting to Azure AD' -Severity 3 -Component $Function
      }
    }
    
    #Check for the 365 Compliance Centre flag
    If ($ConnectToCompliance)
    {
      #Flag is set, connect to Compiance Centre
      Try
      {
        Write-Log -Message "Connecting to Office 365 Compliance Centre" -Severity 2 -Component $Function
        $ComplianceSession = (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection)
        $VerbosePreference = "SilentlyContinue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
        Import-Module (Import-PSSession -Session $ComplianceSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        $VerbosePreference = "Continue" #Todo. fix for  import-psmodule ignoring the -Verbose:$false flag
      }
      Catch
      {
        #We had an issue connecting to the Compliance Centre
        $ErrorMessage = $_.Exception.Message
        Write-log -Message $ErrorMessage -Severity 3 -Component $Function 
        Write-log -Message 'Error connecting to Office 365 Compliance Centre' -Severity 3 -Component $Function
      }
    }
  }
  #endregion NoModern 

}

Function Update-BsAddonMenu 
{

 
  #Check to see if we are loaded, if we are cleanup after ourselves
  if (($psISE.CurrentPowerShellTab.AddOnsMenu.Submenus).displayname -eq "_BounShell") 
  {
    [void]$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.remove($Global:isemenuitem)
  }

  #Create Initial Menu Object
  [void]($Global:IseMenuItem = ($psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('_BounShell',$null ,$null)))

  #Add the Settings Button

  [void]($Global:IseMenuItem.Submenus.add('_Settings...', {Show-BsGuiElements}, $null) )


  #Now add each Tenant

  #Need to put a for each code in here that adds Tenant 1 through 10
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant1.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant1.DisplayName -Tenant 1}, 'Ctrl+Alt+1'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant2.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant2.DisplayName -Tenant 2}, 'Ctrl+Alt+2'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant3.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant3.DisplayName -Tenant 3}, 'Ctrl+Alt+3'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant4.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant4.DisplayName -Tenant 4}, 'Ctrl+Alt+4'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant5.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant5.DisplayName -Tenant 5}, 'Ctrl+Alt+5'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant6.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant6.DisplayName -Tenant 6}, 'Ctrl+Alt+6'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant7.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant7.DisplayName -Tenant 7}, 'Ctrl+Alt+7'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant8.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant8.DisplayName -Tenant 8}, 'Ctrl+Alt+8'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant9.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant9.DisplayName -Tenant 9}, 'Ctrl+Alt+9'))
  [void]($Global:IseMenuItem.Submenus.add("$($global:Config.Tenant10.DisplayName)",{Invoke-BsNewTenantTab -tabname $global:Config.Tenant10.DisplayName -Tenant 10}, 'Ctrl+Alt+0'))
}

Function Test-ManagementTools
{
  Write-Log -component "Test-ManagementTools" -Message "Checking for Lync/Skype management tools"
  $CSManagementTools = $false
  If(!(Get-Module "SkypeForBusiness")) 
  {
    Import-Module SkypeForBusiness -Verbose:$false
  }
  If(!(Get-Module "Lync"))
  {
    Import-Module Lync -Verbose:$false
  }
  If(Get-Module "SkypeForBusiness")
  {
    $CSManagementTools = $true
  }
  If(Get-Module "Lync")
  {
    $CSManagementTools = $true
  }
  If(!$CSManagementTools)
  {
    Write-Log -component "Test-ManagementTools" -Message "Could not locate Lync/Skype4B Management tools" -severity 3 
    Throw  "Could not locate Lync/Skype4B Management tools"
  }

  #Check for the AD Management Tools
  $ADManagementTools = $false
  if(!(Get-Module "ActiveDirectory"))
  {
    Import-Module ActiveDirectory -Verbose:$false
  }
  if(Get-Module "ActiveDirectory") 
  {
    $ADManagementTools = $true
  }
  if(!$ADManagementTools)
  {
    Write-Log -component "Test-ManagementTools" -Message "Could not locate Active Directory Management tools" -severity 3
    Throw  "Could not locate Active Directory Management tools"
  }

  $TeamsManagementTools = $false
  If(!(Get-Module "MicrosoftTeams"))
  {
    Import-Module MicrosoftTeams -Verbose:$false
  }
  If(Get-Module "MicrosoftTeams")
  {
    $TeamsManagementTools = $true
  }
  If(!$TeamsManagementTools)
  {
    Write-Log -component "Test-CsManagementTools" -Message "Could not locate Teams PowerShell Module" -severity 3 
    Write-Log -component "Test-CsManagementTools" -Message "Run the cmdlet 'Install-Module MicrosoftTeams' to install it" -severity 3 
  }
}

Function Import-BsGuiElements 
{
  #First we need to import the Functions so they exist for the GUI items
  Import-BsGuiFunctions

  #region Gui
  [void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
  [void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
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
  ([System.ComponentModel.ISupportInitialize]$Global:grid_Tenants).BeginInit()
  $Global:SettingsForm.SuspendLayout()
  #
  #btn_CancelConfig
  #
  $Global:btn_CancelConfig.BackColor = [System.Drawing.Color]::White
  $Global:btn_CancelConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:btn_CancelConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Pixel,([System.Byte][System.Byte]0)))
  $Global:btn_CancelConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:btn_CancelConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]821,[System.Int32]368))
  $Global:btn_CancelConfig.Name = [System.String]'btn_CancelConfig'
  $Global:btn_CancelConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]94,[System.Int32]23))
  $Global:btn_CancelConfig.TabIndex = [System.Int32]59
  $Global:btn_CancelConfig.Text = [System.String]'Cancel'
  $Global:btn_CancelConfig.UseVisualStyleBackColor = $false
  $Global:btn_CancelConfig.add_Click($Global:btn_CancelConfig_Click)
  #
  #Btn_ReloadConfig
  #
  $Global:Btn_ReloadConfig.BackColor = [System.Drawing.Color]::White
  $Global:Btn_ReloadConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:Btn_ReloadConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Pixel,([System.Byte][System.Byte]0)))
  $Global:Btn_ReloadConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:Btn_ReloadConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]589,[System.Int32]368))
  $Global:Btn_ReloadConfig.Name = [System.String]'Btn_ReloadConfig'
  $Global:Btn_ReloadConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110,[System.Int32]23))
  $Global:Btn_ReloadConfig.TabIndex = [System.Int32]58
  $Global:Btn_ReloadConfig.Text = [System.String]'Reload Config'
  $Global:Btn_ReloadConfig.UseVisualStyleBackColor = $true
  $Global:Btn_ReloadConfig.add_Click($Global:Btn_ConfigReload_Click)
  #
  #Btn_SaveConfig
  #
  $Global:Btn_SaveConfig.BackColor = [System.Drawing.Color]::White
  $Global:Btn_SaveConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:Btn_SaveConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Pixel,([System.Byte][System.Byte]0)))
  $Global:Btn_SaveConfig.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:Btn_SaveConfig.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]705,[System.Int32]368))
  $Global:Btn_SaveConfig.Name = [System.String]'Btn_SaveConfig'
  $Global:Btn_SaveConfig.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110,[System.Int32]23))
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

  $Global:cbx_AutoUpdates.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]370))
  $Global:cbx_AutoUpdates.Name = [System.String]'cbx_AutoUpdates'
  $Global:cbx_AutoUpdates.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]183,[System.Int32]17))
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
  $Global:dataGridViewCellStyle1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel,([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle1.ForeColor = [System.Drawing.SystemColors]::WindowText
  $Global:dataGridViewCellStyle1.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle1.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle1.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
  $Global:grid_Tenants.ColumnHeadersDefaultCellStyle = $Global:dataGridViewCellStyle1
  $Global:grid_Tenants.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
  $Global:grid_Tenants.Columns.AddRange($Global:Tenant_ID,$Global:Tenant_DisplayName,$Global:Tenant_Email,$Global:Tenant_Credentials,$Global:Tenant_ModernAuth,$Global:Tenant_Teams,$Global:Tenant_Skype,$Global:Tenant_Exchange)
  $Global:dataGridViewCellStyle2.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $Global:dataGridViewCellStyle2.BackColor = [System.Drawing.SystemColors]::Window
  $Global:dataGridViewCellStyle2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel,([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle2.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:dataGridViewCellStyle2.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle2.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle2.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
  $Global:grid_Tenants.DefaultCellStyle = $Global:dataGridViewCellStyle2
  $Global:grid_Tenants.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]12))
  $Global:grid_Tenants.Name = [System.String]'grid_Tenants'
  $Global:dataGridViewCellStyle3.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $Global:dataGridViewCellStyle3.BackColor = [System.Drawing.SystemColors]::Control
  $Global:dataGridViewCellStyle3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel,([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle3.ForeColor = [System.Drawing.SystemColors]::WindowText
  $Global:dataGridViewCellStyle3.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle3.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle3.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
  $Global:grid_Tenants.RowHeadersDefaultCellStyle = $Global:dataGridViewCellStyle3
  $Global:grid_Tenants.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]903,[System.Int32]336))
  $Global:grid_Tenants.TabIndex = [System.Int32]76
  $Global:grid_Tenants.add_CellContentClick($Global:grid_Tenants_CellContentClick)
  #
  #Btn_Default
  #
  $Global:Btn_Default.BackColor = [System.Drawing.Color]::White
  $Global:Btn_Default.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Global:Btn_Default.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Pixel,([System.Byte][System.Byte]0)))
  $Global:Btn_Default.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:Btn_Default.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]473,[System.Int32]368))
  $Global:Btn_Default.Name = [System.String]'Btn_Default'
  $Global:Btn_Default.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]110,[System.Int32]23))
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
  #Global:SettingsForm
  #
  $Global:SettingsForm.BackColor = [System.Drawing.Color]::White
  $Global:SettingsForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]925,[System.Int32]404))
  $Global:SettingsForm.Controls.Add($Global:Btn_Default)
  $Global:SettingsForm.Controls.Add($Global:grid_Tenants)
  $Global:SettingsForm.Controls.Add($Global:cbx_AutoUpdates)
  $Global:SettingsForm.Controls.Add($Global:btn_CancelConfig)
  $Global:SettingsForm.Controls.Add($Global:Btn_ReloadConfig)
  $Global:SettingsForm.Controls.Add($Global:Btn_SaveConfig)
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
  Add-Member -InputObject $Global:SettingsForm -Name Tenant_Exchange -Value $Global:Tenant_Exchange -MemberType NoteProperty
  Add-Member -InputObject $Global:SettingsForm -Name cbx_AutoUpdates -Value $Global:cbx_AutoUpdates -MemberType NoteProperty
  #endregion Gui
}

Function Import-BsGuiFunctions 
{
  #Gui Cancel button
  $Global:btn_CancelConfig_Click =
  {
    Read-BsConfigFile
    [void]$Global:SettingsForm.Hide()
  }

  #Gui Save Config Button
  $Global:Btn_SaveConfig_Click =
  {
    $Global:btn_CancelConfig.Text = [System.String]'Close'
    Write-BsConfigFile
    Update-BsAddonMenu
  }

  #Gui Set Defaults Button
  $Global:Btn_Default_Click =
  {
    Import-BsDefaultConfig
    Update-BsAddonMenu
  }

  #Gui Button to Reload Config
  $Global:Btn_ConfigReload_Click =
  {
    Read-BsConfigFile
    Update-BsAddonMenu
  }


}

Function Show-BsGuiElements
{
  #Reset the cancel button
  $Global:btn_CancelConfig.Text = [System.String]'Cancel'
  [void]$Global:SettingsForm.ShowDialog()
}

Function Hide-BsGuiElements
{

  [void]$Global:SettingsForm.Hide()
}

Function Start-BounShell
{
  $function = 'Start-BounShell'
  #Allows us to seperate all the "onetime" run objects incase we get dot sourced.
  Write-Log -component $function -Message "Script executed from $PSScriptRoot" -severity 1
  Write-Log -component $function -Message "Loading BounShell..." -severity 2

  #Check we are actually in the ISE
  If(!$PSISE) 
  {   
    Write-Log -component $function -Message 'Could not locate $PSISE Variable' -severity 3
    Write-Log -component $function -Message 'Sorry, BounShell is designed to be run from the PowerShell ISE' -severity 2
    Write-Log -component $function -Message 'You can however connect to a pre-configured tenant manually using the cmdlet' -severity 2
    Write-Log -component $function -Message 'PS> Connect-BsO365Tenant -Tenant 1'  -severity 2
    Write-Log -component $function -Message 'This will connect to the Tenant stored in slot 1 in the current context'  -severity 2
    Return #Yes I know Return sucks, but I think its better than Throw.
    
  }

  #Load the Gui Elements
  Import-BsGuiElements
  #check for script update
  if ($SkipUpdateCheck -eq $false)
  {
    Get-ScriptUpdate
  } #todo enable update checking

  #check for config file then load the default

  #Check for and load the config file if present
  If(Test-Path $global:ConfigFilePath)
  {
    Write-Log -component $function -Message "Found $ConfigFilePath, loading..." -severity 1
    Read-BsConfigFile
  }

  Else
  {
    Write-Log -component $function -Message "Could not locate $ConfigFilePath, Using Defaults" -severity 3
    #If there is no config file. Load a default
    Import-BsDefaultConfig

    Write-Log -component $function -Message "As we didnt find a config file we will assume this is a first run." -severity 3
    Write-Log -component $function -Message "Thus we will remind you that while all care is taken to store your credentials in a safe manner, I cannot be held responsible for any data breaches" -severity 3
    Write-Log -component $function -Message "If someone was to get a hold of your BounShell.xml AND your user profile private encryption key its possible to reverse engineer stored credentials" -severity 3
    Write-Log -component $function -Message "Seriously, Whilst the password store is encrypted, its not perfect!" -severity 3
  }

  #Check for Management Tools
  #Test-ManagementTools #todo fix
  
  #Now Create the Objects in the ISE
  Update-BsAddonMenu
  Write-Log -component $function -Message "BounShell Loaded" -severity 2
}

Function Watch-BsCredentials
{
  

  [CmdletBinding()]
  PARAM
  (
    $ModernAuthUsername,
    $UnsecurePassword
  )
  [string]$Function = 'Watch-BsCredentials'
  If (!$ModernAuthUsername)
  {
    Write-Log -component $Function -Message "This Cmdlet is for BoundShell's internal use only. Please use 'Start-Bounshell' to launch the tool" -severity 3
    pause
    return
  }
  Write-Log -component $Function -Message "Called to connect to $ModernAuthUsername" -severity 3
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
  Write-Log -component $Function -Message "$ModernAuthUsername placed into Clipboard" -severity 1
  Write-Log -component $Function -Message "Press 'Ctrl+v' to paste the username $ModernAuthUsername in the modern auth window" -severity 3


  do 
  {
    Start-Sleep -Milliseconds 40
  }
  until ($API::GetAsyncKeyState($ascii) -eq -32767)

  Set-Clipboard -Value $UnsecurePassword
  Write-Log -component $Function -Message "Password placed into Clipboard" -severity 1
  Write-Log -component $Function -Message "Press 'Ctrl+v' to paste the password in the modern auth window" -severity 3

  do 
  {
    Start-Sleep -Milliseconds 40
  }
  until ($API::GetAsyncKeyState($ascii) -eq -32767)

  Set-Clipboard -Value "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaaaaaaaaaa"

}

#now we export the relevant stuff

Export-ModuleMember Read-BsConfigFile
Export-ModuleMember Write-BsConfigFile
Export-ModuleMember Import-BsDefaultConfig
Export-ModuleMember Invoke-BsNewTenantTab
Export-ModuleMember Connect-BsO365Tenant
Export-ModuleMember Update-BsAddonMenu
Export-ModuleMember Import-BsGuiElements
Export-ModuleMember Import-BsGuiFunctions
Export-ModuleMember Show-BsGuiElements
Export-ModuleMember Hide-BsGuiElements
Export-ModuleMember Start-BounShell
Export-ModuleMember Watch-BsCredentials
