<#
    .SYNOPSIS

    This is a tool to help users manage multiple office 365 tennants

    .DESCRIPTION

    Created by James Arber. www.skype4badmin.com
    
    .NOTES

    Version      	        : 0.1
    Date			        : 5/11/2018
    Lync Version		    : Tested against Skype4B 2015
    Author    			    : James Arber
    Header stolen from      : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

    

    :v1.0: Initial Release

    .LINK
    https://www.skype4badmin.com

    .KNOWN ISSUES
    Beta, Buggy as all get out.

    .EXAMPLE
    Loads the Module
    PS C:\> Start-BounShell.ps1

#>

[CmdletBinding(DefaultParametersetName="Common")]
param(
  [Parameter(Mandatory=$false)] [switch]$DisableScriptUpdate,
  [Parameter(Mandatory=$false, Position = 1)] [String]$ConfigFilePath = $null,
  [Parameter(Mandatory=$false, Position = 2)] [String]$LogFileLocation = $null,
  [Parameter(Mandatory=$false, Position = 3)] [float]$Tenant = $null

)

#region config
[Net.ServicePointManager]::SecurityProtocol = 'tls12, tls11, tls'
$StartTime                  =  Get-Date
$VerbosePreference          =  "Continue" #TODO
[float]$ScriptVersion       =  '0.1'
[string]$GithubRepo         =  'Start-BounShell' ##todo
[string]$GithubBranch       =  'devel' #todo
[string]$BlogPost           =  'http://www.skype4badmin.com/BounShell/' #todo

#Check to see if paths were specified, Otherwise set defaults
If (!$LogFileLocation) 
{
  $Script:LogFileLocation = $PSCommandPath -replace '.ps1', '.log'
}

If (!$ConfigFilePath) 
{
  $Script:ConfigFilePath = "$ENV:UserProfile\BounShell.xml"
}

#endregion config


Function Write-Log {
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
      http://www.skype4badmin.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>
  [CmdletBinding()]
  PARAM(
    [String]$Message,
    [String]$Path = $Script:LogFileLocation,
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

Function Test-CsManagementTools {
  Write-Log -component "Test-CsManagementTools" -Message "Checking for Lync/Skype management tools"
  $ManagementTools = $false
  if(!(Get-Module "SkypeForBusiness")) {Import-Module SkypeForBusiness -Verbose:$false}
  if(!(Get-Module "Lync")) {Import-Module Lync -Verbose:$false}
  if(Get-Module "SkypeForBusiness") {$ManagementTools = $true}
  if(Get-Module "Lync") {$ManagementTools = $true}
  if(!$ManagementTools) {
    Write-Log 
    Write-Log -component "Test-CsManagementTools" -Message "Could not locate Lync/Skype4B Management tools" -severity 3 
    throw  "Could not locate Lync/Skype4B Management tools"
  }
	
  #Check for the AD Management Tools
  $ADManagementTools = $false
  if(!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory -Verbose:$false}
  if(Get-Module "ActiveDirectory") {$ADManagementTools = $true}
  if(!$ADManagementTools) {
    Write-Log 
    Write-Log -component "Test-CsManagementTools" -Message "Could not locate Active Directory Management tools" -severity 3
    throw  "Could not locate Active Directory Management tools"
  }

}

Function Get-IEProxy {
  Write-Host "Info: Checking for proxy settings" -ForegroundColor Green
  If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0) {
    $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
    if ($proxies) {
      if ($proxies -ilike "*=*") {
        return $proxies -replace "=", "://" -split (';') | Select-Object -First 1
      }
      Else {
        return ('http://{0}' -f $proxies)
      }
    }
    Else {
      return $null
    }
  }
  Else {
    return $null
  }
}

Function Get-ScriptUpdate {
  Write-Log -component "Self Update" -Message "Checking for Script Update" -severity 1
  Write-Log -component "Self Update" -Message "Checking for Proxy" -severity 1
  $ProxyURL = Get-IEProxy
  If ( $ProxyURL) {
    Write-Log -component "Self Update" -Message "Using proxy address $ProxyURL" -severity 1
  }
  Else {
    Write-Log -component "Self Update" -Message "No proxy setting detected, using direct connection" -severity 1
  }
	
  $GitHubScriptVersion = Invoke-WebRequest https://raw.githubusercontent.com/atreidae/BounShell/devel/version -TimeoutSec 10 -Proxy $ProxyURL #todo change back to master!
  If ($GitHubScriptVersion.Content.length -eq 0) {
    Write-Log -component "Self Update" -Message "Error checking for new version. You can check manualy here" -severity 3
    Write-Log -component "Self Update" -Message "http://www.skype4badmin.com/find-and-test-user-ip-addresses-in-the-skype-location-database" -severity 1 #Todo Update URL
    Write-Log -component "Self Update" -Message "Pausing for 5 seconds" -severity 1
    start-sleep 5
  }
  else { 
    if ([single]$GitHubScriptVersion.Content -gt [single]$ScriptVersion) {
      Write-Log -component "Self Update" -Message "New Version Available" -severity 3
      #New Version available

      #Prompt user to download
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
          Write-Log -component "Self Update" -Message "User opted to download update" -severity 1
          start "http://www.skype4badmin.com/australian-holiday-rulesets-for-response-group-service/" #Todo Update URL
          Write-Log -component "Self Update" -Message "Exiting Script" -severity 3
          Exit
        }
        1 {Write-Log -component "Self Update" -Message "User opted to skip update" -severity 1
									
        }
							
      }
    }   
    Else{
      Write-Log -component "Self Update" -Message "Script is upto date" -severity 1
    }
        
  }


}

Function Read-ConfigFile {
  Write-Log -component "Read-ConfigFile" -Message "Reading Config file" -severity 2
  If(!(Test-Path $Script:ConfigFilePath)) {
    Write-Log -component "Read-ConfigFile" -Message "Could not locate config file!" -severity 3
    Write-Log -component "Read-ConfigFile" -Message "Error reading Config, Loading Defaults" -severity 3
    Load-DefaultConfig
  }
  Else {
    Write-Log -component "Read-ConfigFile" -Message "Found Config file in the specified folder" -severity 1
  }

  Write-Log -component "Read-ConfigFile" -Message "Pulling XML File" -severity 1
  [Void](Remove-Variable -Name Config -Scope Script -ErrorAction SilentlyContinue )
  Try{
    #Load the Config
    $Script:Config=@{}
    $Script:Config = (Import-CliXml -Path $Script:ConfigFilePath)
    Write-Log -component "Read-ConfigFile" -Message "Config File Read OK" -severity 2
    Update-AddonMenu

    #Update the Gui options
  
    #Populate with Values
    [void] $Global:grid_Tenants.Rows.Clear()
    [void] $Global:grid_Tenants.Rows.Add("1",$Script:Config.Tenant1.DisplayName,$Script:Config.Tenant1.SignInAddress,"****",$Script:Config.Tenant1.ModernAuth,$Script:Config.Tenant1.ConnectToTeams,$Script:Config.Tenant1.ConnectToSkype,$Script:Config.Tenant1.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("2",$Script:Config.Tenant2.DisplayName,$Script:Config.Tenant2.SignInAddress,"****",$Script:Config.Tenant2.ModernAuth,$Script:Config.Tenant2.ConnectToTeams,$Script:Config.Tenant2.ConnectToSkype,$Script:Config.Tenant2.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("3",$Script:Config.Tenant3.DisplayName,$Script:Config.Tenant3.SignInAddress,"****",$Script:Config.Tenant3.ModernAuth,$Script:Config.Tenant3.ConnectToTeams,$Script:Config.Tenant3.ConnectToSkype,$Script:Config.Tenant3.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("4",$Script:Config.Tenant4.DisplayName,$Script:Config.Tenant4.SignInAddress,"****",$Script:Config.Tenant4.ModernAuth,$Script:Config.Tenant4.ConnectToTeams,$Script:Config.Tenant4.ConnectToSkype,$Script:Config.Tenant4.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("5",$Script:Config.Tenant5.DisplayName,$Script:Config.Tenant5.SignInAddress,"****",$Script:Config.Tenant5.ModernAuth,$Script:Config.Tenant5.ConnectToTeams,$Script:Config.Tenant5.ConnectToSkype,$Script:Config.Tenant5.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("6",$Script:Config.Tenant6.DisplayName,$Script:Config.Tenant6.SignInAddress,"****",$Script:Config.Tenant6.ModernAuth,$Script:Config.Tenant6.ConnectToTeams,$Script:Config.Tenant6.ConnectToSkype,$Script:Config.Tenant6.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("7",$Script:Config.Tenant7.DisplayName,$Script:Config.Tenant7.SignInAddress,"****",$Script:Config.Tenant7.ModernAuth,$Script:Config.Tenant7.ConnectToTeams,$Script:Config.Tenant7.ConnectToSkype,$Script:Config.Tenant7.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("8",$Script:Config.Tenant8.DisplayName,$Script:Config.Tenant8.SignInAddress,"****",$Script:Config.Tenant8.ModernAuth,$Script:Config.Tenant8.ConnectToTeams,$Script:Config.Tenant8.ConnectToSkype,$Script:Config.Tenant8.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("9",$Script:Config.Tenant9.DisplayName,$Script:Config.Tenant9.SignInAddress,"****",$Script:Config.Tenant9.ModernAuth,$Script:Config.Tenant9.ConnectToTeams,$Script:Config.Tenant9.ConnectToSkype,$Script:Config.Tenant9.ConnectToExchange)
    [void] $Global:grid_Tenants.Rows.Add("10",$Script:Config.Tenant10.DisplayName,$Script:Config.Tenant10.SignInAddress,"****",$Script:Config.Tenant10.ModernAuth,$Script:Config.Tenant10.ConnectToTeams,$Script:Config.Tenant10.ConnectToSkype,$Script:Config.Tenant10.ConnectToExchange)
    }
  Catch {
    Write-Log -component "Read-ConfigFile" -Message "Error reading Config or updating GUI, Loading Defaults" -severity 3
    Load-DefaultConfig
  }

}

Function Write-ConfigFile {
  Write-Log -component "Write-ConfigFile" -Message "Writing Config file" -severity 2
  
  #Grab items from the GUI and stuff them into something useful

  $Script:Config.Tenant1.DisplayName = $Global:grid_Tenants.Rows[0].Cells[1].Value
  $Script:Config.Tenant1.SignInAddress = $Global:grid_Tenants.Rows[0].Cells[2].Value
  $Script:Config.Tenant1.ModernAuth = $Global:grid_Tenants.Rows[0].Cells[4].Value
  $Script:Config.Tenant1.ConnectToTeams = $Global:grid_Tenants.Rows[0].Cells[5].Value
  $Script:Config.Tenant1.ConnectToSkype = $Global:grid_Tenants.Rows[0].Cells[6].Value
  $Script:Config.Tenant1.ConnectToExchange = $Global:grid_Tenants.Rows[0].Cells[7].Value
 

  $Script:Config.Tenant2.DisplayName = $Global:grid_Tenants.Rows[1].Cells[1].Value
  $Script:Config.Tenant2.SignInAddress = $Global:grid_Tenants.Rows[1].Cells[2].Value
  $Script:Config.Tenant2.ModernAuth = $Global:grid_Tenants.Rows[1].Cells[4].Value
  $Script:Config.Tenant2.ConnectToTeams = $Global:grid_Tenants.Rows[1].Cells[5].Value
  $Script:Config.Tenant2.ConnectToSkype = $Global:grid_Tenants.Rows[1].Cells[6].Value
  $Script:Config.Tenant2.ConnectToExchange = $Global:grid_Tenants.Rows[1].Cells[7].Value


  $Script:Config.Tenant3.DisplayName = $Global:grid_Tenants.Rows[2].Cells[1].Value
  $Script:Config.Tenant3.SignInAddress = $Global:grid_Tenants.Rows[2].Cells[2].Value
  $Script:Config.Tenant3.ModernAuth = $Global:grid_Tenants.Rows[2].Cells[4].Value
  $Script:Config.Tenant3.ConnectToTeams = $Global:grid_Tenants.Rows[2].Cells[5].Value
  $Script:Config.Tenant3.ConnectToSkype = $Global:grid_Tenants.Rows[2].Cells[6].Value
  $Script:Config.Tenant3.ConnectToExchange = $Global:grid_Tenants.Rows[2].Cells[7].Value

 
  $Script:Config.Tenant4.DisplayName = $Global:grid_Tenants.Rows[3].Cells[1].Value
  $Script:Config.Tenant4.SignInAddress = $Global:grid_Tenants.Rows[3].Cells[2].Value
  $Script:Config.Tenant4.ModernAuth = $Global:grid_Tenants.Rows[3].Cells[4].Value
  $Script:Config.Tenant4.ConnectToTeams = $Global:grid_Tenants.Rows[3].Cells[5].Value
  $Script:Config.Tenant4.ConnectToSkype = $Global:grid_Tenants.Rows[3].Cells[6].Value
  $Script:Config.Tenant4.ConnectToExchange = $Global:grid_Tenants.Rows[3].Cells[7].Value

 
  $Script:Config.Tenant5.DisplayName = $Global:grid_Tenants.Rows[4].Cells[1].Value
  $Script:Config.Tenant5.SignInAddress = $Global:grid_Tenants.Rows[4].Cells[2].Value
  $Script:Config.Tenant5.ModernAuth = $Global:grid_Tenants.Rows[4].Cells[4].Value
  $Script:Config.Tenant5.ConnectToTeams = $Global:grid_Tenants.Rows[4].Cells[5].Value
  $Script:Config.Tenant5.ConnectToSkype = $Global:grid_Tenants.Rows[4].Cells[6].Value
  $Script:Config.Tenant5.ConnectToExchange = $Global:grid_Tenants.Rows[4].Cells[7].Value

 
  $Script:Config.Tenant6.DisplayName = $Global:grid_Tenants.Rows[5].Cells[1].Value
  $Script:Config.Tenant6.SignInAddress = $Global:grid_Tenants.Rows[5].Cells[2].Value
  $Script:Config.Tenant6.ModernAuth = $Global:grid_Tenants.Rows[5].Cells[4].Value
  $Script:Config.Tenant6.ConnectToTeams = $Global:grid_Tenants.Rows[5].Cells[5].Value
  $Script:Config.Tenant6.ConnectToSkype = $Global:grid_Tenants.Rows[5].Cells[6].Value
  $Script:Config.Tenant6.ConnectToExchange = $Global:grid_Tenants.Rows[5].Cells[7].Value

 
  $Script:Config.Tenant7.DisplayName = $Global:grid_Tenants.Rows[6].Cells[1].Value
  $Script:Config.Tenant7.SignInAddress = $Global:grid_Tenants.Rows[6].Cells[2].Value
  $Script:Config.Tenant7.ModernAuth = $Global:grid_Tenants.Rows[6].Cells[4].Value
  $Script:Config.Tenant7.ConnectToTeams = $Global:grid_Tenants.Rows[6].Cells[5].Value
  $Script:Config.Tenant7.ConnectToSkype = $Global:grid_Tenants.Rows[6].Cells[6].Value
  $Script:Config.Tenant7.ConnectToExchange = $Global:grid_Tenants.Rows[6].Cells[7].Value

 
  $Script:Config.Tenant8.DisplayName = $Global:grid_Tenants.Rows[7].Cells[1].Value
  $Script:Config.Tenant8.SignInAddress = $Global:grid_Tenants.Rows[7].Cells[2].Value
  $Script:Config.Tenant8.ModernAuth = $Global:grid_Tenants.Rows[7].Cells[4].Value
  $Script:Config.Tenant8.ConnectToTeams = $Global:grid_Tenants.Rows[7].Cells[5].Value
  $Script:Config.Tenant8.ConnectToSkype = $Global:grid_Tenants.Rows[7].Cells[6].Value
  $Script:Config.Tenant8.ConnectToExchange = $Global:grid_Tenants.Rows[7].Cells[7].Value

 
  $Script:Config.Tenant9.DisplayName = $Global:grid_Tenants.Rows[8].Cells[1].Value
  $Script:Config.Tenant9.SignInAddress = $Global:grid_Tenants.Rows[8].Cells[2].Value
  $Script:Config.Tenant9.ModernAuth = $Global:grid_Tenants.Rows[8].Cells[4].Value
  $Script:Config.Tenant9.ConnectToTeams = $Global:grid_Tenants.Rows[8].Cells[5].Value
  $Script:Config.Tenant9.ConnectToSkype = $Global:grid_Tenants.Rows[8].Cells[6].Value
  $Script:Config.Tenant9.ConnectToExchange = $Global:grid_Tenants.Rows[8].Cells[7].Value


  $Script:Config.Tenant10.DisplayName = $Global:grid_Tenants.Rows[9].Cells[1].Value
  $Script:Config.Tenant10.SignInAddress = $Global:grid_Tenants.Rows[9].Cells[2].Value
  $Script:Config.Tenant10.ModernAuth = $Global:grid_Tenants.Rows[9].Cells[4].Value
  $Script:Config.Tenant10.ConnectToTeams = $Global:grid_Tenants.Rows[9].Cells[5].Value
  $Script:Config.Tenant10.ConnectToSkype = $Global:grid_Tenants.Rows[9].Cells[6].Value
  $Script:Config.Tenant10.ConnectToExchange = $Global:grid_Tenants.Rows[9].Cells[7].Value

  #Encrypt passwords
  If ($Global:grid_Tenants.Rows[0].Cells[3].Value -ne "****") {$Script:Config.Tenant1.Credential = ($Global:grid_Tenants.Rows[0].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[1].Cells[3].Value -ne "****") {$Script:Config.Tenant2.Credential = ($Global:grid_Tenants.Rows[1].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[2].Cells[3].Value -ne "****") {$Script:Config.Tenant3.Credential = ($Global:grid_Tenants.Rows[2].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[3].Cells[3].Value -ne "****") {$Script:Config.Tenant4.Credential = ($Global:grid_Tenants.Rows[3].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[4].Cells[3].Value -ne "****") {$Script:Config.Tenant5.Credential = ($Global:grid_Tenants.Rows[4].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[5].Cells[3].Value -ne "****") {$Script:Config.Tenant6.Credential = ($Global:grid_Tenants.Rows[5].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[6].Cells[3].Value -ne "****") {$Script:Config.Tenant7.Credential = ($Global:grid_Tenants.Rows[6].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[7].Cells[3].Value -ne "****") {$Script:Config.Tenant8.Credential = ($Global:grid_Tenants.Rows[7].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[8].Cells[3].Value -ne "****") {$Script:Config.Tenant9.Credential = ($Global:grid_Tenants.Rows[8].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}
  If ($Global:grid_Tenants.Rows[9].Cells[3].Value -ne "****") {$Script:Config.Tenant10.Credential = ($Global:grid_Tenants.Rows[10].Cells[3].Value | ConvertTo-SecureString -AsPlainText -Force)}

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
  Try{
    $Script:Config| Export-CliXml -Path "$ENV:UserProfile\BounShell.xml"
    Write-Log -component "Write-ConfigFile" -Message "Config File Saved" -severity 2
  }
  Catch {
    Write-Log -component "Write-ConfigFile" -Message "Error writing Config file" -severity 3
  }


}

Function Load-DefaultConfig {
  #Set Variables to Defaults
  #Remove and re-create the Config Array
  [Void](Remove-Variable -Name Config -Scope Script -ErrorAction SilentlyContinue )
  $Script:Config=@{}
  #Populate with Defaults
  [void] $Global:grid_Tenants.Rows.Clear()
  
  $Script:Config.Tenant1 =@{}  
  $Script:Config.Tenant1.DisplayName = "Undefined"
  $Script:Config.Tenant1.SignInAddress = "user1@fabrikam.com"
  $Script:Config.Tenant1.Credential = "****"
  $Script:Config.Tenant1.ModernAuth = $false
  $Script:Config.Tenant1.ConnectToTeams = $false
  $Script:Config.Tenant1.ConnectToSkype = $false
  $Script:Config.Tenant1.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("1",'Undefined','user1@fabrikam.com',"****",$False,$false,$false,$false)
  
  
  $Script:Config.Tenant2 =@{}  
  $Script:Config.Tenant2.DisplayName = "Undefined"
  $Script:Config.Tenant2.SignInAddress = "user2@fabrikam.com"
  $Script:Config.Tenant2.Credential = "****"
  $Script:Config.Tenant2.ModernAuth = $false
  $Script:Config.Tenant2.ConnectToTeams = $false
  $Script:Config.Tenant2.ConnectToSkype = $false
  $Script:Config.Tenant2.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("2",'Undefined','user2@fabrikam.com',"****",$False,$false,$false,$false)

  $Script:Config.Tenant3 =@{}  
  $Script:Config.Tenant3.DisplayName = "Undefined"
  $Script:Config.Tenant3.SignInAddress = "user3@fabrikam.com"
  $Script:Config.Tenant3.Credential = "****"
  $Script:Config.Tenant3.ModernAuth = $false
  $Script:Config.Tenant3.ConnectToTeams = $false
  $Script:Config.Tenant3.ConnectToSkype = $false
  $Script:Config.Tenant3.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("3",'Undefined','user3@fabrikam.com',"****",$False,$false,$false,$false)

  $Script:Config.Tenant4 =@{}  
  $Script:Config.Tenant4.DisplayName = "Undefined"
  $Script:Config.Tenant4.SignInAddress = "user4@fabrikam.com"
  $Script:Config.Tenant4.Credential = "****"
  $Script:Config.Tenant4.ModernAuth = $false
  $Script:Config.Tenant4.ConnectToTeams = $false
  $Script:Config.Tenant4.ConnectToSkype = $false
  $Script:Config.Tenant4.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("4",'Undefined','user4@fabrikam.com',"****",$False,$false,$false,$false)

  $Script:Config.Tenant5 =@{}  
  $Script:Config.Tenant5.DisplayName = "Undefined"
  $Script:Config.Tenant5.SignInAddress = "user5@fabrikam.com"
  $Script:Config.Tenant5.Credential = "****"
  $Script:Config.Tenant5.ModernAuth = $false
  $Script:Config.Tenant5.ConnectToTeams = $false
  $Script:Config.Tenant5.ConnectToSkype = $false
  $Script:Config.Tenant5.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("5",'Undefined','user5@fabrikam.com',"****",$False,$false,$false,$false)

  $Script:Config.Tenant6 =@{}  
  $Script:Config.Tenant6.DisplayName = "Undefined"
  $Script:Config.Tenant6.SignInAddress = "user6@fabrikam.com"
  $Script:Config.Tenant6.Credential = "****"
  $Script:Config.Tenant6.ModernAuth = $false
  $Script:Config.Tenant6.ConnectToTeams = $false
  $Script:Config.Tenant6.ConnectToSkype = $false
  $Script:Config.Tenant6.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("6",'Undefined','user6@fabrikam.com',"****",$False,$false,$false,$false)

  $Script:Config.Tenant7 =@{}  
  $Script:Config.Tenant7.DisplayName = "Undefined"
  $Script:Config.Tenant7.SignInAddress = "user@fabrikam.com"
  $Script:Config.Tenant7.Credential = "****"
  $Script:Config.Tenant7.ModernAuth = $false
  $Script:Config.Tenant7.ConnectToTeams = $false
  $Script:Config.Tenant7.ConnectToSkype = $false
  $Script:Config.Tenant7.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("7",'Undefined','user7@fabrikam.com',"****",$False,$false,$false,$false)

  $Script:Config.Tenant8 =@{}  
  $Script:Config.Tenant8.DisplayName = "Undefined"
  $Script:Config.Tenant8.SignInAddress = "user8@fabrikam.com"
  $Script:Config.Tenant8.Credential = "****"
  $Script:Config.Tenant8.ModernAuth = $false
  $Script:Config.Tenant8.ConnectToTeams = $false
  $Script:Config.Tenant8.ConnectToSkype = $false
  $Script:Config.Tenant8.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("8",'Undefined','user8@fabrikam.com',"****",$False,$false,$false,$false)
    
  $Script:Config.Tenant9 =@{}  
  $Script:Config.Tenant9.DisplayName = "Undefined"
  $Script:Config.Tenant9.SignInAddress = "user@fabrikam.com"
  $Script:Config.Tenant9.Credential = "****"
  $Script:Config.Tenant9.ModernAuth = $false
  $Script:Config.Tenant9.ConnectToTeams = $false
  $Script:Config.Tenant9.ConnectToSkype = $false
  $Script:Config.Tenant9.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("9",'Undefined','user9@fabrikam.com',"****",$False,$false,$false,$false)
  
  $Script:Config.Tenant10 =@{}  
  $Script:Config.Tenant10.DisplayName = "Undefined"
  $Script:Config.Tenant10.SignInAddress = "user@fabrikam.com"
  $Script:Config.Tenant10.Credential = "****"
  $Script:Config.Tenant10.ModernAuth = $false
  $Script:Config.Tenant10.ConnectToTeams = $false
  $Script:Config.Tenant10.ConnectToSkype = $false
  $Script:Config.Tenant10.ConnectToExchange = $false
  [void] $Global:grid_Tenants.Rows.Add("10",'Undefined','user10@fabrikam.com',"****",$False,$false,$false,$false)
    
  [Float]$Script:Config.ConfigFileVersion = "0.1"
  [string]$Script:Config.Description = "BounShell Configuration file. See Skype4BAdmin.com for more information"
  
  
	
}

Function Invoke-NewTenantTab {
  <#
      .SYNOPSIS
      Function to Open new tab in ISE. Connect to a PSsession and import it

      .DESCRIPTION
        

      .PARAMETER Tenant
      The message to write

      .PARAMETER Username
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
      http://www.skype4badmin.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>
  param(
  [Parameter(Mandatory=$true)] [string]$Tabname,
  [Parameter(Mandatory=$true)] [float]$Tenant

    )    
  #kick off a new tab and call it tabname
  $TabNameTab=$psISE.PowerShellTabs.Add()
  $TabNameTab.DisplayName = $Tabname
    
  #Wait for the tab to wake up
  Do 
  {sleep -m 100}
  While (!$TabNameTab.CanInvoke)
    
  #Kick off the connection
  $TabNameTab.Invoke("$PSScriptRoot.\Start-BounShell.ps1 -tenant 1")
  #$TabNameTab.invoke($scriptblock -argumentlist $Pscred)
}

Function Connect-O365Tenant {

  [CmdletBinding()]
  PARAM(
    $Tenant
  )
  $Function= 'Connect-O365Tenant'

  switch ($Tenant)
      {
        1 
        {
            Write-Log -component $function -Message "Connecting to $($Script:Config.Tenant1.DisplayName)" -severity 3
            If (!$Script:Config.Tenant1.ModernAuth) {
                $Script:pscred = New-Object System.Management.Automation.PSCredential($Script:Config.Tenant1.SignInAddress,$Script:Config.Tenant1.Credential)
                }
            Else{
                Write-Log -component $function -Message "Modern Auth Not available in Beta" -severity 3 
                Pause
                Exit
                }   
        }
        2 
        {
            Write-Log -component $function -Message "Connecting to $($Script:Config.Tenant2.DisplayName)" -severity 3
            If (!$Script:Config.Tenant2.ModernAuth) {
                $Script:pscred = New-Object System.Management.Automation.PSCredential($Script:Config.Tenant2.SignInAddress,$Script:Config.Tenant2.Credential)
                }
            Else{
                Write-Log -component $function -Message "Modern Auth Not available in Beta" -severity 3 
                Pause
                Exit
                }   
        }
    }



  #See if we got passed creds
  Write-Log -Message 'Checking for Office365 Credentials' -Severity 1 -Component $Function
  If ($pscred -eq $null) {
    Write-Log -Message 'No Office365 credentials Found, Prompting user for creds' -Severity 3 -Component $Function
  $psCred = Get-Credential}
  Else{
    Write-Log -Message "Found Office365 Creds for Username: $($pscred.username)" -Severity 2 -Component $Function
      
  }
  $Function= 'Connect-O365Tenant-block'
  Get-PSSession | Remove-PSSession
  #Write-Log -Message "Connecting to $tenant Tenant" -Severity 2 -Component $Function
    
    
    
  #`$PSCred = `"$($PScred)` 
  #Exchange section
  If ($O365Session -eq $null) {
    Try {
      $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $pscred -Authentication Basic -AllowRedirection
      (Import-Module -Name (Import-PSSession -Session $O365Session -AllowClobber -DisableNameChecking) -Global)
    } 
    Catch {
      Write-output -Message 'Error connecting to Exchange in Office 365'
    }
  }
   

  #Now we check to see if there is an existing Skype Online Session running already
  If ($S4BOSession -eq $null) {
    Try {
      $S4BOSession = New-CsOnlineSession -Credential $pscred 
      (Import-Module -Name (Import-PSSession -Session $S4BOSession) -Global)
    } 
    Catch {
      #Write-host -Message 'Error connecting to Skype for Business Online' 
      
    }
  }
       
  
}

Function Update-AddonMenu {

 
  #Check to see if we are loaded, if we are cleanup after ourselves
  if (($psISE.CurrentPowerShellTab.AddOnsMenu.Submenus).displayname -eq "_BounShell") {
    [void]$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.remove($Global:isemenuitem)
  }


  #Create Initial Menu Object
  [void]($Global:IseMenuItem = ($psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('_BounShell',$null ,$null)))

  #Add the Settings Button

  [void]($Global:IseMenuItem.Submenus.add('_Settings...', {Show-GuiElements}, $null) )


  #Now add each Tenant

  #For each code in here that adds Tenant 1 through 10
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant1.DisplayName)",{Invoke-NewTenantTab -tabname $Script:Config.Tenant1.DisplayName -Tenant 1}, 'Ctrl+Alt+1'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant2.DisplayName)",{Connect-O365Tenant -Tenant "2"}, 'Ctrl+Alt+2'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant3.DisplayName)",{Connect-O365Tenant -Tenant "3"}, 'Ctrl+Alt+3'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant4.DisplayName)",{Connect-O365Tenant -Tenant "4"}, 'Ctrl+Alt+4'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant5.DisplayName)",{Connect-O365Tenant -Tenant "5"}, 'Ctrl+Alt+5'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant6.DisplayName)",{Connect-O365Tenant -Tenant "6"}, 'Ctrl+Alt+6'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant7.DisplayName)",{Connect-O365Tenant -Tenant "7"}, 'Ctrl+Alt+7'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant8.DisplayName)",{Connect-O365Tenant -Tenant "8"}, 'Ctrl+Alt+8'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant9.DisplayName)",{Connect-O365Tenant -Tenant "9"}, 'Ctrl+Alt+9'))
  [void]($Global:IseMenuItem.Submenus.add("$($Script:Config.Tenant10.DisplayName)",{Connect-O365Tenant -Tenant "10"}, 'Ctrl+Alt+0'))             


}

Function Test-ManagementTools {
  Write-Log -component "Test-ManagementTools" -Message "Checking for Lync/Skype management tools"
  $CSManagementTools = $false
  if(!(Get-Module "SkypeForBusiness")) {Import-Module SkypeForBusiness -Verbose:$false}
  if(!(Get-Module "Lync")) {Import-Module Lync -Verbose:$false}
  if(Get-Module "SkypeForBusiness") {$ManagementTools = $true}
  if(Get-Module "Lync") {$ManagementTools = $true}
  if(!$CSManagementTools) {
    Write-Log 
    Write-Log -component "Test-ManagementTools" -Message "Could not locate Lync/Skype4B Management tools" -severity 3 
    throw  "Could not locate Lync/Skype4B Management tools"
  }
	
  #Check for the AD Management Tools
  $ADManagementTools = $false
  if(!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory -Verbose:$false}
  if(Get-Module "ActiveDirectory") {$ADManagementTools = $true}
  if(!$ADManagementTools) {
    Write-Log 
    Write-Log -component "Test-ManagementTools" -Message "Could not locate Active Directory Management tools" -severity 3
    throw  "Could not locate Active Directory Management tools"
  }

  $TeamsManagementTools = $false
  if(!(Get-Module "MicrosoftTeams")) {Import-Module MicrosoftTeams -Verbose:$false}
  if(Get-Module "MicrosoftTeams") {$TeamsManagementTools = $true}
  if(!$TeamsManagementTools) {
    Write-Log 
    Write-Log -component "Test-CsManagementTools" -Message "Could not locate Teams PowerShell Module" -severity 3 
    Write-Log -component "Test-CsManagementTools" -Message "Run the cmdlet 'Install-Module MicrosoftTeams' to install it" -severity 3 
      
  }
}

Function Import-GuiElements {
    #First we need to import the Functions so they exist for the GUI items
    Import-GuiFunctions

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
  $Global:btn_CancelConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
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
  $Global:Btn_ReloadConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
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
  $Global:Btn_SaveConfig.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
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
  $Global:dataGridViewCellStyle1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle1.ForeColor = [System.Drawing.SystemColors]::WindowText
  $Global:dataGridViewCellStyle1.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle1.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle1.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
  $Global:grid_Tenants.ColumnHeadersDefaultCellStyle = $Global:dataGridViewCellStyle1
  $Global:grid_Tenants.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
  $Global:grid_Tenants.Columns.AddRange($Global:Tenant_ID,$Global:Tenant_DisplayName,$Global:Tenant_Email,$Global:Tenant_Credentials,$Global:Tenant_ModernAuth,$Global:Tenant_Teams,$Global:Tenant_Skype,$Global:Tenant_Exchange)
  $Global:dataGridViewCellStyle2.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $Global:dataGridViewCellStyle2.BackColor = [System.Drawing.SystemColors]::Window
  $Global:dataGridViewCellStyle2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
  $Global:dataGridViewCellStyle2.ForeColor = [System.Drawing.Color]::FromArgb(([System.Int32]([System.Byte][System.Byte]8)),([System.Int32]([System.Byte][System.Byte]116)),([System.Int32]([System.Byte][System.Byte]170)))

  $Global:dataGridViewCellStyle2.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
  $Global:dataGridViewCellStyle2.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
  $Global:dataGridViewCellStyle2.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
  $Global:grid_Tenants.DefaultCellStyle = $Global:dataGridViewCellStyle2
  $Global:grid_Tenants.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]12))
  $Global:grid_Tenants.Name = [System.String]'grid_Tenants'
  $Global:dataGridViewCellStyle3.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $Global:dataGridViewCellStyle3.BackColor = [System.Drawing.SystemColors]::Control
  $Global:dataGridViewCellStyle3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
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
  $Global:Btn_Default.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
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

Function Import-GuiFunctions {
    #Gui Cancel button
    $Global:btn_CancelConfig_Click = {
        Read-ConfigFile
        [void]$Global:SettingsForm.Hide()
    }

    #Gui Save Config Button
    $Global:Btn_SaveConfig_Click = {
        Write-ConfigFile
        Update-AddonMenu
    }

    #Gui Set Defaults Button
    $Global:Btn_Default_Click = {
        Load-DefaultConfig
        Update-AddonMenu
    }

    #Gui Button to Reload Config
    $Global:Btn_ConfigReload_Click = {
        Read-ConfigFile
        Update-AddonMenu
    }


}

Function Show-GuiElements {

  [void]$Global:SettingsForm.ShowDialog()
}

Function Hide-GuiElements {

  [void]$Global:SettingsForm.Hide()
}

Function Start-MainScript {
  #Allows us to seperate all the "onetime" run objects incase we get dot sourced.
  Write-Log -component "Script Block" -Message "Script executed from $PSScriptRoot" -severity 1
  Write-Log -component "Startup" -Message "Loading BounShell..." -severity 2


  #Load the Gui Elements
  Import-GuiElements
  
  #check for script update
  #if ($DisableScriptUpdate -eq $false) {Get-ScriptUpdate} #todo enable update checking

  #check for config file then load the default

  #Check for and load the config file if present
  If(Test-Path $Script:ConfigFilePath) 
  {   
    Write-Log -component "Config" -Message "Found $ConfigFilePath, loading..." -severity 1
    Read-ConfigFile
  }

  Else {
    Write-Log -component "Config" -Message "Could not locate $ConfigFilePath, Using Defaults" -severity 3
    #If there is no config file. Load a default
    Load-DefaultConfig

    Write-Log -component "Config" -Message "As we didnt find a config file we will assume this is a first run." -severity 3
    Write-Log -component "Config" -Message "Thus we will remind you that while all care is taken to store your credentials in a safe manner, I cannot be held responsible for any data breaches" -severity 3
    Write-Log -component "Config" -Message "If someone was to get a hold of your BounShell.xml AND your user profile private encryption key its possible to reverse engineer stored credentials" -severity 3
    Write-Log -component "Config" -Message "Seriously, Whilst the password store is encrypted, its not perfect!" -severity 3
  }

  #Check for Management Tools
  #Test-ManagementTools #todo fix
  
  #Make everything available to ourselves.
  . "$PSScriptRoot.\Start-BounShell.ps1"
  
  #Now Create the Objects in the ISE
  Update-AddonMenu

}

#Start script, check to see if we need to setup the enviroment, if we are dot sourcing ourselves or attempting to load a new tab
if (!($Tenant)){
    if ($MyInvocation.InvocationName -eq '&') {
      Write-Log -component "Startup" -Message "Called using operator - Not supported. YMMV" -severity 3
    
    } elseif ($MyInvocation.InvocationName -eq '.') {
      Write-Log -component "Startup" -Message "Dot sourced, skipping setup" -severity 1
  
    } elseif ((Resolve-Path -Path $MyInvocation.InvocationName).ProviderPath -eq ($MyInvocation.MyCommand.Path)) {
      Start-MainScript
    }
}
Else {
   Write-Log -component "Startup" -Message "Called to connect to Tenant $tenant" -severity 3
   Import-GuiElements
   Read-ConfigFile
   Connect-O365Tenant -Tenant $Tenant
   }