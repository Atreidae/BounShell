<#
    .SYNOPSIS

    This is a tool to help users manage multiple office 365 tennants

    .DESCRIPTION

    Created by James Arber. www.skype4badmin.com
    
    .NOTES

    Version      	        : 0.1
    Date			        : 18/09/2018
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



# Script Config

  #############################
  # Script Specific Variables #
  #############################

  $ScriptVersion              =  0.1
  $StartTime                  =  Get-Date
  $Script:LogFileLocation     =  $PSCommandPath -replace '.ps1','.log' #Where do we store the log files? (In the same folder by default)
  $Script:ConfigPath          =  $PSCommandPath -replace '.ps1','.json' #Where do we store the Config files? (In the same folder by default)
  $VerbosePreference="Continue"


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
  } #end WriteLog
  
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

  Function Read-ConfigFile {
  Write-Log -component "Read-ConfigFile" -Message "Reading Config file" -severity 2
    If(!(Test-Path $Script:ConfigPath)) {
      Write-Log -component "Read-ConfigFile" -Message "Could not locate config file!" -severity 3
      Write-Log -component "Read-ConfigFile" -Message "Error reading Config or Key file, Loading Defaults" -severity 3
      Load-DefaultConfig
      }
    Else {
      Write-Log -component "Read-ConfigFile" -Message "Found Config file in the specified folder" -severity 1
        }

  Write-Log -component "Read-ConfigFile" -Message "Pulling JSON File" -severity 1
  [Void](Remove-Variable -Name Config -Scope Script )
    Try{
      $Script:Config=@{}
      $Script:Config = (Import-CliXml -Path "$ENV:UserProfile\BounShell.xml")
      Write-Log -component "Read-ConfigFile" -Message "Config File Read OK" -severity 2
      }
    Catch {
    Write-Log -component "Read-ConfigFile" -Message "Error reading Config or Key file, Loading Defaults" -severity 3
    Load-DefaultConfig
      }

}

  Function Write-ConfigFile {
  Write-Log -component "Write-ConfigFile" -Message "Writing Config file" -severity 2

  #Write the Json File
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
    [Void](Remove-Variable -Name Config -Scope Script )
    $Script:Config=@{}
    #Populate with Defaults
    
    $Script:Config.Tenant1.DisplayName = "IComm"
    $Script:Config.Tenant1.Credential = $null
    
    $Script:Config.Tenant2.DisplayName = "Skype4BAdmin"
    $Script:Config.Tenant2.Credential = $null
    
    [Float]$Script:Config.ConfigFileVersion = "0.1"
    [string]$Script:Config.Description = "BounShell Configuration file. See Skype4BAdmin.com for more information"
    
    #Now Create the Objects in the ISE
    #Config Page
    $Txt_BotSipAddr.Text = $Script:Config.BotAddress
    $tbx_Autodiscover.text = $Script:Config.AutoDiscover
    $txt_DomainFQDN.text = $Script:Config.DomainFQDN
    $mtxt_MinUpdate.text = $Script:Config.MinUpdate
    $mtxt_MaxChanges.text = $Script:Config.MaxChanges
    
    #Main Page
    $dbx_FePool.text = $Script:Config.SelectedFePool
    $dbx_LocRule.text = $Script:Config.SelectedRule
	
  }

  Function Connect-O365Tenant {
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
    [CmdletBinding()]
    PARAM(
      $pscred,
      $Tabname
    )
    $Function= 'Connect-O365Tenant'
    #See if we got passed creds
    Write-Log -Message 'Checking for Office365 Credentials' -Severity 1 -Component $Function
    If ($pscred -eq $null) {
      Write-Log -Message 'No Office365 credentials Found, Prompting user for creds' -Severity 3 -Component $Function
    $psCred = Get-Credential}
    Else{
      Write-Log -Message "Found Office365 Creds for Username: $pscred.username" -Severity 2 -Component $Function
      
    }

    #Store the invocation code in a script block for execution
    
    $Scriptblock = {
                      
                      $Function= 'Connect-O365Tenant-block'
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
                        Else{
                          #Write-Log -Message 'Found Existing Skype4B Online Remote Session' -Severity 1 -Component $Function
                        }
       
  
                      }
      #kick off a new tab and call it tabname
    $TabNameTab=$psISE.PowerShellTabs.Add()
    $TabNameTab.DisplayName = $Tabname
    #Wait for the tab to wake up
    Do 
    {sleep -m 100}
    While (!$TabNameTab.CanInvoke)
    
    #Kick off the connection
    $TabNameTab.InvokeSynchronous($scriptblock, $false)
    #$TabNameTab.invoke($scriptblock -argumentlist $Pscred)
   }
                    
                    
