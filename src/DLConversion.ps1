#requires -version 2
<#

.DISCLAIMER

    THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT
    PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY
    OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT
    LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR
    PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS
    AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR
    ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE
    FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS
    PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)
    ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,
    EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES 


.SYNOPSIS
This script is designed to assist users in migrating distributon lists to Office 365.
	  
************************************************************************************************
************************************************************************************************
************************************************************************************************
The PSLOGGING module must be installed / updated prior to use.
---install-module PSLOGGING
---update-module PSLOGGING.
************************************************************************************************
************************************************************************************************
************************************************************************************************

Information regarding module can be found here:

https://www.powershellgallery.com/packages/PSLogging/2.5.2
https://9to5it.com/powershell-logging-v2-easily-create-log-files/

All credits for the logging module to the owner and many thanks for allowing it's public use.

.DESCRIPTION

The script performs the following functions:

1)	Log into remote powershell to On-Premises Exchange.
2)	Log into remote powershell to Office 365 / Exchange Online and prefix the command with o365.

Note:	This may require enabling basic authentication on one or more endpoints for powershell remoting.

3)	Obtain the distribution list name from the command line.
4)	Validate the DL exists on premsies.	[Hard fail if not found]
5)	Validate the DL exists in Office 365. [Hard fail if not found]
6)	Capture the distribution list membership on premises and validate that each recipient exists in Exchange Online [Hard fail if recipient not found]
7)	Capture all multivalued attributs of the distribution list, convert the list to primary SMTP addresses, and validate those recipients exist in Exchange Online [Hard fail if recipient not found]
8)	Move the distribution list to a non-sync OU.
9)	Trigger remotely a delta sync of AD connect to remove the distribution list from Exchange Online.
10)	Validate the DL was removed from Exchange Online.
11)	Recreate the distribution list directly in Exchange Online.
12)	Stamp all attributes of the on-premises DL to the new DL in Exchange Online.
13)	Stamp the original legacyExchangeDN to the new DL as an X500 address to preserve reply to functionality.

.INPUTS

Credential files are required to be created prior to script execution as secure string XML files.
Administrator is responsible adjusting the following variables to meet their environment.

Variables proceeded by ##ADMIN##

.OUTPUTS

Operations log file.
XML files for backup of on premises and cloud DLs.

.NOTES

Version:        1.0
Author:         Timothy J. McMichael
Creation Date:  September 5th, 2018
Purpose/Change: Initial script development
  
.EXAMPLE

DLConversion -dlToConvert dl@domain.com -ignoreInvalidDLMembers:$TRUE -ignoreInvalidManagedByMembers:$TRUE
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
    #Script parameters go here

	[Parameter(Mandatory=$True,Position=1)]
    [string]$dlToConvert,
    [Parameter(Mandatory=$True,Position=2)]
    [boolean]$ignoreInvalidDLMember=$FALSE,
    [Parameter(Mandatory=$True,Position=3)]
	[boolean]$ignoreInvalidManagedByMember=$FALSE
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

Import-Module PSLogging

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Log File Info
<###ADMIN###>$script:sLogPath = "C:\Scripts\"
<###ADMIN###>$script:sLogName = "DLConversion.log"
<###ADMIN###>$script:sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#Establish credential information.

<###ADMIN###>$script:credentialFilePath = "C:\Scripts\"  #Path to the local credential files.
<###ADMIN###>$script:onPremisesCredentialFileName = "OnPremises-Credentials.cred"  #OnPremises credential XML file.
<###ADMIN###>$script:office365CredentialFileName = "Office365-Credentials.cred"  #Office365 credential XML file.
$script:onPremisesCredentialFile = Join-Path -Path $script:credentialFilePath -ChildPath $script:onPremisesCredentialFileName  #Full path and file name to onpremises credential file.
$script:office365CredentialFile = Join-Path -Path $script:credentialFilePath -ChildPath $script:office365CredentialFileName  #Full path and file name to Office 365 credential file.
$script:onPremisesCredential = $NULL  #Onpremises credentials
$script:office365Credential = $NULL  #Office 365 creentials

#Establish powershell server references.

<###ADMIN###>$script:onPremisesAADConnectServer = "server.domain.local"  #FQDN of the local AD Connect instance.
<###ADMIN###>$script:onPremisesPowerShell = "https://server.domain.com/powershell"  #URL of on premises powershell instance.
$script:office365Powershell = "https://outlook.office365.com/powershell-liveID/"  #URL of Office 365 powershell instance.

#Establish global variables for powershell sessions.

$script:office365PowershellConfiguration =  "Microsoft.Exchange" #Office 365 powershell configuration.
$script:onPremisesPowershellConfiguration =  "Microsoft.Exchange" #Exchange ON premises powershell configuration.
$script:office365PowershellAuthentication = "Basic" #Office 365 powershell authentication.
$script:onPremisesPowershellAuthentication = "Basic" #Exchange on premises powershell authentication.
$script:office365PowerShellSession = $null #Office 365 powershell session.
$script:onPremisesPowerShellSession = $null #Exchange on premises powershell session.
$script:onPremisesAADConnectPowerShellSession = $null #On premises aad connect powershell session.

#Establish script variables for distribution list operations.

$script:onpremisesdlConfiguration = $NULL  #Gathers the on premises DL information to a variable.
$script:office365DLConfiguration = $NULL  #Gather the Office 365 DL information to a variable.
$script:onpremisesdlconfigurationMembership = $null #On premise dl membership.
$script:newOffice365DLConfiguration = $NULL
$script:newOffice365DLConfigurationMembership = $NULL
[array]$script:onpremisesdlconfigurationMembershipArray = @() #Array of psObjects that represent DL membership.
[array]$script:onpremisesdlconfigurationManagedByArray = @() #Array of psObjects that represent managed by membership.
[array]$script:onpremisesdlconfigurationModeratedByArray = @() #Array of psObjects that represent moderated by membership.
[array]$script:onpremisesdlconfigurationGrantSendOnBehalfTOArray = @() #Array of psObjects that represent grant send on behalf to membership.
[array]$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers = @() #Array of psObjects that represent accept messages only from senders or members membership.
[array]$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers = @() #Array of psObjects that represent reject messages only from senders or members membership.
[array]$script:onPremsiesDLBypassModerationFromSendersOrMembers = @() #Array of psObjects that represent bypass moderation only from senders or members membership.
$script:newOffice365DLConfiguration=$NULL

#Establish script variables for active directory operations.

<###ADMIN###>$script:groupOrganizationalUnit = "OU=ConvertedDL,DC=DOMAIN,DC=LOCAL" #OU to move migrated DLs too.
$script:adDomainControllers = $null #List of domain controllers in domain.
$script:adDomain=$null #AD domain.
<###ADMIN###>[int32]$script:adDomainReplicationTime = 1 #Timeout to wait and allow for ad replication.
<###ADMIN###>[int32]$script:dlDelectionTime = 1 #Timeout to wait before rechecking for deleted DL.

#Establish script variables to backup distribution list information.

<###ADMIN###>$script:backupXMLPath = "C:\Scripts\" #Location of backup XML files.
<###ADMIN###>$script:onpremisesdlconfigurationXMLName = "onpremisesdlConfiguration.XML" #On premises XML file name.
<###ADMIN###>$script:office365DLXMLName = "office365DLConfiguration.XML" #Cloud XML file name.
<###ADMIN###>$script:onPremsiesDLConfigurationMembershipXMLName = "onpremisesDLConfigurationMembership.XML"
<###ADMIN###>$script:newOffice365DLConfigurationXMLName = "newOffice365DLConfiguration.XML"
<###ADMIN###>$script:newOffice365DLConfigurationMembershipXMLName = "newOffice365DLConfigurationMembership.XML"
$script:onPremisesXML = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationXMLName #Full path to on premises XML.
$script:office365XML = Join-Path $script:backupXMLPath -ChildPath $script:office365DLXMLName #Full path to cloud XML.
$script:onPremsiesMembershipXML = Join-Path $script:backupXMLPath -ChildPath $script:onPremsiesDLConfigurationMembershipXMLName
$script:newOffice365XML = Join-Path $script:backupXMLPath -ChildPath $script:newOffice365DLConfigurationXMLName
$script:newOffice365MembershipXML = Join-Path $script:backupXMLPath -ChildPath $script:newOffice365DLConfigurationMembershipXMLName

#Establish misc.

$script:aadconnectRetryRequired = $FALSE #Determines if ad connect sync retry is required.
$script:dlDeletionRetryRequired = $FALSE #Determines if deleted DL retry is required.
[int]$script:forCounter = $NULL #Counter utilized 
<###ADMIN###>[int32]$script:refreshCounter=1000

#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
*******************************************************************************************************

Function Start-PSCountdown

.DESCRIPTION

This function starts a visual countdown timer to show progress at scheduled waits.

All credits to the author

https://gist.github.com/jdhitsolutions/2e58d1aa41f684408b64488259bbeed0

.PARAMETER 

Multiple defined by author.

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
function Start-PSCountdown {
    [cmdletbinding()]
    [OutputType("None")]
    Param(
        [Parameter(Position = 0, HelpMessage = "Enter the number of minutes to countdown (1-60). The default is 5.")]
        [ValidateRange(1, 60)]
        [int32]$Minutes = 5,
        [Parameter(HelpMessage = "Enter the text for the progress bar title.")]
        [ValidateNotNullorEmpty()]
        [string]$Title = "Counting Down ",
        [Parameter(Position = 1, HelpMessage = "Enter a primary message to display in the parent window.")]
        [ValidateNotNullorEmpty()]
        [string]$Message = "Starting soon.",
        [Parameter(HelpMessage = "Use this parameter to clear the screen prior to starting the countdown.")]
        [switch]$ClearHost 
    )
    DynamicParam {
        #this doesn't appear to work in PowerShell core on Linux
        if ($host.PrivateData.ProgressBackgroundColor -And ( $PSVersionTable.Platform -eq 'Win32NT' -OR $PSEdition -eq 'Desktop')) {
            #define a parameter attribute object
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.ValueFromPipelineByPropertyName = $False
            $attributes.Mandatory = $false
            $attributes.HelpMessage = @"
Select a progress bar style. This only applies when using the PowerShell console or ISE.           
Default - use the current value of `$host.PrivateData.ProgressBarBackgroundColor
Transparent - set the progress bar background color to the same as the console
Random - randomly cycle through a list of console colors
"@
            #define a collection for attributes
            $attributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)
            #define the validate set attribute
            $validate = [System.Management.Automation.ValidateSetAttribute]::new("Default", "Random", "Transparent")
            $attributeCollection.Add($validate)
            #add an alias
            $alias = [System.Management.Automation.AliasAttribute]::new("style")
            $attributeCollection.Add($alias)
            #define the dynamic param
            $dynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ProgressStyle", [string], $attributeCollection)
            $dynParam1.Value = "Default"
            #create array of dynamic parameters
            $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add("ProgressStyle", $dynParam1)
            #use the array
            return $paramDictionary     
        } #if
    } #dynamic parameter
    Begin {
        $loading = @(
            'Sleeping'
        )
		if ($ClearHost)
		{
            Clear-Host
        }
        $PSBoundParameters | out-string | Write-Verbose
        if ($psboundparameters.ContainsKey('progressStyle')) { 
            if ($PSBoundParameters.Item('ProgressStyle') -ne 'default') {
                $saved = $host.PrivateData.ProgressBackgroundColor 
            }
            if ($PSBoundParameters.Item('ProgressStyle') -eq 'transparent') {
                $host.PrivateData.progressBackgroundColor = $host.ui.RawUI.BackgroundColor
            }
        }
        $startTime = Get-Date
        $endTime = $startTime.AddMinutes($Minutes)
        $totalSeconds = (New-TimeSpan -Start $startTime -End $endTime).TotalSeconds
        $totalSecondsChild = Get-Random -Minimum 4 -Maximum 30
        $startTimeChild = $startTime
        $endTimeChild = $startTimeChild.AddSeconds($totalSecondsChild)
        $loadingMessage = $loading[(Get-Random -Minimum 0 -Maximum ($loading.Length - 1))]
        #used when progress style is random
        $progcolors = "black", "darkgreen", "magenta", "blue", "darkgray"
    } #begin
    Process {
        #this does not work in VS Code
        if ($host.name -match 'Visual Studio Code') {
            Write-Warning "This command will not work in VS Code."
            #bail out
            Return
        }
        Do {   
            $now = Get-Date
            $secondsElapsed = (New-TimeSpan -Start $startTime -End $now).TotalSeconds
            $secondsRemaining = $totalSeconds - $secondsElapsed
            $percentDone = ($secondsElapsed / $totalSeconds) * 100
            Write-Progress -id 0 -Activity $Title -Status $Message -PercentComplete $percentDone -SecondsRemaining $secondsRemaining
            $secondsElapsedChild = (New-TimeSpan -Start $startTimeChild -End $now).TotalSeconds
            $secondsRemainingChild = $totalSecondsChild - $secondsElapsedChild
            $percentDoneChild = ($secondsElapsedChild / $totalSecondsChild) * 100
            if ($percentDoneChild -le 100) {
                Write-Progress -id 1 -ParentId 0 -Activity $loadingMessage -PercentComplete $percentDoneChild -SecondsRemaining $secondsRemainingChild
            }
            if ($percentDoneChild -ge 100 -and $percentDone -le 98) {
                if ($PSBoundParameters.ContainsKey('ProgressStyle') -AND $PSBoundParameters.Item('ProgressStyle') -eq 'random') {
                    $host.PrivateData.progressBackgroundColor = ($progcolors | Get-Random)
                }
                $totalSecondsChild = Get-Random -Minimum 4 -Maximum 30
                $startTimeChild = $now
                $endTimeChild = $startTimeChild.AddSeconds($totalSecondsChild)
                if ($endTimeChild -gt $endTime) {
                    $endTimeChild = $endTime
                }
                $loadingMessage = $loading[(Get-Random -Minimum 0 -Maximum ($loading.Length - 1))]
            }
            Start-Sleep 0.2
        } Until ($now -ge $endTime)
    } #progress
    End {
        if ($saved) {
            #restore value if it has been changed
            $host.PrivateData.ProgressBackgroundColor = $saved
        }
    } #end
} #end function

<#
*******************************************************************************************************

Function cleanupSessions

.DESCRIPTION

Removes all powershell sessions created by the script.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function cleanupSessions
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function cleans up all powershell sessions....' -toscreen
	}
	Process 
	{
		Try 
		{
			Get-PSSession | Remove-PSSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All powershell sessions have been cleaned up successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "The powershell sessions could not be cleared - manual removal before restarting required" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function removeOffice365PowerShellSession

.DESCRIPTION

Removes only the powershell session associated with Office 365.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function removeOffice365PowerShellSession
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function removes the Office 365 powershell sessions....' -toscreen
	}
	Process 
	{
		Try 
		{
			remove-pssession $script:office365PowerShellSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All powershell sessions have been cleaned up successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "The powershell sessions could not be cleared - manual removal before restarting required" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function refreshOffice365PowerShellSession

.DESCRIPTION

Removes, creates, and imports the Office 365 powershell session.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function refreshOffice365PowerShellSession
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function resets the Office 365 powershell sessions....' -toscreen
	}
	Process 
	{
        removeOffice365PowerShellSession  #Removes the Office 365 powershell session.
        createOffice365PowershellSession  #Creates the Office 365 powershell session.
        importOffice365PowershellSession  #Imports the Office 365 powershell session.
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All Office 365 powershell sessions have been refreshed.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "All Office 365 powershell sessions have not been refreshed." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}


<#
*******************************************************************************************************

Function establishOnPremisesCredentials

.DESCRIPTION

This function imports the on-premises credentials XML file.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function establishOnPremisesCredentials
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function imports the on premises secured credentials file....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onPremisesCredential = Import-Clixml -Path $script:onPremisesCredentialFile #create the credential variable for local Exchange.
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises credentials file was imported successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises credential file could not be imported - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function establishOffice365Credentials

.DESCRIPTION

This function imports the on-premises credentials XML file.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function establishOffice365Credentials
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function imports the Office 365 secured credentials file....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:office365Credential = Import-Clixml -Path $script:office365CredentialFile #create the credential variable for Office 365.
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The Office 365 credentials file was imported successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The Office 365 credential file could not be imported - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function createOnPremisesPowershellSession

.DESCRIPTION

This function creates the on premises powershell session to Exchange.
Recommendation to utilize a server Exchange 2016 or newer for avaialbility of all commands.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOnPremisesPowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the powershell session to on premises Exchange....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onPremisesPowerShellSession = New-PSSession -ConfigurationName $script:onPremisesPowershellConfiguration -ConnectionUri $script:onPremisesPowerShell -Authentication $script:onPremisesPowershellAuthentication -Credential $script:onPremisesCredential -AllowRedirection
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to on premises Exchange was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to on premises Exchange could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function createOffice365PowershellSession

.DESCRIPTION

This function creates the office 365 powershell session to Exchange Online.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOffice365PowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the powershell session to Office 365....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:office365PowerShellSession = New-PSSession -ConfigurationName $script:office365PowershellConfiguration -ConnectionUri $script:office365PowerShell -Authentication $script:office365PowershellAuthentication -Credential $script:office365Credential -AllowRedirection
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to Office 365 was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to Office 365 could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function createOnPremisesAADConnectPowershellSession

.DESCRIPTION

This function creates the on premises AAD Connect powershell session.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOnPremisesAADConnectPowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the powershell session to AAD Connect....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onPremisesAADConnectPowerShellSession = New-PSSession -ComputerName $script:onPremisesAADConnectServer -Credential $script:onPremisesCredential -Verbose
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to AAD Connect was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to AAD Connect could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function importOnPremisesPowershellSession

.DESCRIPTION

This function imports the powershell session to on premises Exchange.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function importOnPremisesPowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function imports the powershell session to on premises Exchange....' -toscreen
	}
	Process 
	{
		Try 
		{
            Import-PSSession $script:onPremisesPowerShellSession
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to on premises Exchange was imported successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to on premises Exchange could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function importOffice365PowershellSession

.DESCRIPTION

This function imports the powershell session to Office 365.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function importOffice365PowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function imports the powershell session to Office 365....' -toscreen
	}
	Process 
	{
		Try 
		{
            Import-PSSession $script:office365PowerShellSession -Prefix o365
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to Office 365 was imported successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to Office 365 could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function collectOnPremsiesDLConfiguration

.DESCRIPTION

This function collects the configuration of the on premises distribution list.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectOnPremsiesDLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the on premises distribution list configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlConfiguration = Get-DistributionGroup -identity $dlToConvert
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was collected successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function collectNewOffice365DLInformation

.DESCRIPTION

This function collects the configuration of the Office 365 distribution list.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectNewOffice365DLInformation
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the new office 365 distribution list configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfiguration = get-o365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was collected successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function collectNewOffice365DLInformation

.DESCRIPTION

This function collects the configuration of the Office 365 distribution group after creation.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectNewOffice365DLMemberInformation
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the new office 365 distribution list member configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfigurationMembership=get-o365DistributionGroupMember -identity $script:onpremisesdlConfiguration.primarySMTPAddress -resultsize unlimited
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was collected successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function collectOffice365DLConfiguation

.DESCRIPTION

This function collects the configuration of the Office 365 distribution list.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectOffice365DLConfiguation
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the Office 365 distribution list configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:office365DLConfiguration = Get-o365DistributionGroup -identity $dlToConvert
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The Office 365 distribution list information was collected successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The Office 365 distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function performOffice365SafetyCheck

.DESCRIPTION

This funciton reviews the settings of the cloud DL for dir sync status.
If dir sync is not true - assume DL specified exsits on premsies and in the cloud with same address.
This would indicated either direct cloud DL creation <or> migrated previously.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function performOffice365SafetyCheck
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function validates a cloud DLs saftey to migrate....' -toscreen
	}
	Process 
	{
			if ($script:office365DLConfiguration.IsDirSynced -eq $FALSE)
			{
				Write-LogError -LogPath $script:sLogFile -Message 'The DL requested for conversion was found in Office 365 and is not directory synced.  Cannot proceed.'
				Write-Error -Message "ERROR"
			}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The DL is safe to proeced for conversion - source of authority is on-premises.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The DL requested for conversion was found in Office 365 and is not directory synced.  Cannot proceed." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function backupOnPremisesDLConfiguration

.DESCRIPTION

This function writes the configuration of the on premies distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOnPremisesDLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the on prmeises distribution list configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlConfiguration | Export-CLIXML -Path $script:onPremisesXML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function backupNewOffice365DLConfiguration

.DESCRIPTION

This function writes the configuration of the new office 365 distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupNewOffice365DLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the new Office 365 distribution list configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfiguration | Export-CLIXML -Path $script:newOffice365XML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function backupNewOffice365DLConfigurationMembership

.DESCRIPTION

This function writes the configuration of the new Office 365 distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupNewOffice365DLConfigurationMembership
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the new Office 365 distribution list membership configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfigurationMembership | Export-CLIXML -Path $script:newOffice365MembershipXML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function backupOffice365DLConfiguration

.DESCRIPTION

This function writes the configuration of the Office 365 distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOffice365DLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the Office 365 distribution list configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:office365DLConfiguration | Export-CLIXML -Path $script:office365XML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The Office 365 distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The Office 365 distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function backupOnPremisesdlMembership

.DESCRIPTION

This function writes the configuration of the on premies distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOnPremisesdlMembership
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the on prmeises distribution list membership configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlconfigurationMembership | Export-CLIXML -Path $script:onPremsiesMembershipXML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list membership information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information membership could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function collectOnPremisesDLConfigurationMembership

.DESCRIPTION

This function collects the on-premises DL membership.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectOnPremisesDLConfigurationMembership
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collections the on premises DL membership....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlconfigurationMembership = get-distributionGroupMember -identity $dlToConvert -resultsize unlimited
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
            Write-LogInfo -LogPath $script:sLogFile -Message 'The DL membership was collected successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The DL membership could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function buildMembershipArray

.DESCRIPTION

This function builds an array of PSObjects representing DL memmbers or multi-valued attributes.

.PARAMETER OperationType

String $operationType - specifies the multi valued attribute we are working with.
String $arrayName - specifies the script variable array name to work with.
String $ignoreVariable - specifies if we should ignore errors found where multi0valued objects have users that are not represented in Office 365

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
Function buildMembershipArray
{
    Param ([string]$operationType,[string]$arrayName,[boolean]$ignoreVariable=$FALSE)

	Begin 
	{
        $functionArray = @()
		$functionOutput = @()
		$functionRecipient = $NULL
		$functionObject= $NULL
		$recipientObject = $NULL
		$userObject = $NULL

		Write-LogInfo -LogPath $script:sLogFile -Message 'This function builds an array of DL members or multivalued attributes....' -toscreen
	}
	Process 
	{
		#Based on the operation performed move a copy of the script array into a function array.
		#Function array allows us to reuse some code below.

        if ($arrayName -eq "onpremisesdlconfigurationMembershipArray")
        {
            $functionArray = $script:onpremisesdlconfigurationMembership
        }
        elseif ($arrayName -eq "onpremisesdlconfigurationManagedByArray")
        {
			$functionArray = $script:onpremisesdlConfiguration.ManagedBy
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationModeratedByArray")
		{
			$functionArray = $script:onpremisesdlConfiguration.ModeratedBy
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationGrantSendOnBehalfTOArray")
		{
			$functionArray = $script:onpremisesdlConfiguration.GrantSendOnBehalfTo
		}
		elseif ($arrayName -eq  "onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers")
		{
			$functionArray = $script:onpremisesdlConfiguration.AcceptMessagesOnlyFromSendersOrMembers
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationRejectMessagesFromSendersOrMembers")
		{
			$functionArray = $script:onpremisesdlConfiguration.RejectMessagesFromSendersOrMembers
		}
		elseif ($arrayName -eq "onPremsiesDLBypassModerationFromSendersOrMembers")
		{
			$functionArray = $script:onpremisesdlConfiguration.BypassModerationFromSendersOrMembers
		}

		#Based on the operation type passed act on the array of members.
		
        if ( $operationType -eq 'DLMembership' ) 
        {
			#Operation type is distribution list membership.

            foreach ( $member in $functionArray )
            {
				#If the recipient returned is not a user, contact, or groups it is mail recipient.
				#Gather information and place into PSObject.
				#Commit PSobject to output array.

                if ( ($member.recipientType.tostring() -ne "USER") -and ($member.recipientType.tostring() -ne "CONTACT") -and ($member.recipientType.tostring() -ne "GROUP") )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing mail enabled DL member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen
                    $functionRecipient = get-recipient -identity $member.PrimarySMTPAddress

                    #$recipientObject = New-Object System.Object

                    #$recipientObject | Add-Member -Type NoteProperty -Name "Alias" -Value $functionRecipient.Alias
                    #$recipientObject | Add-Member -Type NoteProperty -Name "Name" -Value $functionRecipient.Name
                    #$recipientObject | Add-Member -Type NoteProperty -Name "PrimarySMTPAddressOrUPN" -Value $functionRecipient.PrimarySMTPAddress
                    #$recipientObject | add-member -type NoteProperty -Name "RecipientType" -Value $functionRecipient.RecipientType
					#$recipientObject | Add-Member -Type NoteProperty -Name "RecipientOrUser" -Value "Recipient"

					$recipientObject = New-Object PSObject -Property @{
						Alias = $functionRecipient.Alias
						Name = $functionRecipient.Name
						PrimarySMTPAddressOrUPN = $functionRecipient.PrimarySMTPAddress
						RecipientType = $functionRecipient.RecipientType
						RecipientOrUser = "Recipient"
					}

                    $functionOutput += $recipientObject
				}
                elseif ( $member.recipientType.toString() -eq "USER" )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing non-mailenabled DL member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen
                    $functionUser = get-user -identity $member.name

                    #$userObject = New-Object System.Object

                    #$userObject | Add-Member -Type NoteProperty -Name "Alias" -Value $NULL
                    #$userObject | Add-Member -Type NoteProperty -Name "Name" -Value $functionUser.Name
                    #$userObject | Add-Member -Type NoteProperty -Name "PrimarySMTPAddressOrUPN" -Value $functionUser.UserprincipalName
                    #$userObject | add-member -type NoteProperty -Name "RecipientType" -Value "User"
					#$userObject | Add-Member -Type NoteProperty -Name "RecipientOrUser" -Value "User"

					$userObject = New-Object PSObject -Property @{
						Alias = $NULL
						Name = $functionRecipient.Name
						PrimarySMTPAddressOrUPN = $functionUser.UserprincipalName
						RecipientType = "User"
						RecipientOrUser = "User"
					}
					
                    $functionOutput += $userObject
                }
                elseif ($ignoreVariable -eq $FALSE )
                {
                    Write-LogError -LogPath $script:sLogFile -Message $member.name -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "A non-mail enabled or Office 365 object was found in the group." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "Script invoked without skipping invalid DL Member." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "The object must be removed or mail enabled." -ToScreen   
                    Write-LogError -LogPath $script:sLogFile -Message "EXITING." -ToScreen 
                    cleanupSessions
                    Stop-Log -LogPath $script:sLogFile -ToScreen  
                }
                else 
                {
                    #DO nothing - the user indicated that invalid recipients that cannot be migrated can be skipped.

                    Write-LogInfo -LogPath $script:sLogFile -Message "The following object was intentionally skipped - object type not replicated to Exchange Online" -ToScreen
                    Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen
                }
            }
		}
		
		#Operation type is managed by array.

        elseif ( $operationType -eq "ManagedBy"  )
        {
            foreach ( $member in $functionArray )
            {
				#Attempt to get a recipient.  If a recipient is returned the object is mail enabled.
				#This approach was taken becuase the attributes are stored in DN format on premises - and may not reflect a mail enabled object.
				#Managed by can be edited directly through AD and may contain non-mail enabled objects.

                if ( Get-Recipient -identity $member -errorAction SilentlyContinue )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing Managed By member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen
					$functionRecipient = get-recipient -identity $member

                    #$recipientObject = New-Object System.Object

                    #$recipientObject | Add-Member -Type NoteProperty -Name "Alias" -Value $functionRecipient.Alias
                    #$recipientObject | Add-Member -Type NoteProperty -Name "Name" -Value $functionRecipient.Name
                    #$recipientObject | Add-Member -Type NoteProperty -Name "PrimarySMTPAddressOrUPN" -Value $functionRecipient.PrimarySMTPAddress
                    #$recipientObject | add-member -type NoteProperty -Name "RecipientType" -Value $functionRecipient.RecipientType
					#$recipientObject | Add-Member -Type NoteProperty -Name "RecipientOrUser" -Value "Recipient"

					$recipientObject = New-Object PSObject -Property @{
						Alias = $functionRecipient.Alias
						Name = $functionRecipient.Name
						PrimarySMTPAddressOrUPN = $functionRecipient.PrimarySMTPAddress
						RecipientType = $functionRecipient.RecipientType
						RecipientOrUser = "Recipient"
					}

					
					$functionOutput += $recipientObject
                }
                elseif ($ignoreVariable -eq $FALSE )
                {
                    Write-LogError -LogPath $script:sLogFile -Message $member -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "A non-mail enabled or Office 365 object was found in ManagedBy." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "Script invoked without skipping invalid DL Member." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "The object must be removed or mail enabled." -ToScreen   
                    Write-LogError -LogPath $script:sLogFile -Message "EXITING." -ToScreen 
                    cleanupSessions
                    Stop-Log -LogPath $script:sLogFile -ToScreen  
                }
                else 
                {
                    #DO nothing - the user indicated that invalid recipients that cannot be migrated can be skipped.

                    Write-LogInfo -LogPath $script:sLogFile -Message "The following object was intentionally skipped - object type not replicated to Exchange Online" -ToScreen
                    Write-LogInfo -LogPath $script:sLogFile -Message $member -ToScreen
                }
            }
		}

		#Operation is a remaining multivalued attribute.

		elseif ( ( $operationType -eq "ModeratedBy" ) -or ( $operationType  -eq "GrantSendOnBehalfTo" ) -or ( $operationType -eq "AcceptMessagesOnlyFromSendersOrMembers") -or ($operationType -eq "RejectMessagesFromSendersOrMembers" ) -or ($operationType -eq "BypassModerationFromSendersOrMembers") )
		{
			foreach ( $member in $functionArray )
            {
				#Test to ensure that the object is a recipient.
				#In theory this is not required since Exchange commandlets will not let you set a non-mail enabled object on these properties...but...
				#For consistency sake we're doing the same thing...

                if ( Get-Recipient -identity $member -errorAction SilentlyContinue )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing ModeratedBy, GrantSendOnBehalfTo, AcceptMessagesOnlyFromSendersorMembers, RejectMessagesFromSendersOrMembers, or BypassModerationFromSendersOrMembers member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen
					$functionRecipient = get-recipient -identity $member

                    #$recipientObject = New-Object System.Object

                    #$recipientObject | Add-Member -Type NoteProperty -Name "Alias" -Value $functionRecipient.Alias
                    #$recipientObject | Add-Member -Type NoteProperty -Name "Name" -Value $functionRecipient.Name
                    #$recipientObject | Add-Member -Type NoteProperty -Name "PrimarySMTPAddressOrUPN" -Value $functionRecipient.PrimarySMTPAddress
                    #$recipientObject | add-member -type NoteProperty -Name "RecipientType" -Value $functionRecipient.RecipientType
					#$recipientObject | Add-Member -Type NoteProperty -Name "RecipientOrUser" -Value "Recipient"

					$recipientObject = New-Object PSObject -Property @{
						Alias = $functionRecipient.Alias
						Name = $functionRecipient.Name
						PrimarySMTPAddressOrUPN = $functionRecipient.PrimarySMTPAddress
						RecipientType = $functionRecipient.RecipientType
						RecipientOrUser = "Recipient"
					}

					
					$functionOutput += $recipientObject
                }
			}
		}

		#Based on the array name iterate and log the items found to process.

        if ( $arrayName -eq "onpremisesdlconfigurationMembershipArray" )
        {
			$script:onpremisesdlconfigurationMembershipArray = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationMembershipArray )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
        }
        elseif ($arrayName -eq "onpremisesdlconfigurationManagedByArray")
        {
			$script:onpremisesdlconfigurationManagedByArray = $functionOutput
			
			foreach ( $member in $script:onpremisesdlconfigurationManagedByArray )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationModeratedByArray")
		{
			$script:onpremisesdlconfigurationModeratedByArray = $functionOutput
			
			foreach ( $member in $script:onpremisesdlconfigurationModeratedByArray )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationGrantSendOnBehalfTOArray")
		{
			$script:onpremisesdlconfigurationGrantSendOnBehalfTOArray = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers")
		{
			$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationRejectMessagesFromSendersOrMembers")
		{
			$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers)
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
				Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen
			}
		}
		elseif ($arrayName -eq "onPremsiesDLBypassModerationFromSendersOrMembers")
		{
			$script:onPremsiesDLBypassModerationFromSendersOrMembers = $functionOutput

			foreach ( $member in $script:onPremsiesDLBypassModerationFromSendersOrMembers)
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
				Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen
			}
		}
	}
	End 
	{
		If ($?) 
		{
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message 'The array was built successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The array could not be built successfully - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function testOffice365Recipient

.DESCRIPTION

This function collects the on-premises DL membership.

.PARAMETER <Parameter_Name>

[string]$primarySMTPAddressOrUPN is the UPN or Primary SMTP address we are testing office 365 for.
[string]$userOrRecipient determines if we test with get-recipient (recipient) or get-user (user)

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function testOffice365Recipient
{
	Param ([string]$primarySMTPAddressOrUPN,[string]$UserorRecipient)

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function validates that all objects in the passed array exist in Office 365....' -toscreen
	}
	Process 
	{
		Try 
		{
			#Test to ensure the mail enabled or user object exists in Office 365.

			Write-LogInfo -LogPath $script:sLogFile -Message "Testing user in Office 365..." -ToScreen
			Write-LogInfo -LogPath $script:sLogFile -Message $primarySMTPAddressOrUPN -ToScreen
			Write-LogInfo -LogPath $script:sLogFile -Message $UserorRecipient
			
			if ( $UserorRecipient -eq "Recipient")
			{
				$functionTest=get-o365Recipient -identity $primarySMTPAddressOrUPN
			}
			elseif ($UserorRecipient -eq "User")
			{
				$functionTest=get-o365User -identity $primarySMTPAddressOrUPN
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The recipients were found in Office 365.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The recipients were not found in Office 365 - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function testOffice365GroupMigrated

.DESCRIPTION

This function tests to ensure that a sub group or permissions groups was migrated before the DL that references it.

If we did not do this migration of sub groups woudl be possible - membership and permissions would be lost.

.PARAMETER <Parameter_Name>

[string]$primarySMTPAddressOrUPN - group reference to test for.

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function testOffice365GroupMigrated
{
	Param ([string]$primarySMTPAddress)

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function tests to see if any sub groups or groups assigned permissions have been migrated....' -toscreen
	}
	Process 
	{
		#Test the dir sync flag to see if TRUE.  If TRUE the list was not migrated - abend.

		$functionTest = Get-o365DistributionGroup -identity $primarySMTPAddress
		Write-LogInfo -LogPath $script:sLogFile -Message 'Now testing group...' -ToScreen
		Write-LogInfo -LogPath $script:sLogFile -Message $functionTest.primarySMTPAddress -ToScreen
		Write-LogInfo -LogPath $script:sLogFile -Message $functionTest.IsDirSynced -ToScreen

		if ( $functionTest.IsDirSynced -eq $TRUE)
		{
			Write-LogError -LogPath $script:sLogFile -Message 'A distribution list was found as a sub-member or on a multi-valued attribute.' -ToScreen
			Write-LogError -LogPath $script:sLogFile -Message 'The distribution list has not been migrated to Office 365 (DirSync Flag is TRUE)' -ToScreen
			Write-LogError -LogPath $script:sLogFile -Message 'All sub lists or lists with permissions must be migrated before proceeding.' -ToScreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The recipients were found in Office 365.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The recipients were not found in Office 365 - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function moveGroupToOU

.DESCRIPTION

This function moves the group to an OU that is not synchonized by AAD Connect

This requires prior configuration of this setting in AAD Connect.

This requires the OU exist prior to script execution.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function moveGroupToOU
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function moves group to the non-sync OU' -toscreen
	}
	Process 
	{
		Try 
		{
			Move-ADObject -Identity $script:onpremisesdlConfiguration.distinguishedName -TargetPath $script:groupOrganizationalUnit
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The group has been moved successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Group Could Not Be Moved...exiting" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function getADDomainControllers

.DESCRIPTION

This function gathers all of the Active Directory Domain Controllers.

Note:	If the administrator desires that a smaller subset of domain controllers be selected or only one the filter should be updated.

Note:	May require enterprise admin rights.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function getADDomainControllers
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'Gets active directory domain controllers...' -toscreen
	}
	Process 
	{
		Try 
		{
			$script:adDomainControllers = Get-ADDomainController -Filter *
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Succesfully obtained domain controllers.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Domain controllers could not be obtained..." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function getADDomain

.DESCRIPTION

This function gathers the AD Domain information and assumes the workstation is in the same domain.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function getADDomain
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'Gets active directory domain...' -toscreen
	}
	Process 
	{
		Try 
		{
			$script:adDomain = Get-ADDomain
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Succesfully obtained domain.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Domain could not be obtained..." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function replicateDomainControllers

.DESCRIPTION

This function replicates domain controllers in the active directory.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function replicateDomainControllers

{
	Param ([string]$domainController,[string]$distinguishedName)

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'Replicates the specified domain controller...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $domainControllerName -toscreen
	}
	Process 
	{
		Try 
		{
			repadmin /syncall $domainController $distinguishedName
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Successfully replicated the domain controller.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "The domain controller could not be replicated - this does not cause the script to abend..." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
		}
	}
}


<#
*******************************************************************************************************

Function invokeADConnect

.DESCRIPTION

This function invokes a delta sync through remote powershell.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function invokeADConnect
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function triggers the ad connect process to sync changes...' -toscreen
	}
	Process 
	{
		Try 
		{
			Invoke-Command -Session $script:onPremisesAADConnectPowerShellSession -ScriptBlock {Import-Module -Name 'AdSync'}
			Invoke-Command -Session $script:onPremisesAADConnectPowerShellSession	-ScriptBlock {start-adsyncsynccycle -policyType Delta}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The AD Connect instance has been successfully initiated.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "ADConnect sync could not be triggered...this does not cause the script to abend." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
		}
	}
}

<#
*******************************************************************************************************

Function createOffice365DistributionList

.DESCRIPTION

This function creates the new cloud DL with the minimum attributes. 

Detailed attributes will be handeled later.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOffice365DistributionList
{
	Param ()
	
	Begin 
	{
		$functionGroupType = $NULL #Utilized to establish group type distribution or group type security.

		Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the cloud DL with the minimum settings...' -toscreen

		if ( $script:onpremisesdlConfiguration.GroupType -eq "Universal, SecurityEnabled" )
		{
			$functionGroupType="Security"
		}
		elseif (  $script:onpremisesdlConfiguration.GroupType -eq "Universal" )
		{
			$functionGroupType="Distribution"
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "This script only supports universal distribution and universal security group conversions." -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
		$script:onpremisesdlConfiguration.GroupType
	}
	Process 
	{
		Try 
		{
			new-o365DistributionGroup -name $script:onpremisesdlConfiguration.Name -alias $script:onpremisesdlConfiguration.Alias -primarySMTPAddress $script:onpremisesdlConfiguration.PrimarySmtpAddress -type $functionGroupType
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution list created successfully in Exchange Online / Office 365.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Distribution list could not be created in Exchange Online / Office 365...exiting" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
		}
	}
}

<#
*******************************************************************************************************

Function setOffice365DistributionListSettings

.DESCRIPTION

This function sets the single attribute settings on the new Cloud DL based on the previous on-premises DL.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function setOffice365DistributionListSettings
{
	Param ()

	Begin 
	{
		$functionX500 = $NULL
		$functionEmailAddresses = $NULL
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function updates the cloud DL settings to match on premise...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This does not update the multivalued attributes...' -ToScreen

		#Build the X500 address for the new email address based off on premises values.

		$functionX500 = "X500:"+$script:onpremisesdlConfiguration.legacyExchangeDN
		$functionEmailAddresses=$script:onpremisesdlconfiguration.emailAddresses
		$functionEmailAddresses+=$functionX500
	}
	Process 
	{
		Try 
		{
			Set-O365DistributionGroup -Identity $script:onpremisesdlConfiguration.primarySMTPAddress -BypassNestedModerationEnabled $script:onpremisesdlconfiguration.BypassNestedModerationEnabled -MemberJoinRestriction $script:onpremisesdlconfiguration.MemberJoinRestriction -MemberDepartRestriction $script:onpremisesdlconfiguration.MemberDepartRestriction -ReportToManagerEnabled $script:onpremisesdlconfiguration.ReportToManagerEnabled -ReportToOriginatorEnabled $script:onpremisesdlconfiguration.ReportToOriginatorEnabled -SendOofMessageToOriginatorEnabled $script:onpremisesdlconfiguration.SendOofMessageToOriginatorEnabled -Alias $script:onpremisesdlconfiguration.Alias -CustomAttribute1 $script:onpremisesdlconfiguration.CustomAttribute1 -CustomAttribute10 $script:onpremisesdlconfiguration.CustomAttribute10 -CustomAttribute11 $script:onpremisesdlconfiguration.CustomAttribute11 -CustomAttribute12 $script:onpremisesdlconfiguration.CustomAttribute12 -CustomAttribute13 $script:onpremisesdlconfiguration.CustomAttribute13 -CustomAttribute14 $script:onpremisesdlconfiguration.CustomAttribute14 -CustomAttribute15 $script:onpremisesdlconfiguration.CustomAttribute15 -CustomAttribute2 $script:onpremisesdlconfiguration.CustomAttribute2 -CustomAttribute3 $script:onpremisesdlconfiguration.CustomAttribute3 -CustomAttribute4 $script:onpremisesdlconfiguration.CustomAttribute4 -CustomAttribute5 $script:onpremisesdlconfiguration.CustomAttribute5 -CustomAttribute6 $script:onpremisesdlconfiguration.CustomAttribute6 -CustomAttribute7 $script:onpremisesdlconfiguration.CustomAttribute7 -CustomAttribute8 $script:onpremisesdlconfiguration.CustomAttribute8 -CustomAttribute9 $script:onpremisesdlconfiguration.CustomAttribute9 -ExtensionCustomAttribute1 $script:onpremisesdlconfiguration.ExtensionCustomAttribute1 -ExtensionCustomAttribute2 $script:onpremisesdlconfiguration.ExtensionCustomAttribute2 -ExtensionCustomAttribute3 $script:onpremisesdlconfiguration.ExtensionCustomAttribute3 -ExtensionCustomAttribute4 $script:onpremisesdlconfiguration.ExtensionCustomAttribute4 -ExtensionCustomAttribute5 $script:onpremisesdlconfiguration.ExtensionCustomAttribute5 -DisplayName $script:onpremisesdlconfiguration.DisplayName -EmailAddresses $functionEmailAddresses -HiddenFromAddressListsEnabled $script:onpremisesdlconfiguration.HiddenFromAddressListsEnabled -ModerationEnabled $script:onpremisesdlconfiguration.ModerationEnabled -RequireSenderAuthenticationEnabled $script:onpremisesdlconfiguration.RequireSenderAuthenticationEnabled -SimpleDisplayName $script:onpremisesdlconfiguration.SimpleDisplayName -SendModerationNotifications $script:onpremisesdlconfiguration.SendModerationNotifications -WindowsEmailAddress $script:onpremisesdlconfiguration.WindowsEmailAddress -MailTipTranslations $script:onpremisesdlconfiguration.MailTipTranslations -Name $script:onpremisesdlconfiguration.Name
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group properties updated successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Cannot update properties of distribution group." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
		}
	}
}

<#
*******************************************************************************************************

Function setOffice365DistributionListSettings

.DESCRIPTION

This function sets the multi valued attribute settings on the new Cloud DL based on the previous on-premises DL.

.PARAMETER 

[string]$operationType is the type of operation - for example moderateydBy or managedBy.
[string]$primarySMTPAddressOrUPN is the primary SMTP address of a recipient or UPN of the user object that we are operating on.

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function setOffice365DistributionlistMultivaluedAttributes
{
	Param ([string]$operationType,[string]$primarySMTPAddressOrUPN)

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function sets the multi-valued attributes' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $operationType -ToScreen
		Write-LogInfo -LogPath $script:sLogFile -Message $primarySMTPAddressOrUPN
	}
	Process 
	{
		#Based on the operation type specified utilize the appropriate array and iterate through each member adding to the attribute in the service.

		if ( $operationType -eq "DLMembership")
		{
			Try
			{
				add-o365DistributionGroupMember -identity $script:onpremisesdlConfiguration.primarySMTPAddress -member $PrimarySMTPAddressOrUPN
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				Break
			}
		}
		elseif ( $operationType -eq "ManagedBy")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -managedBy @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				Break
			}
		}
		elseif ( $operationType -eq "ModeratedBy")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -ModeratedBy @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				Break
			}
		}
		elseif ( $operationType -eq "GrantSendOnBehalfTo")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -GrantSendOnBehalfTo @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				Break
			}
		}
		elseif ( $operationType -eq "AcceptMessagesOnlyFromSendersOrMembers")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -AcceptMessagesOnlyFromSendersOrMembers @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				Break
			}
		}
		elseif ( $operationType -eq "RejectMessagesFromSendersOrMembers")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -RejectMessagesFromSendersOrMembers @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				Break
			}
		}
		elseif ( $operationType -eq "BypassModerationFromSendersOrMembers")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -BypassModerationFromSendersOrMembers @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				Break
			}
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The mutilvalued attribute was updated successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $operationType -ToScreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The mutilvalued attribute could not be updated successfully - exiting." -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $operationType -ToScreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Create log file for operations within this script.

Start-Log -LogPath $script:sLogPath -LogName $script:sLogName -ScriptVersion $script:sScriptVersion -ToScreen

establishOnPremisesCredentials  #Function call to import and populate on premises credentials.

establishOffice365Credentials  #Function call to import and populate Office 365 credentials.

createOnPremisesPowershellSession  #Creates the on premises powershell session to Exchange.

createOffice365PowershellSession  #Creates the Office 365 powershell session.

createOnPremisesAADConnectPowershellSession  #Create the on premises AAD Connect powershell session.

importOnPremisesPowershellSession  #Function call to import the on premises powershell session.

importOffice365PowershellSession  #Function call to import the Office 365 powershell session.

collectOnPremsiesDLConfiguration  #Function call to gather the on premises DL information.

collectOffice365DLConfiguation  #Function call to gather the Office 365 DL information.

performOffice365SafetyCheck #Checks to see if the distribution list provided has already been migrated.

backuponpremisesdlConfiguration  #Function call to write the on premises DL information to XML.

backupOffice365DLConfiguration  #Function call to write the Office 365 DL information to XML.

collectonpremisesdlconfigurationMembership  #Function collects the membership of the on premise DL.

backupOnPremisesdlMembership #Writes the on premises DL membership to XML for protection and auditing.

removeOffice365PowerShellSession  #Remove the office 365 powershell session.  It is not needed at this time and will be recreated after long running operations.

#Begin processing the on premises membership and multi-valued array.
#We take the array and call build array to create a normalized list of SMTP addresses for later operations.
#This is required since many of the multi-valued attributes store their references as disinguished names - which will not translate to Office 365.
#We then test to ensure that the target recipient exists in office 365 as a mail enabled object or a user.
#We then test to see if the recipient is a group has the group already been migrated.  Groups that are members must be migrated first.
#Several set functions have a counter routine set to the defined value.  When we hit this value - we refresh powershell to office 365.

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a DL membership array." -ToScreen

if ( $script:onpremisesdlconfigurationMembership -ne $NULL )
{
    buildMembershipArray ( "DLMembership" ) ( "onpremisesdlconfigurationMembershipArray") ($ignoreInvalidDLMember) #This function builds an array of members for the DL <or> multivalued attributes.
    
    refreshOffice365PowerShellSession #Refreshing the sessio here since building the membership array can take a while depending on array size.

    $script:forCounter=0
	
	foreach ($member in $script:onpremisesdlconfigurationMembershipArray)
	{
        if ($script:forCounter -gt $script:refreshCounter)
        {
            refreshOffice365PowerShellSession
            $script:forCounter = 0
        }

		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
        }
        
        $script:forCounter+=1
    }
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a ManagedBy array." -ToScreen

if ( $script:onpremisesdlConfiguration.ManagedBy -ne $NULL )
{
    buildMembershipArray ( "ManagedBy" ) ( "onpremisesdlconfigurationManagedByArray" ) ( $ignoreInvalidManagedByMember )
    
    refreshOffice365PowerShellSession

	foreach ($member in $script:onpremisesdlconfigurationManagedByArray)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}
	}
}

Write-LogInfo -logpath $script:sLogFile -Message "Begin processing a ModeratedBy array." -ToScreen

if ( $script:onpremisesdlConfiguration.ModeratedBy -ne $NULL )
{
    buildmembershipArray ( "ModeratedBy" ) ( "onpremisesdlconfigurationModeratedByArray" ) 
    
    refreshOffice365PowerShellSession

	foreach ($member in $script:onpremisesdlconfigurationModeratedByArray)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)
		
		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a GrantSendOnBehalfTo array" -ToScreen

if ( $script:onpremisesdlConfiguration.GrantSendOnBehalfTo -ne $NULL )
{
    buildmembershipArray ( "GrantSendOnBehalfTo" ) ( "onpremisesdlconfigurationGrantSendOnBehalfTOArray" )
    
    refreshOffice365PowerShellSession

	foreach ($member in $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a AcceptMessagesOnlyFromSendersOrMembers array" -ToScreen

if ( $script:onpremisesdlConfiguration.AcceptMessagesOnlyFromSendersOrMembers -ne $NULL )
{
    buildMembershipArray ( "AcceptMessagesOnlyFromSendersOrMembers" ) ( "onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers" )
    
    refreshOffice365PowerShellSession

	foreach ($member in $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing RejectMessagesFromSendersOrMembers array" -ToScreen

if ( $script:onpremisesdlConfiguration.RejectMessagesFromSendersOrMembers -ne $NULL)
{
    buildMembershipArray ( "RejectMessagesFromSendersOrMembers" ) ( "onpremisesdlconfigurationRejectMessagesFromSendersOrMembers" )
    
    refreshOffice365PowerShellSession

	foreach ($member in $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing BypassModerationFromSendersOrMembers array" -ToScreen

if ( $script:onpremisesdlConfiguration.BypassModerationFromSendersOrMembers -ne $NULL)
{
    buildMembershipArray ( "BypassModerationFromSendersOrMembers") ( "onPremsiesDLBypassModerationFromSendersOrMembers" )
    
    refreshOffice365PowerShellSession

	foreach ($member in $script:onPremsiesDLBypassModerationFromSendersOrMembers)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}
	}
}

moveGroupToOU  #Move the group to a non-sync OU to preserve it.

getADDomainControllers  #Gather active directory domain controllers.

getADDomain #Gather the ad domain name.

#Replicate each domain controller in the domain.
#Administrators may choose to refresh this to a single domain controller or set of domain controllers by calling the function with static values and more than once.

foreach ( $DC in $script:adDomainControllers )
{
	replicateDomainControllers ( $dc.HostName ) ( $script:adDomain.DistinguishedName )
}

#Start countdown for the period of time specified by the variable for post domain controller replication.

Start-PSCountdown -Minutes 1 -Title "Waiting for domain controller replication" -Message "Waiting for domain controller replication"

Write-LogInfo -LogPath $script:sLogFile -Message "Invoking AADConnect Delta Sync Remotely" -ToScreen

#Invoke ad sync.
#If a sync is already in progress this will return an retryable error condition.
#Set the retry variable to true and loop back through.
#The script enforces that at least one delta sync was triggered by it to ensure that it was run post DL movement to non-sync OU.

do
{
	if ( $script:aadconnectRetryRequired -eq $TRUE )
	{
		Start-PSCountdown -Minutes $script:adDomainReplicationTime -Title "Waiting for previous sync to finishn <or> allowing time for invoked sync to run" -Message "Waiting for previous sync to to finishn <or> allowing time for invoked sync to run"
	}
	invokeADConnect	
	$script:aadconnectRetryRequired = $TRUE	
} while ( $error -ne $NULL )

refreshOffice365PowerShellSession

$error.clear()

#Test to wait until the DL is removed from the service.
#This code continues to try to find the original DL in office 365.
#If the DL is found the retry variable will tirgger us to loop back around and try again.
#When the error condition is encountered the DL is no longer there - good - we can move on.

Write-LogInfo -LogPath $script:sLogFile -Message "Wating for original DL deletion from Office 365" -ToScreen

do
{
	if ( $script:dlDeletionRetryRequired -eq $TRUE)
	{
		Write-LogInfo -LogPath $script:sLogFile -Message "Wating for original DL deletion from Office 365" -ToScreen
		Start-PSCountdown -Minutes $script:dlDelectionTime -Title "Waiting for DL deletion to process in Office 365" -Message "Waiting for DL deletion to process in Office 365"
		$error.clear()
	}
	$script:dlDeletionRetryRequired = $TRUE
	$scriptTest=get-o365Recipient -identity $script:onpremisesdlConfiguration.primarySMTPAddress
	
} until ( $error -ne $NULL )

refreshOffice365PowerShellSession

$error.clear()

createOffice365DistributionList

#Set the settings of the distrbution list.
#For multivalued attributes that are not NULL set the individual multivalued attribute.
#For multivalued attributes trigger the appropriate add function with the operation name and the recipient to add.

setOffice365DistributionListSettings

if ( $script:onpremisesdlconfigurationMembershipArray -ne $NULL)
{
    $script:forCounter=0

	foreach ($member in $script:onpremisesdlconfigurationMembershipArray)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing DL Membership member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
        if ($script:forCounter -gt 1000)
        {
            refreshOffice365PowerShellSession
            $script:forCounter=0
        }
        setOffice365DistributionlistMultivaluedAttributes ( "DLMembership" ) ( $member.PrimarySMTPAddressOrUPN )
        $script:forCounter+=1
	}
}

if ( $script:onpremisesdlconfigurationManagedByArray -ne $NULL)
{
	foreach ($member in $script:onpremisesdlconfigurationManagedByArray )
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Bypass Managed By member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "ManagedBy" ) ( $member.PrimarySMTPAddressOrUPN )
	}
}

if ( $script:onpremisesdlconfigurationModeratedByArray -ne $NULL)
{
	foreach ($member in $script:onpremisesdlconfigurationModeratedByArray)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Moderated By member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "ModeratedBy" ) ( $member.PrimarySMTPAddressOrUPN  )
	}
}

if ( $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray -ne $NULL )
{
	foreach ($member in $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Grant Send On Behalf To Array member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "GrantSendOnBehalfTo" ) ( $member.PrimarySMTPAddressOrUPN  )
	}
}

if ( $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers -ne $NULL )
{
	foreach ($member in $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Accept Messages Only From Senders Or Members member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "AcceptMessagesOnlyFromSendersOrMembers" ) ( $member.PrimarySMTPAddressOrUPN  )
	}
}

if ( $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers -ne $null)
{
	foreach ($member in $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Reject Messages From Senders Or Members member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "RejectMessagesFromSendersOrMembers" ) ( $member.PrimarySMTPAddressOrUPN  )
	}
}

if ( $script:onPremsiesDLBypassModerationFromSendersOrMembers -ne $NULL )
{
	foreach ($member in $script:onPremsiesDLBypassModerationFromSendersOrMembers )
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Bypass Moderation From Senders Or Members member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "BypassModerationFromSendersOrMembers" ) ( $member.PrimarySMTPAddressOrUPN  )
	}
}

refreshOffice365PowerShellSession

collectNewOffice365DLInformation  #Collect the new office 365 dl configuration.

collectNewOffice365DLMemberInformation #Collect the new office 365 dl membership information.

backupNewOffice365DLConfiguration  #Backup the office 365 DL configuration.

#Assuming there were members in the DL (why would you migrate otherwise but we'll check anyway...write those members to XML.)

if ($script:newOffice365DLConfigurationMembership -ne $NULL)
{
	backupNewOffice365DLConfigurationMembership
}

cleanupSessions  #Clean up - were outta here.