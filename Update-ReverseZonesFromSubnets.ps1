#Requires -Version 4.0
#Requires -Module ActiveDirectory
#Requires -Module DnsServer


<#
.SYNOPSIS
	Script to get subnets from Sites & Services, see if a matching AD DNS Reverse Lookup 
	Zone exists, and if not, create the reverse zone.
.DESCRIPTION
	This script was created for a client who granted permission to share with the 
	community.

	This script reads the list of Subnets in AD Sites & Services, checks if a matching DNS 
	Reverse Lookup Zone exists, if the zone does not exist, attempts to create it.
	
	The Reverse Zones created are created with a Replication Scope of "Forest" and Dynamic 
	Update set to "Secure". By default, Aging and Scavenging is enabled with the default 
	7 days No-refresh and Refresh intervals.

	The script requires at least PowerShell version 4 but runs best in version 5.

	This script requires Domain Admin rights and an elevated PowerShell session.

	Creates an output file named UpdateReverseZonesFromSubnetsScriptResults_YYYYMMDDHHSS.txt.

	You do NOT have to run this script on a domain controller. This script was developed 
	and run from a Windows 10 VM.

	To run the script from a workstation, RSAT is required.

	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520

.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output log file. 
.PARAMETER From
	Specifies the username for the From email address.
	
	Note: To use unauthenticated email, the From must be Anonymous@emaildomain.tld.
	
	If SmtpServer or To are used, this is a required parameter.
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER SmtpPort
	Specifies the SMTP port for the SmtpServer. 
	The default is 25.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report(s). 
	
	If From or To are used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	
	If SmtpServer or From are used, this is a required parameter.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1

	The script will create AD DNS Reverse Lookup zones for every AD Sites & Services subnets
	that do not exist.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 -Whatif

	The script will show what it would have done to create AD DNS Reverse Lookup zones for 
	every AD Sites & Services subnets that do not exist.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 -Confirm

	The script will ask to create AD DNS Reverse Lookup zones for every AD Sites & Services 
	subnets that do not exist.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 -Folder \\FileServer\ShareName
	
	Output log file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 -Dev -ScriptInfo -Log
	
	Creates a text file named UpdateReverseZonesFromSubnetsScriptErrors_yyyy-MM-dd_HHmm.txt 
	that contains up to the last 250 errors reported by the script.
	
	Creates a text file named UpdateReverseZonesFromSubnetsScriptInfo_yyyy-MM-dd_HHmm.txt 
	that contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	UpdateReverseZonesFromSubnetsScriptTranscript_yyyy-MM-dd_HHmm.txt.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 
	-SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld 
	-To ITGroup@domain.tld	

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and will not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 
	-SmtpServer mailrelay.domain.tld
	-From Anonymous@domain.tld 
	-To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script will use the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script will use the default SMTP port 25 and will not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script will generate an anonymous secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 
	-SmtpServer labaddomain-com.mail.protection.outlook.com
	-UseSSL
	-From SomeEmailAddress@labaddomain.com 
	-To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script will use the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script will use the default SMTP port 25 and will use SSL.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 
	-SmtpServer smtp.office365.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Update-ReverseZonesFromSubnets.ps1 
	-SmtpServer smtp.gmail.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script will use the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a plain text document.
.NOTES
	NAME: Update-ReverseZonesFromSubnets.ps1
	VERSION: 1.10
	AUTHOR: Carl Webster
	LASTEDIT: April 29, 2020
#>


[CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = "Medium", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[string]$User=$env:username,
	
	[parameter(Mandatory=$False)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False
	
	)

	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on March 23, 2020

#Version 1.0 released to the community on 20-Apr-2020
#This script was created for a customer whose CIO requested this script be given to the community.
#
#Version 1.10 29-Apr-2020
#	Cleaned up some code and typos driving my OCD up the wall
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#

Set-StrictMode -Version Latest

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$ErrorActionPreference = 'SilentlyContinue'
$ConfirmPreference = "High"

#needs to run from an elevated PowerShell session
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
{
	Write-Verbose "$(Get-Date): This is an elevated PowerShell session"
}
Else
{
	Write-Error "
	`n`n
	`t`t
	This is NOT an elevated PowerShell session.
	`n`n
	`t`t
	Script will exit.
	`n`n
	"
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "
			`n`n
			`tFolder $Folder is a file, not a folder.
			`n`n
			`tScript cannot continue.
			`n`n"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
		`tFolder $Folder does not exist.
		`n`n
		`tScript cannot continue.
		`n`n"
		Exit
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$($Script:pwdpath)\UpdateReverseZonesFromSubnetsScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($Script:pwdpath)\UpdateReverseZonesFromSubnetsScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}

#region email function
Function SendEmail
{
	Param([string]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date): Email was not sent:"
					Write-Warning "$(Get-Date): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date): Email was not sent:"
				Write-Warning "$(Get-Date): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region script start function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

	#find PDCe DC
	$Script:ServerName = (Get-ADDomain -EA 0).PDCEmulator
	
	If($null -eq $Script:ServerName -or $Script:ServerName -eq "")
	{
		Write-Error "
		`n`n
		`t`t
		Unable to find the PDCe Domain Controller.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
	}
	
	Write-Verbose "$(Get-Date): Will use Domain Controller $Script:ServerName"
	
	$Script:Title = "Update Reverse Zones from Subnets_$(Get-Date -f yyyy-MM-dd_HHmm)"
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Dev                  : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile         : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Domain Controller    : $($Script:ServerName)"
	Write-Verbose "$(Get-Date): Folder               : $($Folder)"
	Write-Verbose "$(Get-Date): From                 : $($From)"
	Write-Verbose "$(Get-Date): Log                  : $($Log)"
	Write-Verbose "$(Get-Date): ScriptInfo           : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port            : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server          : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title                : $($Script:Title)"
	Write-Verbose "$(Get-Date): To                   : $($To)"
	Write-Verbose "$(Get-Date): Use SSL              : $($UseSSL)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected          : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version         : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture            : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture          : $($PSUICulture)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start         : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

#endregion

#region process subnets
Function ProcessSubnets
{
	$Results = Get-ADReplicationSubnet -filter * -Server $Script:ServerName -EA 0

	If(!$?)
	{
		Write-Warning "Unable to retrieve AD Sites & Service Subnets"
	}
	ElseIf($? -and $Null -eq $Results)
	{
		Write-Warning "There were no AD Sites & Service Subnets found"
	}
	Else
	{
		$Subnets = @(Get-ADReplicationSubnet -filter * -Server $Script:ServerName -EA 0).name
		
		$DNSResults = New-Object System.Collections.ArrayList
	
        If($WhatIfPreference)
        {
		    $null = $DNSResults.Add("WhatIf Enabled for this script run. No changes made.")
		    $null = $DNSResults.Add("")
        }
	
		$cnt = 1
		If($Subnets -is [array])
		{
			$cnt = $Subnets.Count
		}
		
		$Subnets = $Subnets | Sort-Object
		
		Write-Verbose "$(Get-Date): Successfully retrieved $cnt AD Sites & Service Subnets"
		Write-Verbose "$(Get-Date): "
		
		ForEach($Subnet in $Subnets)
		{
			Write-Verbose "$(Get-Date): Processing subnet $Subnet"
			
			$null = $DNSResults.Add("Processing subnet $Subnet.")

			$SubnetArray = $Subnet.split("./")
			$SubnetMask = [int]$SubnetArray[($SubnetArray.count-1)]
			
			If($SubnetMask -le 8)
			{
				$RevZone = "$($SubnetArray[0]).in-addr.arpa"
			}
			ElseIf($SubnetMask -le 16)
			{
				$RevZone = "$($SubnetArray[1]).$($SubnetArray[0]).in-addr.arpa"
			}
			ElseIf($SubnetMask -le 24)
			{
				$RevZone = "$($SubnetArray[2]).$($SubnetArray[1]).$($SubnetArray[0]).in-addr.arpa"
			}
			Else
			{
				$RevZone = "$($SubnetArray[3]).$($SubnetArray[2]).$($SubnetArray[1]).$($SubnetArray[0]).in-addr.arpa"
			}
			
			Write-Verbose "$(Get-Date): `tSearch for Reverse Lookup Zone $RevZone"
			$null = $DNSResults.Add("`tSearch for Reverse Lookup Zone $RevZone.")
			
			$Results = Get-DnsServerZone -name $RevZone -ComputerName $Script:ServerName -EA 0
			
			If($? -and $null -ne $results)
			{
				Write-Verbose "$(Get-Date): `t`t$RevZone was found, nothing to do."
				$null = $DNSResults.Add("`t`t$RevZone was found, nothing to do.")
			}
			Else
			{
				Write-Verbose "$(Get-Date): `t`t$RevZone was not found. Attempt to create it."
				$null = $DNSResults.Add("`t`t$RevZone was not found. Attempt to create it.")
				
				If($PSCmdlet.ShouldProcess($Subnet,'Create Reverse Lookup Zone'))
				{
					Try
					{
						$Results = Add-DnsServerPrimaryZone -NetworkID $Subnet -ReplicationScope "Forest" -DynamicUpdate "Secure" -ComputerName $Script:ServerName -PassThru -EA 0 *>$Null

						Write-Verbose "$(Get-Date): `t`t`tSuccessfully created Reverse Lookup Zone $RevZone for subnet $Subnet"
						$null = $DNSResults.Add("`t`t`tSuccessfully created Reverse Lookup Zone $RevZone for subnet $Subnet.")
					}
					
					Catch
					{
						Write-Warning "`t`t`tFailed to create Reverse Lookup Zone $RevZone for subnet $Subnet"
						$null = $DNSResults.Add("`t`t`tFailed to create Reverse Lookup Zone $RevZone for subnet $Subnet.")
					}
				}
				Else
				{
					If($WhatIfPreference)
					{
						$null = $DNSResults.Add("`t`t`tWhatIf used so no creating Reverse Lookup Zone $RevZone for subnet $Subnet.")
					}
					Else
					{
						#negative response to confirm
						$null = $DNSResults.Add("`t`t`tAnswered NO to Confirm prompt for creating Reverse Lookup Zone $RevZone for subnet $Subnet.")
					}
				}
			}
		}
		
		$File = Join-Path $Script:pwdpath "UpdateReverseZonesFromSubnetsScriptResults_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		$DNSResults | Out-File -FilePath $File -Force -WhatIf:$False *>$Null
		Write-Verbose "$(Get-Date): "
		Write-Verbose "$(Get-Date): $File is ready for use"
		Write-Verbose "$(Get-Date): "

		#email output file if requested
		If(![System.String]::IsNullOrEmpty( $SmtpServer ))
		{
			SendEmail $File
		}
	}
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$($Script:pwdpath)\UpdateReverseZonesFromSubnetScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                  : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile         : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Domain Controller    : $($Script:ServerName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Folder               : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From                 : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                  : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info          : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port            : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server          : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title                : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                   : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL              : $($UseSSL)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected          : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version         : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture            : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture          : $($PSUICulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start         : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time         : $($Str)" 4>$Null
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date): Transcript/log stop failed"
			}
		}
	}
	
	#cleanup obj variables
	$Script:Output = $Null
	$runtime = $Null
	$Str = $Null
}
#endregion

ProcessScriptSetup

ShowScriptOptions

ProcessSubnets

ProcessScriptEnd
