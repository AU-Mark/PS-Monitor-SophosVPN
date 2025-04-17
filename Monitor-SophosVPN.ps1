<#
.SYNOPSIS
Check the Sophos VPN status and trigger a toast notification if the VPN requires reauthentication

.DESCRIPTION
Uses output from the SCCLI console application to list all VPN connections. Parses the output and
loops indefinitely polling VPN status every 10 seconds. When the VPN is disconnected, if the cause
is that it requres reauthentication then a toast notification is triggered.

.INPUTS
None

.OUTPUTS
"$env:TEMP\ToastPicture.png" and "$env:TEMP\ToastIcon.png" for toast notification images

.NOTES
Version:        1.0
Author:         Mark Newton
Creation Date:  02/14/2025
Purpose/Change: Initial script development
#>

##############################################################################################################
#                                                Globals                                                     #
##############################################################################################################

if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
 { $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
 else
 { $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
     if (!$ScriptPath){ $ScriptPath = "." } }

##############################################################################################################
#                                             Base 64 Images                                                 #
##############################################################################################################

# Toast header in Base 64 format
$Picture_Base64 = "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"

# Toast Icon in Base64 format
$Icon_Base64 = "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"

##############################################################################################################
#                                                Functions                                                   #
##############################################################################################################
Function Write-Log {
    Param(
        [Parameter(Mandatory=$True)][String]$LogMsg,
        [Parameter(Mandatory=$False)][String]$LogLvl="INFO",
        [Parameter(Mandatory=$False)][String]$LogName="Debug.log"
        
    )

    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"

    If (!(Test-Path -Path "$ScriptPath\$LogName")) {
        Write-Output "[$TimeStamp][INFO] Logging started" | Out-File -FilePath "$ScriptPath\$LogName" -Append
    }

    Write-Output "[$TimeStamp][$LogLvl] $LogMsg" | Out-File -FilePath "$ScriptPath\$LogName" -Append
}

function Ensure-WatchSophosVPNTask {
    $taskName = "Watch-SophosVPN"
    $taskPath = "C:\Scripts\Watch-SophosVPN.exe"

    # Check if the task exists
    $taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if (-not $taskExists) {
        # Create the task if it does not exist
        $action = New-ScheduledTaskAction -Execute $taskPath
        $trigger = New-ScheduledTaskTrigger -AtLogon
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Description "Watch Sophos VPN Task"
        Write-Output "Task '$taskName' created successfully."
    } else {
        Write-Output "Task '$taskName' already exists."
    }

    # Check if the task is running
    $taskInfo = Get-ScheduledTaskInfo -TaskName $taskName
    if ($taskInfo.State -ne "Running") {
        # Start the task if it is not running
        Start-ScheduledTask -TaskName $taskName
        Write-Output "Task '$taskName' started."
    } else {
        Write-Output "Task '$taskName' is already running."
    }
}

function Copy-ScriptIfNotExists {
    $destinationPath = "C:\Scripts\Watch-SophosVPN.exe"

    # Get the path of the currently executing script
    $currentScriptPath = $MyInvocation.MyCommand.Path

    # Check if the destination file exists
    if (-not (Test-Path -Path $destinationPath)) {
        # Ensure the destination directory exists
        $destinationDirectory = [System.IO.Path]::GetDirectoryName($destinationPath)
        if (-not (Test-Path -Path $destinationDirectory)) {
            New-Item -ItemType Directory -Path $destinationDirectory -Force
        }

        # Copy the script to the destination
        Copy-Item -Path $currentScriptPath -Destination $destinationPath
        Write-Output "Script copied to $destinationPath."
    } else {
        Write-Output "Script already exists at $destinationPath."
    }
}

# Call the function
Copy-ScriptIfNotExists

Function Get-VPNStatus {
    param(
        [Parameter(Mandatory = $False)][string]$ConnectionName
    )

    # Initialize an empty dictionary
    $VPNStatus = @{}

    # Initialize VPN name variable for nested dictionary
    $currentConnection = ""

    # Get output on VPN state from SCCLI
    $SCCLIOutput = & "C:\Program Files (x86)\Sophos\Connect\sccli.exe" "list" "-d"

    # Regex pattern to capture IP address or name
    $patternIP = '.*Display Name: (.*)'

    # Regex pattern to capture key/value pairs
    $patternKV = '\s\s\s\s(.*):\s(.*)'

    # Find matches for display names
    $matchesDisplays = [regex]::Matches($SCCLIOutput, $patternIP)

        # Split the input into lines
    $lines = $matchesDisplays.Value -split '\s{2,}'

    # Initialize variables
    $currentConnection = ""

    # Regex patterns
    $connectionPattern = '^(?!Display Name:|Gateway:|VPN Type:|Auto-Connect:|IKE version:|Last connect time:|Last connect result:|Latency:|Favicon:|User authentication type:|IKE authentication type:|VPN state:|Auth state:)(.+)$'
    $keyValuePattern = '([^:]+):\s*(.*)'

    foreach ($line in $lines) {
        if ($line -match $connectionPattern) {
            # New connection found
            $currentConnection = $line.Trim()
            $VPNStatus[$currentConnection] = @{}
        } elseif ($line -match $keyValuePattern) {
            # Key-value pair found
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $VPNStatus[$currentConnection][$key] = $value
        }
    }

    $VPNStatus.Remove("Connections:")

    If ($ConnectionName) {
        Return $VPNStatus["$ConnectionName"]
    } Else {
        # Output the dictionary
        Return $VPNStatus
    }
}

# Function to check VPN connection status and handle reauthentication notifications
function Monitor-VPNConnections {
    $VPNConnections = Get-VPNStatus
    $oldStatus = @{}
    
    ForEach ($VPN in $VPNConnections.Keys) {
        $oldStatus[$VPN] = $VPNConnections[$VPN]["VPN state"]
    }

    While ($True) {
        $VPNConnections = Get-VPNStatus

        ForEach ($VPN in $VPNConnections.Keys) {
            $ConnectionName = $VPN
            $VPNDetails = $VPNConnections[$VPN]

            If ($VPNDetails["VPN state"] -eq "connected" -and $VPNDetails["VPN state"] -ne $oldStatus[$VPN]) {
                Write-Host "$VPN VPN has connected"
                Write-Log -LogLvl "INFO" -LogName "Monitor-SophosVPN.log" -LogMsg "The $ConnectionName VPN has connected"
                $oldStatus[$VPN] = "connected"
            }

            If ($VPNDetails["VPN state"] -eq "disconnected" -and $VPNDetails["VPN state"] -ne $oldStatus[$VPN]) {
                Write-Host "$VPN VPN has disconnected"
                Write-Log -LogLvl "INFO" -LogName "Monitor-SophosVPN.log" -LogMsg "The $ConnectionName VPN has disconnected"
                $oldStatus[$VPN] = "disconnected"

                # Handle authentication notifications
                If ($VPNDetails["Auth state"] -like "*need username/password*") {
                    $Text_AppName = "Sophos Connect"
                    $VPNType = ($VPNDetails["VPN Type"]).ToUpper()
                    $Title = "The $($VPNDetails["Display Name"]) $VPNType VPN requires reauthentication"
                    $Message = "`nYou need to reauthenticate to the VPN. Check the Sophos Connect application and/or your mobile device for DUO notifications."
                    Show-ToastNotification -Text_AppName $Text_AppName -Title $Title -Message $Message
                    Write-Log -LogLvl "INFO" -LogName "Monitor-SophosVPN.log" -LogMsg "Displaying notification for $ConnectionName $VPNType VPN disconnected"
                } ElseIf ($VPNDetails["Auth state"] -like "*need one-time password*") {
                    $Text_AppName = "Sophos Connect"
                    $VPNType = ($VPNDetails["VPN Type"]).ToUpper()
                    $Title = "The $($VPNDetails["Display Name"]) $VPNType VPN requires OTP reauthentication"
                    $Message = "`nYou need to reauthenticate to the VPN. Check the Sophos Connect application and enter the one-time password."
                    Show-ToastNotification -Text_AppName $Text_AppName -Title $Title -Message $Message
                    Write-Log -LogLvl "INFO" -LogName "Monitor-SophosVPN.log" -LogMsg "Displaying notification for $ConnectionName $VPNType VPN disconnected"
                } Else {
                    $Text_AppName = "Sophos Connect"
                    $VPNType = ($VPNDetails["VPN Type"]).ToUpper()
                    $Title = "The $($VPNDetails["Display Name"]) $VPNType VPN has disconnected"
                    $Message = "`nThe VPN has disconnected. If you disconnected the VPN manually you can close this notification"
                    Show-ToastNotification -Text_AppName $Text_AppName -Title $Title -Message $Message -Scenario "default"
                    Write-Log -LogLvl "INFO" -LogName "Monitor-SophosVPN.log" -LogMsg "Displaying notification for $ConnectionName $VPNType VPN disconnected"
                }
            }
        }
        Start-Sleep 5
    }
}

Function Register-NotificationApp($AppID,$AppDisplayName,$ToastIcon) {
    [int]$ShowInSettings = 0

    [int]$IconBackgroundColor = 0
	$IconUri = $ToastIcon
	
    $AppRegPath = "HKCU:\Software\Classes\AppUserModelId"
    $RegPath = "$AppRegPath\$AppID"
	
	$Notifications_Reg = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings'
	If(!(Test-Path -Path "$Notifications_Reg\$AppID")) 
		{
			New-Item -Path "$Notifications_Reg\$AppID" -Force
			New-ItemProperty -Path "$Notifications_Reg\$AppID" -Name 'ShowInActionCenter' -Value 1 -PropertyType 'DWORD' -Force
		}

	If((Get-ItemProperty -Path "$Notifications_Reg\$AppID" -Name 'ShowInActionCenter' -ErrorAction SilentlyContinue).ShowInActionCenter -ne '1') 
		{
			New-ItemProperty -Path "$Notifications_Reg\$AppID" -Name 'ShowInActionCenter' -Value 1 -PropertyType 'DWORD' -Force
		}	
		
    try {
        if (-NOT(Test-Path $RegPath)) {
            New-Item -Path $AppRegPath -Name $AppID -Force | Out-Null
        }
        $DisplayName = Get-ItemProperty -Path $RegPath -Name DisplayName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue
        if ($DisplayName -ne $AppDisplayName) {
            New-ItemProperty -Path $RegPath -Name DisplayName -Value $AppDisplayName -PropertyType String -Force | Out-Null
        }
        $ShowInSettingsValue = Get-ItemProperty -Path $RegPath -Name ShowInSettings -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ShowInSettings -ErrorAction SilentlyContinue
        if ($ShowInSettingsValue -ne $ShowInSettings) {
            New-ItemProperty -Path $RegPath -Name ShowInSettings -Value $ShowInSettings -PropertyType DWORD -Force | Out-Null
        }
		
		New-ItemProperty -Path $RegPath -Name IconUri -Value $IconUri -PropertyType ExpandString -Force | Out-Null	
		New-ItemProperty -Path $RegPath -Name IconBackgroundColor -Value $IconBackgroundColor -PropertyType ExpandString -Force | Out-Null		
		
    }
    catch {}
}

Function Show-ToastNotification {
    param(
		[Parameter(Mandatory = $True)] [string]$Text_AppName,
		[Parameter(Mandatory = $True)] [string]$Title,
		[Parameter(Mandatory = $True)] [string]$Message,
        [Parameter(Mandatory = $False)] [string]$Scenario = "reminder"
	)

    # Export the toast picture png to disk
    $ToastImage = "$env:TEMP\ToastPicture.png"
    [byte[]]$Bytes = [convert]::FromBase64String($Picture_Base64)
    [System.IO.File]::WriteAllBytes($ToastImage,$Bytes)

    # Export the toast picture png to disk
    $IconImage = "$env:TEMP\ToastIcon.png"
    [byte[]]$Bytes = [convert]::FromBase64String($Icon_Base64)
    [System.IO.File]::WriteAllBytes($IconImage,$Bytes)	

    # Configure toast actions (xml)
    $Actions = 
@"
<actions>
        <action activationType="protocol" arguments="Dismiss" content="Dismiss" />
</actions>	
"@


    [xml]$Toast = @"
<toast scenario="$Scenario">
    <visual>
    <binding template="ToastGeneric">
        <image placement="hero" src="$ToastImage"/>
        <image placement="appLogoOverride" hint-crop="circle" src="$ToastIcon"/>
        <text placement="attribution">$Attribution</text>
        <text>$Title</text>
        <group>
            <subgroup>
                <text hint-style="body" hint-wrap="true" >$Message</text>
            </subgroup>
        </group>
    </binding>
    </visual>
    $Actions
</toast>
"@	

    # Register notification app to allow toast notification
    Register-NotificationApp -AppID $Text_AppName -AppDisplayName $Text_AppName -ToastIcon $IconImage

    # Toast creation and display
    $Load = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    $Load = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    $ToastXml = New-Object -TypeName Windows.Data.Xml.Dom.XmlDocument
    $ToastXml.LoadXml($Toast.OuterXml)

    # Display the Toast
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($Text_AppName).Show($ToastXml)
}

##############################################################################################################
#                                                   Main                                                     #
##############################################################################################################

Write-Log -LogLvl "INFO" -LogName "Monitor-SophosVPN.log" -LogMsg "Starting Sophos VPN Status Monitor"
    
# Parse the output from SCCLI into a nested dictionary 
$VPNStatus = Get-VPNStatus

ForEach ($VPN in $VPNStatus.Keys) {
    $CurrentStatus = $($VPNStatus[$VPN]["VPN state"])
    Write-Log -LogLvl "INFO" -LogName "Monitor-SophosVPN.log" -LogMsg "Detected VPN: $VPN`tCurrent Status: $CurrentStatus"
}

# Infinite loop
Monitor-VPNConnections
