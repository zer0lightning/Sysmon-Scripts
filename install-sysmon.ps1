# Install Sysmon if it isn't installed, and if it is, update the config (if it needs it)
# (C) Nathaniel Roach 2021

#Requires -RunAsAdministrator
[cmdletbinding()]
Param(
    [Parameter(Mandatory=$false)] [Switch]$ForceUpdate = $false,
    [Parameter(Mandatory=$false)] [Switch]$ForceConfig = $false,
    [Parameter(Mandatory=$false)] [Switch]$ForceBinary = $false
    )

Set-strictmode -version latest

$ErrorActionPreference = "Stop"

$SysmonRemotePath = "\\domain\DFS\ComputerAccessibleShare\sysmon"
$SysmonExe32 = "sysmon\sysmon.exe"
$SysmonExe64 = "sysmon\sysmon64.exe"
$SysmonConfig = "sysmon-config.xml"
$SysmonLocalPath = "C:\Windows"

$SysmonConfigRemotePath = ($SysmonRemotePath + "\" + $SysmonConfig)
$SysmonConfigLocalPath = ($SysmonLocalPath + "\" + $SysmonConfig)

If ($ForceUpdate) {
    $ForceConfig = $true
    $ForceBinary = $true
}

$IntConfigUpdated = $false
$IntBinaryUpdated = $false
$IntServiceRegistered = $false

$HostBitness = (gwmi win32_operatingsystem | select osarchitecture).osarchitecture
if ($HostBitness -ne "64-bit")
{
	$SysmonRemoteExeFullPath = ($SysmonRemotePath + "\" + $SysmonExe32)
    $SysmonServiceName = "sysmon"
} else {
    $SysmonRemoteExeFullPath = ($SysmonRemotePath + "\" + $SysmonExe64)
    $SysmonServiceName = "sysmon64"
}

Write-Verbose ("Using the following EXE for service install: " + $SysmonRemoteExeFullPath)

function Update-SysmonConfig {
    Param(
    [Parameter(Mandatory=$false)]
    [Switch]$NoService = $false
    )
    if (! $IntConfigUpdated) {
        try {
            Write-Verbose "Copying config..."
            Copy-Item -Path $SysmonConfigRemotePath -Destination $SysmonLocalPath -Force
            $IntConfigUpdated = $true
            if (! $NoService) {
                Write-Verbose "Informing service of update..."
                try {
                    & $SysmonRemoteExeFullPath -c $SysmonConfigLocalPath
                } catch {
                    Write-Warning "Error applying config"
                }
            }
        } catch {
            Write-Error ("Error copying config: " + $_.Exception)
        }
    }
}

function Update-SysmonBinary {
    Param(
    [Parameter(Mandatory=$false)]
    [Switch]$NoService = $false
    )
    if (! $IntBinaryUpdated) {
        try {        
            Write-Verbose "Copying binary..."
            Copy-Item -Path $SysmonRemoteExeFullPath -Destination ($SysmonLocalPath + "\Sysmon.exe" + ".new") -Verbose
            
            if (! $NoService) {
                Write-Verbose "Uninstalling and swapping files"
                try {
                    Write-Verbose "Attempting uninstall..."
                    & $SysmonRemoteExeFullPath -u force
                    #Move-Item -Path ($SysmonLocalPath + "\Sysmon.exe") -Destination ($SysmonLocalPath + "\Sysmon.exe" + ".old") -Verbose -Force
                    Move-Item -Path ($SysmonLocalPath + "\Sysmon.exe" + ".new") -Destination ($SysmonLocalPath + "\Sysmon.exe") -Verbose -Force

                    Write-Verbose "Installing update"
                    & $SysmonRemoteExeFullPath -accepteula -i $SysmonConfigLocalPath
                } catch {
                    Write-Warning "Error restarting Sysmon during Update-SysmonBinary"
                }
            } else {   
                Move-Item -Path ($SysmonLocalPath + "\Sysmon.exe" + ".new") -Destination ($SysmonLocalPath + "\Sysmon.exe") -Verbose -Force
            }
            $IntBinaryUpdated = $true
        } catch {
            Write-Error ("Error copying binary: " + $_.Exception)
        }
    }
}

function Install-SysmonService {
    Write-Verbose "Starting install..."

    try { # making the directory if it doesn't exist
        If (!(Test-Path -Path $SysmonLocalPath -PathType Container)){
            Write-Verbose "Making directory"

            Remove-Item -Path $SysmonLocalPath -ErrorAction Ignore
            New-Item -Path $SysmonLocalPath -ItemType Directory
        }

        Update-SysmonConfig -NoService:$true
        Update-SysmonBinary -NoService:$true
                
    } catch {
        Write-Error ("Error copying files: " + $_.exception)
    }
    
    try { # registering the service
        & $SysmonRemoteExeFullPath -accepteula -i $SysmonConfigLocalPath
        $IntServiceRegistered = $true
    } catch {
        Write-Error ("Error registring Sysmon service: " + $_.Exception)
    }
}

Try { 
    Write-Verbose "Checking presence of Sysmon service"
    Get-Service $SysmonServiceName -ErrorAction Stop | Out-Null
    Write-Verbose "Sysmon service already installed"

    Write-Verbose "Checking if binary needs updating, might take a second"
    #If ((Get-FileHash $($SysmonLocalPath + "\Sysmon.exe")).hash -ne (Get-FileHash $SysmonRemoteExeFullPath).hash) {
    If (!(Test-Path ($SysmonLocalPath + "\Sysmon.exe"))) {
        Update-SysmonBinary
    } elseif ((Get-Item ($SysmonLocalPath + "\Sysmon.exe")).LastWriteTimeUtc -lt (Get-Item $SysmonRemoteExeFullPath).LastWriteTimeUtc -or $ForceBinary) {
        Write-Host "Binary being updated..."
        Update-SysmonBinary
    }

    Write-Verbose "Checking if config need updating"
    If (!(Test-Path ($SysmonConfigLocalPath))) {
        Update-SysmonConfig
    } elseif ((Get-FileHash $SysmonConfigLocalPath).hash -ne (Get-FileHash $SysmonConfigRemotePath).hash -or $ForceConfig) {
        Write-Host "Config being updated..."
        Update-SysmonConfig
    }

} catch [Microsoft.PowerShell.Commands.ServiceCommandException] {
    Write-Host "Sysmon service missing, beginning install"
    Install-SysmonService
}

Write-Host "Done."
