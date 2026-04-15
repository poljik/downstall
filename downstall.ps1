#Requires -RunAsAdministrator
#Requires -Version 5.1   # Windows 10 LTSC

<#
.SYNOPSIS
    Downstall - Script for downloading and installing software on Windows
.DESCRIPTION
    Supports downloading software from various sources (GitHub, Yandex Disk, Google Drive, etc.)
    and automatic installation with customizable parameters
.PARAMETER DownloadOnly
    Download files only without installation
.PARAMETER InstallOnly
    Install only without downloading (uses already downloaded files)
.PARAMETER Install
    Array of software names for installation/download
.EXAMPLE
    .\downstall.ps1 -DownloadOnly
    .\downstall.ps1 -InstallOnly opera,firefox
.NOTES
    Author: poljik 2019-2026
#>

[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Download files only without installation")]
    [Switch]$DownloadOnly,

    [Parameter(HelpMessage = "Install only without downloading")]
    [Switch]$InstallOnly
)

DynamicParam {
    $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
    
    $SoftwareSet = @()
    foreach ($file in @("soft.json", "soft+.json")) {
        if (Test-Path "$PSScriptRoot\$file") {
            $SoftwareSet += (Get-Content "$PSScriptRoot\$file" -Raw | ConvertFrom-Json).SoftwareName
        }
    }

    if ($SoftwareSet.Count -gt 0) {
        $Attr = New-Object System.Management.Automation.ParameterAttribute
        $Attr.Mandatory = $false
        $Attr.Position = 1
        
        $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($SoftwareSet)
        $AttrCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $AttrCollection.Add($Attr)
        $AttrCollection.Add($ValidateSet)

        $RuntimeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Install', [string[]], $AttrCollection)
        $RuntimeParameterDictionary.Add('Install', $RuntimeParam)
    }
    return $RuntimeParameterDictionary
}

begin {
    [string[]]$InstallList = $PsBoundParameters['Install']
    $Script:SystemArchitecture = if ([System.Environment]::Is64BitOperatingSystem) { 64 } else { 86 }
    
    # Mask PowerShell as a regular Google Chrome browser so websites don't block downloads
    $Global:UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

    # ==========================================
    # HELPER FUNCTIONS
    # ==========================================

    function Find-TargetDirectory {
        param ([String]$SearchPattern, [String]$DirectoryName)
        
        $FoundFiles = Get-ChildItem -Path $PSScriptRoot -Include $SearchPattern -Recurse -Force -ErrorAction Ignore
        if ($FoundFiles.Count -gt 0) {
            $TargetFile = $FoundFiles[0]
            $ResultFileName = $TargetFile.Name
            $ResultFilePath = if ($TargetFile.Attributes -match 'Directory') { $TargetFile.FullName } else { $TargetFile.DirectoryName }
        }
        else {
            $ResultFileName = $null
            $FoundDirectories = Get-ChildItem -Path $PSScriptRoot -Include $DirectoryName -Recurse -Force -ErrorAction Ignore
            if ($FoundDirectories.Count -gt 0) {
                $ResultFilePath = $FoundDirectories[0].FullName
            }
            else {
                $ResultFilePath = Join-Path -Path $PSScriptRoot -ChildPath $DirectoryName.Replace("*", " ")
            }
        }
        return @{ FilePath = $ResultFilePath; FileName = $ResultFileName }
    }

    function Test-FileUpdateRequired {
    param (
        [String]$LocalFilePath, 
        [String]$DownloadUrl
    )

    $LocalFile = Get-Item -Path $LocalFilePath -ErrorAction SilentlyContinue
    if (-not $LocalFile) { return $true }

    $OldSize = $LocalFile.Length
    
    $SkipCheck = if ($PSVersionTable.PSVersion.Major -ge 6) { @{ SkipCertificateCheck = $true } } else { @{} }
    $CommonHeaders = @{ Referer = $DownloadUrl }

    # Attempt 1: HEAD request
    try {
        $WebResponse = Invoke-WebRequest -Uri $DownloadUrl -Method Head -Headers $CommonHeaders -UserAgent $Global:UserAgent @SkipCheck -ErrorAction Stop
        $NewSize = $WebResponse.Headers.'Content-Length'
    }
    catch {
        # Extract HTTP Status Code cross-platform (works in both PS 5.1 and PS 7+)
        $StatusCode = $null
        if ($_.Exception.Response) {
            $StatusCode = [int]$_.Exception.Response.StatusCode # PS 5.1
        } elseif ($null -ne $_.Exception.StatusCode) {
            $StatusCode = [int]$_.Exception.StatusCode # PS 7+
        }

        # If it is an HTTP error (4xx or 5xx), the server might be blocking HEAD requests. Try GET fallback.
        if ($StatusCode -match "^(4|5)\d\d$") {
            Write-Warning "Test-FileUpdateRequired: HEAD failed for '$($LocalFile.Name)' (HTTP $StatusCode), trying GET..."
            
            try {
                # IMPORTANT: Request only 1 byte to prevent downloading the entire file into RAM
                $GetHeaders = $CommonHeaders.Clone()
                $GetHeaders['Range'] = "bytes=0-0"
                
                $GetResponse = Invoke-WebRequest -Uri $DownloadUrl -Method Get -Headers $GetHeaders -UserAgent $Global:UserAgent @SkipCheck -MaximumRedirection 5 -ErrorAction Stop
                
                # A successful Range request returns HTTP 206 and a Content-Range header (e.g., "bytes 0-0/1234567")
                if ($GetResponse.Headers.'Content-Range') {
                    $NewSize = ($GetResponse.Headers.'Content-Range' -split '/')[-1]
                } else {
                    $NewSize = $GetResponse.Headers.'Content-Length'
                }
            }
            catch {
                Write-Warning "Test-FileUpdateRequired: GET fallback failed: $($_.Exception.Message)"
                return $false
            }
        } 
        else {
            # Network error (Timeout, DNS resolution failure, offline)
            Write-Warning "Test-FileUpdateRequired: Network error for '$($LocalFile.Name)': $($_.Exception.Message)"
            Write-Warning "Skipping download - cannot verify file size."
            return $false
        }
    }

    if ([string]::IsNullOrEmpty($NewSize)) {
        Write-Warning "Test-FileUpdateRequired: Content-Length unknown for '$($LocalFile.Name)', skipping download."
        return $false
    }

    # IMPORTANT: Use [long] (Int64) to prevent overflow exceptions on files larger than 2GB
    $NewSizeLong = if ($NewSize -is [Array]) { [long]$NewSize[0] } else { [long]$NewSize }

    if ($OldSize -ge 1GB) {
        $FormattedOldSize = "{0:N2} GB" -f ($OldSize / 1GB)
    }
    elseif ($OldSize -ge 1MB) {
        $FormattedOldSize = "{0:N2} MB" -f ($OldSize / 1MB)
    }
    elseif ($OldSize -ge 1KB) {
        $FormattedOldSize = "{0:N2} KB" -f ($OldSize / 1KB)
    }
    else {
        $FormattedOldSize = "$OldSize bytes"
    }

    if (($OldSize -eq $NewSizeLong) -or ($NewSizeLong -eq 0)) {
        Write-Host "Skip download '$($LocalFile.Name)' - Sizes match ($FormattedOldSize)."
        return $false
    }

    if ($NewSizeLong -ge 1GB) {
        $FormattedNewSizeLong = "{0:N2} GB" -f ($NewSizeLong / 1GB)
    }
    elseif ($NewSizeLong -ge 1MB) {
        $FormattedNewSizeLong = "{0:N2} MB" -f ($NewSizeLong / 1MB)
    }
    elseif ($NewSizeLong -ge 1KB) {
        $FormattedNewSizeLong = "{0:N2} KB" -f ($NewSizeLong / 1KB)
    }
    else {
        $FormattedNewSizeLong = "$NewSizeLong bytes"
    }

    Write-Host "Size mismatch: '$($SoftwareItem.SoftwareName)' - local=$FormattedOldSize, remote=$FormattedNewSizeLong. Download required."
    return $true
}

    function Get-AbsoluteUri {
        param ([String]$DownloadUrl)

        $SkipCheck = if ($PSVersionTable.PSVersion.Major -ge 6) { @{ SkipCertificateCheck = $true } } else { @{} }
        try {
            if ($PSVersionTable.PSVersion.Major -ge 6) {
                $Response = Invoke-WebRequest -Uri $DownloadUrl -Method Head @SkipCheck -UserAgent $Global:UserAgent -ErrorAction Stop
            } else {
                $Response = Invoke-WebRequest -Uri $DownloadUrl -UserAgent $Global:UserAgent -ErrorAction Stop
            }
        }
        catch {
            Write-Warning "'$($SoftwareItem.SoftwareName)' - Error: $($_.Exception.Message)"   
            return $null
        }
        
        $ResolvedUrl = $Response.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
        if ($ResolvedUrl) {
            return [string]$ResolvedUrl
        } else {
            return [string]$Response.BaseResponse.ResponseUri.AbsoluteUri
        }
    }

    function Invoke-Office365Setup {
        param ([String]$SetupDirectory)

        $ConfigXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
    <Language ID="MatchOS" />
    <Language ID="en-us" />
    <ExcludeApp ID="Access" />
    <ExcludeApp ID="Bing" />
    <ExcludeApp ID="Groove" />
    <ExcludeApp ID="Lync" />
    <ExcludeApp ID="OneDrive" />
    <ExcludeApp ID="OneNote" />
    <ExcludeApp ID="Outlook" />
    <ExcludeApp ID="Publisher" />
    <ExcludeApp ID="Teams" />
  </Product>
  <Product ID="ProofingTools">
    <Language ID="MatchOS" />
    <Language ID="en-us" />
  </Product>
</Add>
<Property Name="PinIconsToTaskbar" Value="FALSE" />
<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
<Updates Enabled="TRUE" />
<Display Level="Full" AcceptEULA="TRUE" />
<Logging Level="Off" />
</Configuration>
"@
        $RootDirectory = $SetupDirectory | Split-Path -Parent | Split-Path -Parent
        if (Test-Path "$RootDirectory\setup.exe") { return }

        "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" | ForEach-Object {
            $Null = New-Item -Path $_ -ItemType Directory -Force
            Set-ItemProperty -Path $_ -Name CountryCode -Value "std::wstring|US" -Force
        }
  
        if (-not (Test-Path $SetupDirectory)) { $Null = New-Item $SetupDirectory -ItemType Directory }
        
        Push-Location -Path $SetupDirectory
        try {
            $WebResult = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing -UserAgent $Global:UserAgent
            $DownloadUrl = ($WebResult.Links | Where-Object { $_.outerHTML -match "officedeploymenttoll" }).href
            Invoke-WebRequest -Uri $DownloadUrl -OutFile "officedeploymenttool.exe" -UseBasicParsing -UserAgent $Global:UserAgent -Verbose

            Start-Process -FilePath "officedeploymenttool.exe" -ArgumentList "/quiet /extract:officedeploymenttool" -Wait

            Move-Item -Path "officedeploymenttool\setup.exe" -Destination "./" -Force
            Start-Sleep -Seconds 1
            Remove-Item -Path "officedeploymenttool*" -Recurse -Force
            
            $ConfigXmlPath = Join-Path -Path $SetupDirectory -ChildPath "config.xml"
            Set-Content -Path $ConfigXmlPath -Value $ConfigXml
            
            Start-Process -FilePath "setup.exe" -ArgumentList "/download `"config.xml`"" -Wait
        }
        finally {
            Pop-Location
        }
    }

    function Install-SoftwarePackage {
        param (
            [Parameter(Mandatory=$true)] [Object]$SoftwareItem,
            [Parameter(Mandatory=$true)] [String]$FilePath,
            [Parameter(Mandatory=$true)] [String]$FileName
        )

        $FullFilePath = Join-Path -Path $FilePath -ChildPath $FileName
        $SoftwareName = $SoftwareItem.SoftwareName
        Write-Warning "Install $SoftwareName"

        # Special Cases
        switch -Wildcard ($SoftwareName) {
            "imagine_plugin_*" {
                if ((Test-Path "$Env:Programfiles\WinRAR\Winrar.exe") -and (Test-Path "$Env:LOCALAPPDATA\Imagine\Plugin")) {
                    Start-Process -FilePath "$Env:Programfiles\WinRAR\Winrar.exe" -ArgumentList "x -o+", "`"$FullFilePath`"", "$Env:LOCALAPPDATA\Imagine\Plugin" -WindowStyle Hidden
                }
                return
            }
            "mas" {
                $OsMajor = [System.Environment]::OSVersion.Version.Major
                Write-Host "You have Windows $OsMajor"
                if ($OsMajor -ge 10) {
                    Write-Host "KMS38 windows activation and Office Ohook activation"
                    & ([ScriptBlock]::Create((Invoke-RestMethod https://get.activated.win))) /KMS38 /Ohook
                } else {
                    Write-Host "KMS windows activation only"
                    & ([ScriptBlock]::Create((Invoke-RestMethod https://get.activated.win))) /KMS-WindowsOffice /KMS-RenewalTask
                }
                return
            }
            "office365" {
                $SetupPath = $FilePath | Split-Path -Parent | Split-Path -Parent
                Start-Process -FilePath "$SetupPath\setup.exe" -ArgumentList "/configure `"$SetupPath\config.xml`"" -Wait
                return
            }
            "office*" {
                @("Word.lnk", "Excel.lnk", "Powerpoint.lnk") | ForEach-Object {
                    $ShortcutPath = "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\$_"
                    if (Test-Path $ShortcutPath) { Copy-Item -Path $ShortcutPath -Destination $Env:PUBLIC\Desktop -Force }
                }
                return
            }
            "ventoy_linux" {
                Write-Warning "$SoftwareName is just for Linux :)"
                return
            }
            "platelschik_eaes_rate" {
                $EaesPath = "${Env:ProgramFiles(x86)}\МНС\Плательщик ЕАЭС*\description"
                if (Test-Path $EaesPath) { $SetupDir = Get-ChildItem -Path $EaesPath -Force }
                elseif ($Found = Get-ChildItem -Path D:\* -Include $SoftwareItem.SetupPattern -Recurse -Force) { $SetupDir = $Found | Split-Path -Parent }
                else { Write-Warning "Can't find platelschik_eaes directory"; return }
                
                Remove-Item "$SetupDir\reduced_rate_nds.stbl" -Force -ErrorAction Ignore
                Remove-Item "$SetupDir\$($SoftwareItem.SetupPattern)" -Force
                if (Test-Path "$Env:Programfiles\WinRAR\Winrar.exe") {
                    Start-Process -FilePath "$Env:Programfiles\WinRAR\Winrar.exe" -ArgumentList "x -o+", "`"$FullFilePath`"", "`"$SetupDir`"" -WindowStyle Hidden
                }
                return
            }
            "rdpwrap_ini" {
                $RdpPath = "$Env:ProgramFiles\RDP Wrapper\"
                if (Test-Path $RdpPath) { $SetupDir = $RdpPath }
                elseif ($Found = Get-ChildItem -Path D:\* -Directory -Recurse -Force | Where-Object { $_.Name -match "RDP Wrapper" }) { $SetupDir = $Found.FullName }
                else { Write-Warning "Can't find RDPWrap directory"; return }
                
                Stop-Service TermService -Force
                Remove-Item "$SetupDir\$($SoftwareItem.SearchPattern)" -Force
                Move-Item -Path $FullFilePath -Destination $SetupDir -Force
                Start-Service TermService
                return
            }
        }

        # Standard Installation
        $BypassInstall = $false
        if (Test-Path "$PSScriptRoot/downstall+.ps1") {
            $BypassInstall = & "$PSScriptRoot/downstall+.ps1"
        }

        if (-not $BypassInstall) {
            $IsArchive = $FileName -match "\.(zip|7z|rar|gz)$"
            $ExtractSingleExe = [bool]$SoftwareItem.ExtractSingleExe

            if ($IsArchive) {
                # Determine file extension
                $Extension = [System.IO.Path]::GetExtension($FileName).ToLower()
                
                # Determine extraction destination
                $DestDir = if ($SoftwareItem.SetupPattern) {
                    $TempDir = "$Env:Temp\temp_downstall_$($SoftwareItem.SoftwareName)"
                    $Null = New-Item -ItemType Directory -Force -Path $TempDir
                    $TempDir
                }
                else {
                    $DesktopDir = "$Env:PUBLIC\Desktop\$SoftwareName"
                    $Null = New-Item -ItemType Directory -Force -Path $DesktopDir
                    $DesktopDir
                }

                # 1. EXTRACT BASED ON FORMAT
                $ExtractSuccess = $false
                switch -Regex ($Extension) {
                    "\.zip$" {
                        Write-Host "Extracting ZIP using built-in Expand-Archive..."
                        Expand-Archive -Path $FullFilePath -DestinationPath $DestDir -Force
                        $ExtractSuccess = $true
                    }
                    "\.(gz|tgz|tar)$" {
                        Write-Host "Extracting GZ/TAR using built-in Windows tar.exe..."
                        Start-Process -FilePath "tar.exe" -ArgumentList "-xf `"$FullFilePath`" -C `"$DestDir`"" -Wait -NoNewWindow
                        $ExtractSuccess = $true
                    }
                    "\.(rar|7z)$" {
                        # Look for WinRAR or 7-Zip (Windows 10 lacks native support for these formats)
                        $WinRarPath = "$Env:Programfiles\WinRAR\Winrar.exe"
                        $7zPath = "$Env:Programfiles\7-Zip\7z.exe"
                        
                        if (Test-Path $WinRarPath) {
                            $WinRarArgs = if ($SoftwareItem.ArchiveArgs) { "x -o+ $($SoftwareItem.ArchiveArgs)" } else { "x -o+" }
                            Start-Process -FilePath $WinRarPath -ArgumentList "$WinRarArgs `"$FullFilePath`" `"$DestDir`"" -WindowStyle Hidden -Wait
                            $ExtractSuccess = $true
                        }
                        elseif (Test-Path $7zPath) {
                            Start-Process -FilePath $7zPath -ArgumentList "x `"$FullFilePath`" -o`"$DestDir`" -y" -WindowStyle Hidden -Wait
                            $ExtractSuccess = $true
                        }
                        else {
                            Write-Warning "Cannot extract '$Extension'. Please install WinRAR or 7-Zip first! (Windows 10 lacks native support)."
                        }
                    }
                    default {
                        Write-Warning "Unknown archive format: $Extension"
                    }
                }

                # 2. POST-EXTRACTION ACTIONS
                if ($ExtractSuccess -and $SoftwareItem.SetupPattern) {
                    $ExtractedItems = Get-ChildItem -Path "$DestDir\*" -Include $SoftwareItem.SetupPattern -Recurse -Force
                    
                    if ($ExtractedItems) {
                        $TargetFile = $ExtractedItems | Select-Object -First 1
                        
                        if ($ExtractSingleExe) {
                            # If only extraction is needed (e.g., Portable version to Desktop)
                            Copy-Item -Path $TargetFile.FullName -Destination "$Env:PUBLIC\Desktop" -Force
                            Write-Host "Copied $($TargetFile.Name) to Desktop."
                        }
                        else {
                            # Normal mode: run the installer
                            $LaunchArgs = if ($SoftwareItem.InstallArgs) { $SoftwareItem.InstallArgs } else { "" }
                            Start-Process -FilePath $TargetFile.FullName -ArgumentList $LaunchArgs -Wait
                            
                            # Specific action for Avest
                            if ($SoftwareName -eq "avest") {
                                $AvestSetup = Get-ChildItem -Path "$DestDir\*" -Include "setupAvCSPBel*.exe" -Recurse -Force | Select-Object -First 1
                                if ($AvestSetup) { 
                                    Start-Process -FilePath $AvestSetup.FullName -ArgumentList "/verysilent /devices=avToken,avPass,iKey" -Wait 
                                }
                            }
                        }
                    }
                    else {
                        Write-Warning "Setup pattern '$($SoftwareItem.SetupPattern)' not found in extracted archive."
                    }
                    
                    # Clean up temp directory
                    Remove-Item $DestDir -Recurse -Force
                }
            }
            elseif ($ExtractSingleExe) {
                if (Test-Path $FullFilePath) { Copy-Item -Path $FullFilePath -Destination $Env:PUBLIC\Desktop -Force }
            }
            else {
                # PROTECTION: Check for junk files (web pages instead of EXE)
                if ($FileName -match "\.(php|html|htm)$") {
                    Write-Warning "File '$FileName' appears to be a web page, not an installer. Download failed or blocked by server."
                    return
                }
                
                # PROTECTION: Strict check for the MZ magic bytes in executables
                if ($FileName -match "\.exe$") {
                    $Header = Get-Content -Path $FullFilePath -TotalCount 2 -Encoding Byte -ErrorAction SilentlyContinue
                    if ($Header -and ($Header[0] -ne 77 -or $Header[1] -ne 90)) {
                        # MZ = 77 90
                        Write-Warning "File '$FileName' is not a valid executable (Missing MZ header). Download was corrupted or blocked by bot-protection."
                        return
                    }
                }

                $LaunchArgs = if ($SoftwareItem.InstallArgs) { $SoftwareItem.InstallArgs } else { "" }
                Start-Process -FilePath $FullFilePath -ArgumentList $LaunchArgs -Wait
            }
        }

        # Post-Installation Actions
        switch ($SoftwareName) {
            "true_image" {
                Set-Service afcdpsrv, syncagentsrv -StartupType Disabled -PassThru -Confirm:$false -ErrorAction SilentlyContinue | Stop-Service > $null
            }
            "total_commander" {
                $Dest = if ($Script:SystemArchitecture -eq 64) { "${Env:Programfiles(x86)}\Total Commander" } else { "$Env:Programfiles\Total Commander" }
                Get-ChildItem -Path $FilePath -Include wincmd.key, TOTALCMD*.EXE -Recurse -Force -ErrorAction Ignore | Copy-Item -Destination $Dest -Force
            }
            "far" {
                $FarProfile = "$Env:APPDATA\Far Manager\Profile"
                if ($LuaFile = Get-ChildItem -Path $FilePath -Include Panel.Esc.lua -Recurse -Force -ErrorAction Ignore) {
                    $MacroPath = "$FarProfile\Macros\scripts\"
                    $Null = New-Item -Path $MacroPath -ItemType Directory -Force
                    Copy-Item -Path $LuaFile -Destination $MacroPath -Force
                }
                if ($DbFile = Get-ChildItem -Path $FilePath -Include generalconfig.db -Recurse -Force -ErrorAction Ignore) {
                    $Null = New-Item -Path $FarProfile -ItemType Directory -Force
                    Copy-Item -Path $DbFile -Destination $FarProfile -Force
                }
                
                $BitSuffix = if ($Script:SystemArchitecture -eq 64) { " (x64)" } else { "" }
                $ShortcutTarget = "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Far Manager 3$BitSuffix\Far Manager 3$BitSuffix.lnk"
                
                $retry = 0
                while (-not (Test-Path $ShortcutTarget) -and $retry -lt 10) { Start-Sleep -Seconds 1; $retry++ }
                if (Test-Path $ShortcutTarget) { Copy-Item -Path $ShortcutTarget -Destination $Env:PUBLIC\Desktop -Force }
            }
            "adobe_reader" {
                $AcrobatShortcut = "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Acrobat.lnk"
                if (Test-Path $AcrobatShortcut) { Copy-Item -Path $AcrobatShortcut -Destination $Env:PUBLIC\Desktop -Force }
                $RegPath = "HKLM:\Software\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown"
                $Null = New-Item -Path $RegPath -ItemType Directory -Force
                Set-ItemProperty -Path $RegPath -Name bAcroSuppressUpsell -Value 1 -Force
            }
            "viber" {
                $retry = 0
                while ((Get-Process | Where-Object Path -match "vibersetup") -and $retry -lt 60) { Start-Sleep -Seconds 1; $retry++ }
                Get-Process | Where-Object Path -match "Viber" | Stop-Process -Force
            }
            "winrar" {
                if ((Test-Path "$Env:Programfiles\WinRAR\Winrar.exe") -and ($KeyFile = Get-ChildItem -Path $FilePath -Include "rarreg.key" -Recurse -Force -ErrorAction Ignore)) {
                    Copy-Item $KeyFile "$Env:Programfiles\WinRAR" -Force
                    $WinRarShortcut = "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\WinRAR\WinRAR.lnk"
                    if (Test-Path $WinRarShortcut) { Copy-Item -Path $WinRarShortcut -Destination $Env:PUBLIC\Desktop -Force }
                }
            }
            "edeclaration" {
                $Target = "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\EDeclaration\Запустить EDeclaration.lnk"
                if (Test-Path $Target) { 
                    Copy-Item -Path $Target -Destination "$Env:PUBLIC\Desktop\EDeclaration.lnk" -Force 
                }
            }
            "platelschik_eaes" {
                $EaesShortcut = Get-ChildItem -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\МНС\" -Include "Плательщик ЕАЭС*.lnk" -Recurse -ErrorAction Ignore | Select-Object -First 1
                if ($EaesShortcut) {
                    Copy-Item -Path $EaesShortcut.FullName -Destination "$Env:PUBLIC\Desktop\Плательщик ЕАЭС.lnk" -Force
                }
            }
        }
    }

    function Update-DynamicSoftwareUrls {
        param ([Array]$SoftwareArray, [Array]$ActiveInstalls)
        
        $IrfanViewVersion = $null
        
        foreach ($Soft in $SoftwareArray) {
            if ($Soft.SoftwareName -notin $ActiveInstalls) { continue }

            switch ($Soft.SoftwareName) {
                { $_ -match "^irfanview" } {
                    if (-not $IrfanViewVersion) {
                        $HtmlContent = (Invoke-WebRequest -Uri "https://www.irfanview.com/" -UserAgent $Global:UserAgent).Content
                        $Match = [regex]::Match($HtmlContent, "(?i)version\s*([\d\.]+)")
                        if ($Match.Success) { $IrfanViewVersion = $Match.Groups[1].Value.Replace(".", "") }
                    }
                    $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#irfanviewVersion', $IrfanViewVersion)
                }
                "opera" {
                    $BaseUri = 'https://get.geo.opera.com/pub/opera/desktop/'
                    $Links = (Invoke-WebRequest -Uri $BaseUri -UseBasicParsing -UserAgent $Global:UserAgent).links.href
    
                    $LatestVer = @($Links | Where-Object { $_ -match "^\d+\.\d+\.\d+\.\d+/?$" }) | 
                    Sort-Object -Property { [version]($_ -replace '/', '') } | 
                    Select-Object -Last 1
        
                    $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#operaUrl', "$BaseUri$LatestVer" + "win/")
                }
                { $_ -match "^softethervpn" } {
                    $BaseUri = 'http://www.softether-download.com'
                    $Links = (Invoke-WebRequest -Uri "$BaseUri/files/softether/" -UseBasicParsing -UserAgent $Global:UserAgent).links.href
                    
                    # Filter RTM versions and sort them correctly by extracting the embedded date (YYYY.MM.DD)
                    $LatestNode = @($Links | Where-Object { $_ -match "v\d+\.\d+-\d+-rtm-(\d{4}\.\d{2}\.\d{2})-tree" }) | 
                    Sort-Object -Property { 
                        $DateStr = [regex]::Match($_, "rtm-(\d{4}\.\d{2}\.\d{2})-tree").Groups[1].Value
                        [datetime]::ParseExact($DateStr, "yyyy.MM.dd", $null) 
                    } | Select-Object -Last 1

                    if ($Soft.SoftwareName -eq "softethervpn") {
                        $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#softethervpnUrl', "$BaseUri$LatestNode" + "Windows/SoftEther_VPN_Client/")
                    }
                    else {
                        $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#softethervpnServerUrl', "$BaseUri$LatestNode" + "Windows/SoftEther_VPN_Server_and_VPN_Bridge/")
                    }
                }
            }
        }
    }

    function Invoke-SoftwareDeployment {
        param ([Object]$SoftwareItem)

        $DirInfo = Find-TargetDirectory -SearchPattern $SoftwareItem.SearchPattern -DirectoryName $SoftwareItem.DirectoryName
        $FilePath = $DirInfo.FilePath
        $CurrentFileName = $DirInfo.FileName

        if ($SoftwareItem.SoftwareName -eq "office365") {
            Invoke-Office365Setup -SetupDirectory $FilePath
            $DirInfo = Find-TargetDirectory -SearchPattern $SoftwareItem.SearchPattern -DirectoryName $SoftwareItem.DirectoryName
            $CurrentFileName = $DirInfo.FileName
        }

        $LinkFileName = $CurrentFileName
        [string]$DownloadUrl = $SoftwareItem.DownloadUrl

        if ($DownloadUrl -and -not $InstallOnly) {
            $DownloadLink = $DownloadUrl
            
            # Resolve Yandex/Google/GitHub links
            if ($DownloadUrl -match '\.yandex\.|yadi\.sk') {
                $UrlToken = "/d/"
                $BeginIndex = $DownloadUrl.IndexOf($UrlToken) + $UrlToken.Length
                $EndIndex = $DownloadUrl.IndexOf("/", $BeginIndex)
                
                if ($EndIndex -eq -1) {
                    $UrlDecoded = [System.Web.HttpUtility]::UrlEncode($DownloadUrl)
                    $YandexApi = Invoke-RestMethod "https://cloud-api.yandex.net/v1/disk/public/resources?public_key=$UrlDecoded&limit=200" -UserAgent $Global:UserAgent -ErrorAction SilentlyContinue
                }
                else {
                    $UrlDecoded = [System.Web.HttpUtility]::UrlEncode($DownloadUrl.Substring(0, $EndIndex))
                    $YandexPath = $DownloadUrl.Substring($EndIndex)
                    $YandexApi = Invoke-RestMethod "https://cloud-api.yandex.net/v1/disk/public/resources?public_key=$UrlDecoded&path=$YandexPath&limit=200" -UserAgent $Global:UserAgent -ErrorAction SilentlyContinue
                }
                
                if ($YandexApi.type -eq "file") {
                    $DownloadLink = (Invoke-RestMethod "https://cloud-api.yandex.net:443/v1/disk/public/resources/download?public_key=$UrlDecoded" -UserAgent $Global:UserAgent).href
                    $LinkFileName = [uri]::UnescapeDataString(($DownloadLink -split "filename=")[1].Split("&")[0])
                }
                else {
                    $LatestItem = $YandexApi._embedded.items | Where-Object { $_.name -like $SoftwareItem.SearchPattern } | Sort-Object -Property created | Select-Object -Last 1
                    if ($LatestItem) {
                        $LinkFileName = $LatestItem.name
                        $DownloadLink = $LatestItem.file
                    }
                }
            }
            elseif ($DownloadUrl -match "drive\.google\.com/file/d") {
                $GoogleResponse = Invoke-WebRequest $DownloadUrl -UseBasicParsing -MaximumRedirection 0 -UserAgent $Global:UserAgent -ErrorAction SilentlyContinue
                $DriveId = ($GoogleResponse.links.href | Where-Object { $_ -match "drive\.google\.com/file/d" } | Select-Object -First 1) -split "/" | Select-Object -Last 2 | Select-Object -First 1
                
                $TitleMatch = [regex]::Match($GoogleResponse.Content, 'meta property="og:title" content="(.*?)"')
                if ($TitleMatch.Success) { $LinkFileName = $TitleMatch.Groups[1].Value }
                $DownloadLink = "https://drive.google.com/uc?export=download&id=$DriveId"
            }
            else {
                # Force strictly to string to prevent array injection
                if (-not ([string]$AbsUrl = Get-AbsoluteUri -DownloadUrl $DownloadUrl)) { return }
                $DownloadUrl = $AbsUrl -replace "/tag/", "/expanded_assets/"
                $DownloadLink = $DownloadUrl
                
                $SkipCheck = if ($PSVersionTable.PSVersion.Major -ge 6) { @{ SkipCertificateCheck = $true } } else { @{} }
                $WebResponse = Invoke-WebRequest -Uri $DownloadUrl -UseBasicParsing -DisableKeepAlive -UserAgent $Global:UserAgent @SkipCheck -ErrorAction SilentlyContinue

                $Pattern = if ($SoftwareItem.SoftwareName -eq "imagine") { "*download.php?arch=x64&unicode=1&full=0&setup=1*" } else { $SoftwareItem.SearchPattern }
                $FoundLinks = @($WebResponse.links.href | Where-Object { $_ -like $Pattern })
        
                if ($FoundLinks.Count -gt 0) {
                    $DownloadLink = [string]$FoundLinks[0]
                    if ($SoftwareItem.SoftwareName -eq "vvoddpu" -and $FoundLinks[0] -like "*0.zip") { $DownloadLink = [string]$FoundLinks[1] }
                    if ($SoftwareItem.SoftwareName -eq "victoria") { $DownloadLink = [string]$FoundLinks[1].Substring($FoundLinks[1].LastIndexOf("https:")) }
            
                    if (-not $DownloadLink.StartsWith("http")) {
                        if ($DownloadLink.Contains("/")) {
                            $DownloadLink = if ($DownloadLink.StartsWith("/")) { $DownloadUrl.Substring(0, $DownloadUrl.IndexOf("/", 8)) + $DownloadLink } else { $DownloadUrl.Substring(0, $DownloadUrl.IndexOf("/", 8) + 1) + $DownloadLink }
                        }
                        else {
                            $DownloadLink = $DownloadUrl + $DownloadLink
                        }
                    }
                }
                elseif ($DownloadLink -match '//github\.com/') {
                    $DownloadLink = $null
                }
        
                if ($DownloadLink) {
                    $LinkFileName = ($DownloadLink | Split-Path -Leaf) -replace "\?.*$", "" -replace "%20", " "
                }
        
                if ($SoftwareItem.SoftwareName -eq "imagine" -and $FoundLinks.Count -gt 0) {
                    $Ver = ($FoundLinks[0] -split "version=")[1]
                    $LinkFileName = "Imagine_${Ver}_x64_Unicode.exe"
                }
                
                if ($SoftwareItem.SoftwareName -eq "cloudflare_warp" -and $LinkFileName) {
                    $LinkFileName = "cloudflare_warp_$LinkFileName.msi"
                }
            }

            if ($SoftwareItem.SoftwareName -eq "platelschik_eaes" -and $DownloadLink) { $DownloadLink = [System.Uri]::EscapeDataString($DownloadLink) }

            # Download Logic
            if ($LinkFileName -and ($CurrentFileName -ne $LinkFileName -or (Test-FileUpdateRequired -LocalFilePath "$FilePath\$CurrentFileName" -DownloadUrl $DownloadLink))) {
                Write-Warning "+$LinkFileName"
                if (-not (Test-Path $FilePath)) { $Null = New-Item $FilePath -ItemType Directory }

                try {
                    Invoke-WebRequest -Uri $DownloadLink -OutFile "$FilePath\$LinkFileName" -Headers @{Referer = $DownloadLink } -UserAgent $Global:UserAgent -ErrorAction Stop
                }
                catch {
                    Write-Warning "$CurrentFileName - Error: $($_.Exception.Message)"
                    $LinkFileName = $null # Reset so it doesn't try to install a failed download
                }

                if ($LinkFileName -and $CurrentFileName -and ($CurrentFileName -ne $LinkFileName)) {
                    Write-Warning "-$CurrentFileName deleted!"
                    Remove-Item -Path "$FilePath\$CurrentFileName" -Force -Recurse
                }
            }
        }

        # Install Phase (With protection against empty filenames)
        if (-not $DownloadOnly) {
            if ([string]::IsNullOrWhiteSpace($LinkFileName)) {
                Write-Warning "Skipping installation of '$($SoftwareItem.SoftwareName)': File not found locally and download failed/unavailable."
            } else {
                Install-SoftwarePackage -SoftwareItem $SoftwareItem -FilePath $FilePath -FileName $LinkFileName
            }
        }
    }
}

# ==========================================
# MAIN EXECUTION
# ==========================================
process {
    $Host.UI.RawUI.WindowTitle = "Downstall script 4 Windows $([char]0x00A9) poljik 2019-2026"
 
    # Enforce TLS 1.2 for PS 5.1 compatibility with modern HTTPS servers
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = "SilentlyContinue"

    if ($DownloadOnly -and $InstallOnly) {
        Write-Warning "ok, do nothing!"
        exit
    }
  
    while (-not $InstallOnly -and -not ([System.Net.Dns]::GetHostAddresses('google.com'))) {
        $InstallOnly = $true
        Write-Warning "Internet is missing, -InstallOnly activated"
    }

    $SoftwareList = @()
    if (Test-Path "$PSScriptRoot\soft.json") { $SoftwareList += Get-Content -Path "$PSScriptRoot\soft.json" | ConvertFrom-Json }
    else { Write-Warning "soft.json is missing" }
    
    if (Test-Path "$PSScriptRoot\soft+.json") { $SoftwareList += Get-Content -Path "$PSScriptRoot\soft+.json" | ConvertFrom-Json }

    if (-not $InstallList) {
        $DownloadOnly = $true
        $InstallList = $SoftwareList.SoftwareName
    }

    if (-not $InstallOnly) {
        Update-DynamicSoftwareUrls -SoftwareArray $SoftwareList -ActiveInstalls $InstallList
    }

    foreach ($InstallName in $InstallList) {
        $TargetSoftware = $SoftwareList | Where-Object { $_.SoftwareName -eq $InstallName } | Select-Object -First 1
        if ($TargetSoftware) {
            Invoke-SoftwareDeployment -SoftwareItem $TargetSoftware
        }
    }
}