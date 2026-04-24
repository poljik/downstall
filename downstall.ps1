#Requires -RunAsAdministrator
#Requires -Version 5.1  # Windows 10 LTSC

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
    $Files = "soft.json", "soft+.json"
    $SoftwareSet = foreach ($File in $Files) {
        $FilePath = Join-Path $PSScriptRoot $File
        if (Test-Path $FilePath) {
            $Json = Get-Content $FilePath -Raw | ConvertFrom-Json
            $Json.SoftwareName
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
    # ==========================================
    # ENVIRONMENT CONFIGURATION
    # ==========================================
    function Import-EnvironmentVariables {
        $EnvPath = Join-Path $PSScriptRoot ".env"
        if (Test-Path $EnvPath) {
            $Null = Get-Content -Path $EnvPath | Where-Object { $_ -match '^\s*[^#]' -and $_ -match '=' } | ForEach-Object {
                $Name, $Value = $_ -split '=', 2
                $Name = $Name.Trim()
                $Value = $Value.Trim() -replace '^["'']|["'']$', ''
                Set-Item -Path "Env:\$Name" -Value $Value
            }
        }
    }

    Import-EnvironmentVariables
    
    [string[]]$InstallList = $PsBoundParameters['Install']
    $Script:SystemArchitecture = if ([System.Environment]::Is64BitOperatingSystem) { 64 } else { 86 }
    
    $Global:UserAgent = $env:DOWNSTALL_USER_AGENT
    if (-not $Global:UserAgent) {
        $Global:UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    # ==========================================
    # HELPER FUNCTIONS
    # ==========================================
    function Start-DownloadWithProgress {
        param(
            [Parameter(Mandatory = $true)] [string]$Uri,
            [Parameter(Mandatory = $true)] [string]$OutFile,
            [string]$SoftwareName = "File"
        )
        # Enable progress bar locally (active only within this function!)
        $ProgressPreference = 'Continue'
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            try {
                $Null = Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UserAgent $Global:UserAgent -ErrorAction Stop
            }
            catch {
                throw $_
            }
        }
        else {
            $WebClient = [System.Net.HttpWebRequest]::Create($Uri)
            $WebClient.UserAgent = $Global:UserAgent
            $WebClient.Timeout = 60000
            $WebClient.AllowAutoRedirect = $true
            try {
                $Response = $WebClient.GetResponse()
                $TotalSize = $Response.ContentLength
                $Stream = $Response.GetResponseStream()
                $FileStream = [System.IO.File]::Create($OutFile)
                $Buffer = New-Object Byte[] 65536
                $TotalRead = 0
                $SwTotal = [System.Diagnostics.Stopwatch]::StartNew()
                $SwUI = [System.Diagnostics.Stopwatch]::StartNew()

                while (($BytesRead = $Stream.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
                    $FileStream.Write($Buffer, 0, $BytesRead)
                    $TotalRead += $BytesRead
                    # Refresh interface every 500 ms only
                    if ($SwUI.ElapsedMilliseconds -ge 500) {
                        $TotalSeconds = $SwTotal.Elapsed.TotalSeconds
                        $SpeedBps = if ($TotalSeconds -gt 0) { $TotalRead / $TotalSeconds } else { 0 }
                        $SpeedMbs = "{0:N2}" -f ($SpeedBps / 1MB)
                        $Percent = if ($TotalSize -gt 0) { [int](($TotalRead / $TotalSize) * 100) } else { 0 }
                        $ReadMB = "{0:N1}" -f ($TotalRead / 1MB)
                        $TotalMB = if ($TotalSize -gt 0) { "{0:N1}" -f ($TotalSize / 1MB) } else { "?" }
                        $RemainingSeconds = if ($SpeedBps -gt 0 -and $TotalSize -gt 0) { [int](($TotalSize - $TotalRead) / $SpeedBps) } else { -1 }

                        Write-Progress -Activity "Downloading $SoftwareName" `
                            -Status "$ReadMB / $TotalMB MB ($Percent%) | Speed: $SpeedMbs MB/s" `
                            -PercentComplete $Percent `
                            -SecondsRemaining $RemainingSeconds
                        $SwUI.Restart()
                    }
                }
                $FileStream.Close()
                $Stream.Close()
                $Response.Close()
                Write-Progress -Activity "Downloading $SoftwareName" -Completed
            }
            catch {
                if ($FileStream) { $FileStream.Close() }
                throw $_
            }
        }
    }

    function Find-TargetDirectory {
        param ([String]$SearchPattern, [String]$DirectoryName)
        $FoundFiles = Get-ChildItem -Path $PSScriptRoot -Include $SearchPattern -Recurse -Force -ErrorAction Ignore
        if ($FoundFiles.Count -gt 0) {
            $TargetFile = $FoundFiles[0]
            $ResultFileName = $TargetFile.Name
            $ResultFilePath = if ($TargetFile.Attributes -match 'Directory') { $TargetFile.FullName } else { $TargetFile.DirectoryName }
        }
        else {
            $ResultFileName = $Null
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
        param ([String]$LocalFilePath, [String]$DownloadUrl)
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
            $StatusCode = $Null
            if ($_.Exception.Response) { $StatusCode = [int]$_.Exception.Response.StatusCode } # PS 5.1 
            elseif ($Null -ne $_.Exception.StatusCode) { $StatusCode = [int]$_.Exception.StatusCode } # PS 7+
            
            # Fallback to GET with Range if HEAD is blocked (HTTP 4xx or 5xx)
            if ($StatusCode -match "^(4|5)\d\d$") {
                Write-Warning "Test-FileUpdateRequired: HEAD failed for '$($LocalFile.Name)' (HTTP $StatusCode), trying GET..."
                try {
                    # IMPORTANT: Request only 1 byte to prevent downloading the entire file into RAM
                    $GetHeaders = $CommonHeaders.Clone()
                    $GetHeaders['Range'] = "bytes=0-0"
                    $GetResponse = Invoke-WebRequest -Uri $DownloadUrl -Method Get -Headers $GetHeaders -UserAgent $Global:UserAgent @SkipCheck -MaximumRedirection 5 -ErrorAction Stop
                    
                    # A successful Range request returns HTTP 206 and a Content-Range header (e.g., "bytes 0-0/1234567")
                    if ($GetResponse.Headers.'Content-Range') { $NewSize = ($GetResponse.Headers.'Content-Range' -split '/')[-1] } 
                    else { $NewSize = $GetResponse.Headers.'Content-Length' }
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

        if ($OldSize -ge 1GB) { $FormattedOldSize = "{0:N2} GB" -f ($OldSize / 1GB) }
        elseif ($OldSize -ge 1MB) { $FormattedOldSize = "{0:N2} MB" -f ($OldSize / 1MB) }
        elseif ($OldSize -ge 1KB) { $FormattedOldSize = "{0:N2} KB" -f ($OldSize / 1KB) }
        else { $FormattedOldSize = "$OldSize bytes" }

        if (($OldSize -eq $NewSizeLong) -or ($NewSizeLong -eq 0)) {
            Write-Host "Skip download '$($LocalFile.Name)' - Sizes match ($FormattedOldSize)."
            return $false
        }

        if ($NewSizeLong -ge 1GB) { $FormattedNewSizeLong = "{0:N2} GB" -f ($NewSizeLong / 1GB) }
        elseif ($NewSizeLong -ge 1MB) { $FormattedNewSizeLong = "{0:N2} MB" -f ($NewSizeLong / 1MB) }
        elseif ($NewSizeLong -ge 1KB) { $FormattedNewSizeLong = "{0:N2} KB" -f ($NewSizeLong / 1KB) }
        else { $FormattedNewSizeLong = "$NewSizeLong bytes" }

        Write-Host "Size mismatch: '$($SoftwareItem.SoftwareName)' - local=$FormattedOldSize, remote=$FormattedNewSizeLong. Download required."
        return $true
    }

    function Get-AbsoluteUri {
        param ([String]$DownloadUrl)
        $SkipCheck = if ($PSVersionTable.PSVersion.Major -ge 6) { @{ SkipCertificateCheck = $true } } else { @{} }
        try {
            if ($PSVersionTable.PSVersion.Major -ge 6) { $Response = Invoke-WebRequest -Uri $DownloadUrl -Method Head @SkipCheck -UserAgent $Global:UserAgent -ErrorAction Stop } 
            else { $Response = Invoke-WebRequest -Uri $DownloadUrl -UserAgent $Global:UserAgent -ErrorAction Stop }
        }
        catch {
            Write-Warning "'$($SoftwareItem.SoftwareName)' - Error: $($_.Exception.Message)"   
            return $Null 
        }
        $ResolvedUrl = $Response.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
        if ($ResolvedUrl) { return [string]$ResolvedUrl } else { return [string]$Response.BaseResponse.ResponseUri.AbsoluteUri }
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
        # Check if setup.exe already exists in the root directory
        $RootDirectory = $SetupDirectory | Split-Path -Parent | Split-Path -Parent
        if (Test-Path (Join-Path $RootDirectory "setup.exe")) { return }
        
        # Configure Registry for Office Experiment Configs (ECS) to prevent setup issues
        $RegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs"
        if (-not (Test-Path $RegPath)) { 
            $Null = New-Item -Path $RegPath -ItemType Directory -Force 
        }
        Set-ItemProperty -Path $RegPath -Name CountryCode -Value "std::wstring|US" -Force

        # Ensure the setup directory exists
        if (-not (Test-Path $SetupDirectory)) { 
            $Null = New-Item $SetupDirectory -ItemType Directory 
        }
        Push-Location -Path $SetupDirectory
        try {
            # 1. Scrape the Microsoft download page for the latest ODT link
            $WebResult = Invoke-WebRequest -Uri $env:DOWNSTALL_OFFICE_ODT_URL -UseBasicParsing -UserAgent $Global:UserAgent
            # Filter links to find the direct .exe download URL
            $DownloadUrl = $WebResult.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | Select-Object -ExpandProperty href -First 1
            if (-not $DownloadUrl) {
                throw "Unable to find the Office Deployment Tool download link on the page."
            }

            # 2. Download the tool
            $ExePath = Join-Path $SetupDirectory "officedeploymenttool.exe"
            Start-DownloadWithProgress -Uri $DownloadUrl -OutFile $ExePath -SoftwareName "Office Setup Tool"

            # 3. Create a temp folder for extraction to avoid file conflicts
            $ExtractPath = Join-Path $SetupDirectory "odt_temp_extract"
            if (Test-Path $ExtractPath) { Remove-Item $ExtractPath -Recurse -Force }
            $Null = New-Item -Path $ExtractPath -ItemType Directory
            
            # Extract the contents silently
            Start-Process -FilePath $ExePath -ArgumentList "/quiet /extract:`"$ExtractPath`"" -Wait

            # Move only the necessary setup.exe to the main directory
            $ExtractedSetup = Join-Path $ExtractPath "setup.exe"
            if (Test-Path $ExtractedSetup) {
                Move-Item -Path $ExtractedSetup -Destination $SetupDirectory -Force
            }
            else {
                throw "Extraction failed: setup.exe not found in extracted files."
            }

            # 4. Generate the config.xml and begin downloading Office installation bits
            $ConfigXmlPath = Join-Path $SetupDirectory "config.xml"
            Set-Content -Path $ConfigXmlPath -Value $ConfigXml -Encoding UTF8
            $SetupExePath = Join-Path $SetupDirectory "setup.exe"
            Start-Process -FilePath $SetupExePath -ArgumentList "/download `"$ConfigXmlPath`"" -Wait
        }
        catch {
            Write-Warning "Office ODT Setup Error: $($_.Exception.Message)"
        }
        finally {
            # Cleanup: Remove installer and temp extraction folder
            Get-ChildItem -Path $SetupDirectory -Filter "officedeploymenttool*" | Remove-Item -Recurse -Force
            if (Test-Path $ExtractPath) { Remove-Item $ExtractPath -Recurse -Force }
            Pop-Location
        }
    }

    # ==========================================
    # INSTALLATION FUNCTIONS
    # ==========================================
    function Invoke-PostInstall {
        param (
            [Parameter(Mandatory = $true)] [Object]$SoftwareItem,
            [Parameter(Mandatory = $true)] [String]$DownloadDir,
            [Parameter(Mandatory = $true)] [String]$FileName
        )
        $SoftwareName = $SoftwareItem.SoftwareName
        $FullFilePath = Join-Path -Path $DownloadDir -ChildPath $FileName

        if ($SoftwareItem.PostInstall) {
            foreach ($Step in $SoftwareItem.PostInstall) {
                switch ($Step.Action) {
                    "Shortcut" {
                        $Src = $Null
                        $Retry = 0
                        while (-not $Src -and $Retry -lt 10) {
                            if ($Step.From -eq "StartMenu") { 
                                $Src = Get-ChildItem -Path $CommonPrograms -Filter $Step.File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                            }
                            else { 
                                $Src = Get-Item -Path (Join-Path $DownloadDir $Step.File) -ErrorAction SilentlyContinue
                            }
                            if (-not $Src) { Start-Sleep -Seconds 1; $Retry++ }
                        }
                        if ($Src) { 
                            $DestName = if ($Step.Rename) { $Step.Rename } else { $Src.Name }
                            $Null = Copy-Item -Path $Src.FullName -Destination (Join-Path $CommonDesktop $DestName) -Force 
                        }
                        else {
                            Write-Warning "Invoke-PostInstall: Shortcut source '$($Step.File)' not found after 10 seconds."
                        }
                    }
                    "Registry" {
                        if (-not (Test-Path $Step.Path)) {
                            $Null = New-Item -Path $Step.Path -ItemType Directory -Force -ErrorAction SilentlyContinue
                        }
                        Set-ItemProperty -Path $Step.Path -Name $Step.Name -Value $Step.Value -Force
                    }
                    "Service" {
                        $Step.Name | ForEach-Object {
                            Set-Service $_ -StartupType $Step.State -ErrorAction SilentlyContinue
                            if ($Step.Stop) { Stop-Service $_ -Force -ErrorAction SilentlyContinue }
                        }
                    }
                    "Copy" {
                        $DestinationPath = Invoke-Expression "`"$($Step.Dest)`""
                        $FoundFiles = Get-ChildItem -Path $DownloadDir -Filter $Step.Source -Recurse -Force -ErrorAction SilentlyContinue
                        if ($FoundFiles) {
                            if (-not (Test-Path $DestinationPath)) {
                                $Null = New-Item -Path $DestinationPath -ItemType Directory -Force -ErrorAction SilentlyContinue
                            }
                            foreach ($File in $FoundFiles) {
                                Copy-Item -Path $File.FullName -Destination $DestinationPath -Force -ErrorAction SilentlyContinue
                            }
                        }
                    }
                    "StopProcess" {
                        $Step.Name | ForEach-Object {
                            Get-Process -Name $_ -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                        }
                    }
                    "Remove" {
                        $RemovePath = Invoke-Expression "`"$($Step.Path)`""
                        if (Test-Path $RemovePath) { Remove-Item -Path $RemovePath -Recurse -Force -ErrorAction SilentlyContinue }
                    }
                    "StartProcess" {
                        Start-Process -FilePath $Step.FilePath -ArgumentList $Step.Arguments -Wait -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        # Special Cases
        switch -Wildcard ($SoftwareName) {
            "mas" { 
                $OsMajor = [System.Environment]::OSVersion.Version.Major
                Write-Host "You have Windows $OsMajor"
                $Choice = Read-Host "Do you really want to download and execute the remote script? (Y/N)"
                if ($Choice -match "^[yY]$") {
                    Write-Host "Starting script..." -ForegroundColor Cyan
                    if ($OsMajor -ge 10) {
                        Write-Host "KMS38 windows activation and Office Ohook activation"
                        & ([ScriptBlock]::Create((Invoke-RestMethod $env:DOWNSTALL_MAS_URL))) /KMS38 /Ohook
                    }
                    else {
                        Write-Host "KMS windows activation only"
                        & ([ScriptBlock]::Create((Invoke-RestMethod $env:DOWNSTALL_MAS_URL))) /KMS-WindowsOffice /KMS-RenewalTask
                    }
                }
                else {
                    Write-Host "Execution canceled by the user." -ForegroundColor Yellow
                }
            }
            "office365" {
                $SetupPath = $DownloadDir | Split-Path -Parent | Split-Path -Parent 
                Start-Process -FilePath (Join-Path $SetupPath "setup.exe") -ArgumentList "/configure `"$SetupPath\config.xml`"" -Wait
            }
            "office*" {
                @("Word.lnk", "Excel.lnk", "Powerpoint.lnk") | ForEach-Object {
                    $ShortcutPath = Join-Path $CommonPrograms $_
                    if (Test-Path $ShortcutPath) { Copy-Item -Path $ShortcutPath -Destination $CommonDesktop -Force }
                }
            }
            "true_image" {
                $Null = Set-Service afcdpsrv, syncagentsrv -StartupType Disabled -PassThru -Confirm:$false -ErrorAction SilentlyContinue | Stop-Service
            }
            "total_commander" {
                # Check native x64 path first, fallback to x86 path
                $Dest = Join-Path $ProgramFiles "Total Commander"
                if (-not (Test-Path $Dest)) {
                    $Dest = Join-Path $ProgramFilesX86 "Total Commander"
                }
                $Null = Get-ChildItem -Path $DownloadDir -Include wincmd.key, TOTALCMD*.EXE -Recurse -Force -ErrorAction Ignore | Copy-Item -Destination $Dest -Force
            }
            "viber" {
                $Retry = 0
                while ((Get-Process | Where-Object Path -match "vibersetup" -ErrorAction SilentlyContinue) -and $Retry -lt 60) { Start-Sleep -Seconds 1; $Retry++ }
                Get-Process | Where-Object Path -match "Viber" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            }
            "imagine_plugin_*" {
                $PluginDir = Join-Path $LocalAppData "Imagine\Plugin"
                if (-not (Test-Path $PluginDir)) {
                    $Null = New-Item -ItemType Directory -Path $PluginDir -Force -ErrorAction SilentlyContinue
                }
                $Null = Expand-Archive -Path $FullFilePath -DestinationPath $PluginDir -Force -ErrorAction SilentlyContinue
            }
            "ventoy_linux" { Write-Warning "$SoftwareName is just for Linux :)" }
            "platelschik_eaes_rate" {
                $EaesPath = "${Env:ProgramFiles(x86)}\МНС\Плательщик ЕАЭС*\description"
                $SetupDir = $Null
                if (Test-Path $EaesPath) { $SetupDir = (Get-Item -Path $EaesPath -Force | Select-Object -First 1).FullName }
                else {
                    $SearchDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -ne "C:\" } | Select-Object -ExpandProperty Root
                    foreach ($Drive in $SearchDrives) {
                        if ($Found = Get-ChildItem -Path "$Drive*" -Include $SoftwareItem.SetupPattern -Recurse -Force -ErrorAction Ignore) {
                            $SetupDir = $Found | Select-Object -First 1 | Split-Path -Parent; break
                        }
                    }
                }
                if ($SetupDir) {
                    Remove-Item (Join-Path $SetupDir "reduced_rate_nds.stbl") -Force -ErrorAction Ignore
                    Remove-Item (Join-Path $SetupDir $SoftwareItem.SetupPattern) -Force -ErrorAction Ignore
                    Write-Host "Extracting Eaes rate update to $SetupDir..."
                    Expand-Archive -Path $FullFilePath -DestinationPath $SetupDir -Force -ErrorAction SilentlyContinue
                }
                else { Write-Warning "Can't find platelschik_eaes directory for rate update." }
            }
            "rdpwrap_ini" {
                $RdpPath = Join-Path $ProgramFiles "RDP Wrapper"
                $SetupDir = $Null
                if (Test-Path $RdpPath) { $SetupDir = $RdpPath }
                else {
                    $SearchDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -ne "C:\" } | Select-Object -ExpandProperty Root
                    foreach ($Drive in $SearchDrives) {
                        if ($Found = Get-ChildItem -Path "$Drive*" -Directory -Recurse -Force -ErrorAction Ignore | Where-Object { $_.Name -match "RDP Wrapper" }) {
                            $SetupDir = $Found.FullName | Select-Object -First 1; break
                        }
                    }
                }
                if ($SetupDir) {
                    Stop-Service TermService -Force -ErrorAction SilentlyContinue
                    Remove-Item (Join-Path $SetupDir $SoftwareItem.SearchPattern) -Force -ErrorAction SilentlyContinue
                    Move-Item -Path $FullFilePath -Destination $SetupDir -Force -ErrorAction SilentlyContinue
                    Start-Service TermService -ErrorAction SilentlyContinue
                }
                else { Write-Warning "Can't find RDPWrap directory." }
            }
        }
    }

    function Install-SoftwarePackage {
        param (
            [Parameter(Mandatory = $true)] [Object]$SoftwareItem,
            [Parameter(Mandatory = $true)] [String]$FilePath,
            [Parameter(Mandatory = $true)] [String]$FileName
        )
        $FullFilePath = Join-Path -Path $FilePath -ChildPath $FileName
        $SoftwareName = $SoftwareItem.SoftwareName
        Write-Warning "Install $SoftwareName"
        
        $CustomInstallOnly = @("imagine_plugin_*", "mas", "office365", "office*", "ventoy_linux", "platelschik_eaes_rate", "rdpwrap_ini")
        $BypassStandardInstall = $false
        foreach ($Pattern in $CustomInstallOnly) {
            if ($SoftwareName -like $Pattern) { $BypassStandardInstall = $true; break }
        }

        if (Test-Path (Join-Path $PSScriptRoot "downstall+.ps1")) {
            $BypassStandardInstall = & (Join-Path $PSScriptRoot "downstall+.ps1")
        }

        if (-not $BypassStandardInstall) {
            $IsArchive = $FileName -match "\.(zip|7z|rar|gz|tgz|tar)$"
            $ExtractSingleExe = [bool]$SoftwareItem.ExtractSingleExe

            if ($IsArchive) {
                # Determine file extension
                $Extension = [System.IO.Path]::GetExtension($FileName).ToLower()
                # Determine extraction destination
                $DestDir = if ($SoftwareItem.SetupPattern -or $ExtractSingleExe) {
                    $TempDir = Join-Path $TempPath "temp_downstall_$SoftwareName"
                    $Null = New-Item -ItemType Directory -Force -Path $TempDir
                    $TempDir
                }
                else {
                    $DesktopDir = Join-Path $CommonDesktop $SoftwareName
                    $Null = New-Item -ItemType Directory -Force -Path $DesktopDir
                    $DesktopDir
                }

                # 1. EXTRACT BASED ON FORMAT
                $ExtractSuccess = $false
                switch -Regex ($Extension) {
                    "\.zip$" { 
                        Write-Host "Extracting ZIP using built-in Expand-Archive..."
                        Expand-Archive -Path $FullFilePath -DestinationPath $DestDir -Force -ErrorAction SilentlyContinue; $ExtractSuccess = $true 
                    }
                    "\.(gz|tgz|tar)$" { 
                        Write-Host "Extracting GZ/TAR using built-in Windows tar.exe..."
                        Start-Process -FilePath "tar.exe" -ArgumentList "-xf `"$FullFilePath`" -C `"$DestDir`"" -Wait -NoNewWindow -ErrorAction SilentlyContinue; $ExtractSuccess = $true 
                    }
                    "\.(rar|7z)$" {
                        $ProgramFilesPaths = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ } | Select-Object -Unique
                        $WinRarPath = $ProgramFilesPaths | ForEach-Object { Join-Path $_ "WinRAR\Winrar.exe" } | Where-Object { Test-Path $_ } | Select-Object -First 1
                        $7zPath = $ProgramFilesPaths | ForEach-Object { Join-Path $_ "7-Zip\7z.exe" } | Where-Object { Test-Path $_ } | Select-Object -First 1

                        if (Test-Path $WinRarPath) {
                            $WinRarArgs = if ($SoftwareItem.ArchiveArgs) { "x -o+ $($SoftwareItem.ArchiveArgs)" } else { "x -o+" }
                            Start-Process -FilePath $WinRarPath -ArgumentList "$WinRarArgs `"$FullFilePath`" `"$DestDir`"" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
                            $ExtractSuccess = $true
                        }
                        elseif (Test-Path $7zPath) {
                            Start-Process -FilePath $7zPath -ArgumentList "x `"$FullFilePath`" -o`"$DestDir`" -y" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
                            $ExtractSuccess = $true
                        }
                        else { Write-Warning "Cannot extract '$Extension'. Please install WinRAR or 7-Zip first! (Windows 10 lacks native support)." }
                    }
                    default { Write-Warning "Unknown archive format: $Extension" }
                }

                # 2. POST-EXTRACTION ACTIONS
                if ($ExtractSuccess) {
                    if ($SoftwareItem.SetupPattern) {
                        $ExtractedItems = Get-ChildItem -Path "$DestDir\*" -Include $SoftwareItem.SetupPattern -Recurse -Force
                        if ($ExtractedItems) {
                            $TargetFile = $ExtractedItems | Select-Object -First 1
                            if ($ExtractSingleExe) {
                                # If only extraction is needed (e.g., Portable version to Desktop)
                                Copy-Item -Path $TargetFile.FullName -Destination $CommonDesktop -Force -ErrorAction SilentlyContinue
                                Write-Host "Copied $($TargetFile.Name) to Desktop."
                            }
                            else {
                                # Normal mode: run the installer
                                $LaunchArgs = if ($SoftwareItem.InstallArgs) { $SoftwareItem.InstallArgs } else { "" }
                                Start-Process -FilePath $TargetFile.FullName -ArgumentList $LaunchArgs -Wait -ErrorAction SilentlyContinue
                                # Specific action for Avest
                                if ($SoftwareName -eq "avest") {
                                    $AvestSetup = Get-ChildItem -Path "$DestDir\*" -Include "setupAvCSPBel*.exe" -Recurse -Force | Select-Object -First 1
                                    if ($AvestSetup) { Start-Process -FilePath $AvestSetup.FullName -ArgumentList "/verysilent /devices=avToken,avPass,iKey" -Wait -ErrorAction SilentlyContinue }
                                }
                            }
                        }
                        else { Write-Warning "Setup pattern '$($SoftwareItem.SetupPattern)' not found in extracted archive." }
                    }
                    elseif ($ExtractSingleExe) {
                        $FirstExe = Get-ChildItem $DestDir -Filter "*.exe" -Recurse | Select-Object -First 1
                        if ($FirstExe) { Copy-Item -Path $FirstExe.FullName -Destination $CommonDesktop -Force -ErrorAction SilentlyContinue }
                    }

                    if ($SoftwareItem.SetupPattern -or $ExtractSingleExe) {
                        if (Test-Path $DestDir) { Remove-Item $DestDir -Recurse -Force -ErrorAction SilentlyContinue }
                    }
                }
            }
            else {
                if ($ExtractSingleExe) {
                    if (Test-Path $FullFilePath) { 
                        Copy-Item -Path $FullFilePath -Destination $CommonDesktop -Force -ErrorAction SilentlyContinue 
                    }
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
                            Write-Warning "File '$FileName' is not a valid executable (Missing MZ header). Download was corrupted or blocked by bot-protection."
                            return 
                        }
                    }
                    $LaunchArgs = if ($SoftwareItem.InstallArgs) { $SoftwareItem.InstallArgs } else { "" }
                    Start-Process -FilePath $FullFilePath -ArgumentList $LaunchArgs -Wait -ErrorAction SilentlyContinue
                }
            }
        }
        Invoke-PostInstall -SoftwareItem $SoftwareItem -DownloadDir $FilePath -FileName $FileName
    }

    # ==========================================
    # URL & DEPLOYMENT LOGIC
    # ==========================================
    function Update-DynamicSoftwareUrls {
        param ([Array]$SoftwareArray, [Array]$ActiveInstalls)
        $IrfanViewVersion = $Null
        
        foreach ($Soft in $SoftwareArray) {
            if ($Soft.SoftwareName -notin $ActiveInstalls) { continue }
            switch ($Soft.SoftwareName) {
                { $_ -match "^irfanview" } {
                    if (-not $IrfanViewVersion) {
                        $HtmlContent = (Invoke-WebRequest -Uri $env:DOWNSTALL_IRFANVIEW_URL -UserAgent $Global:UserAgent).Content
                        $Match = [regex]::Match($HtmlContent, "(?i)version\s*([\d\.]+)")
                        if ($Match.Success) { $IrfanViewVersion = $Match.Groups[1].Value.Replace(".", "") }
                    }
                    $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#irfanviewVersion', $IrfanViewVersion)
                }
                "opera" {
                    $BaseUri = $env:DOWNSTALL_OPERA_URL
                    $Links = (Invoke-WebRequest -Uri $BaseUri -UseBasicParsing -UserAgent $Global:UserAgent).links.href
                    $LatestVer = @($Links | Where-Object { $_ -match "^\d+\.\d+\.\d+\.\d+/?$" }) | Sort-Object -Property { [version]($_ -replace '/', '') } | Select-Object -Last 1
                    $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#operaUrl', "$BaseUri$LatestVer" + "win/")
                }
                { $_ -match "^softethervpn" } {
                    $BaseUri = $env:DOWNSTALL_SOFTETHER_URL
                    $Links = (Invoke-WebRequest -Uri $BaseUri -UseBasicParsing -UserAgent $Global:UserAgent).links.href
                    # Filter RTM versions and sort them correctly by extracting the embedded date (YYYY.MM.DD)
                    $LatestNode = @($Links | Where-Object { $_ -match "v\d+\.\d+-\d+-rtm-(\d{4}\.\d{2}\.\d{2})-tree" }) | Sort-Object -Property { 
                        $DateStr = [regex]::Match($_, "rtm-(\d{4}\.\d{2}\.\d{2})-tree").Groups[1].Value
                        [datetime]::ParseExact($DateStr, "yyyy.MM.dd", $Null) 
                    } | Select-Object -Last 1

                    $RootUri = ([uri]$BaseUri).GetLeftPart([System.UriPartial]::Authority) + "/"
                    if ($Soft.SoftwareName -eq "softethervpn") { 
                        $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#softethervpnUrl', "$RootUri$LatestNode" + "Windows/SoftEther_VPN_Client/") 
                    }
                    else { 
                        $Soft.DownloadUrl = $Soft.DownloadUrl.Replace('#softethervpnServerUrl', "$RootUri$LatestNode" + "Windows/SoftEther_VPN_Server_and_VPN_Bridge/") 
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
                    $YandexApi = Invoke-RestMethod "$($env:DOWNSTALL_YANDEX_API_URL)?public_key=$UrlDecoded&limit=200" -UserAgent $Global:UserAgent -ErrorAction SilentlyContinue
                }
                else {
                    $UrlDecoded = [System.Web.HttpUtility]::UrlEncode($DownloadUrl.Substring(0, $EndIndex))
                    $YandexPath = $DownloadUrl.Substring($EndIndex)
                    $YandexApi = Invoke-RestMethod "$($env:DOWNSTALL_YANDEX_API_URL)?public_key=$UrlDecoded&path=$YandexPath&limit=200" -UserAgent $Global:UserAgent -ErrorAction SilentlyContinue
                }
                if ($YandexApi.type -eq "file") {
                    $DownloadLink = (Invoke-RestMethod "$($env:DOWNSTALL_YANDEX_API_DL)?public_key=$UrlDecoded" -UserAgent $Global:UserAgent).href
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
                $DownloadLink = "$($env:DOWNSTALL_GOOGLE_DRIVE_API)$DriveId"
            }
            elseif ($DownloadUrl -match "^https?://github\.com/([^/]+)/([^/]+)") {
                $RepoOwner = $Matches[1]
                $RepoName = $Matches[2].Replace(".git", "")
                $ApiUrl = "$($env:DOWNSTALL_GITHUB_API_BASE)/repos/$RepoOwner/$RepoName/releases/latest"
                $ApiHeaders = @{ Accept = "application/vnd.github.v3+json" }
                if ($env:GITHUB_TOKEN) {
                    $ApiHeaders.Add("Authorization", "token $($env:GITHUB_TOKEN)")
                }
                $ReleaseData = Invoke-RestMethod -Uri $ApiUrl -Headers $ApiHeaders -ErrorAction SilentlyContinue
                if ($ReleaseData.assets) {
                    $TargetAsset = $ReleaseData.assets | Where-Object { $_.name -like $SoftwareItem.SearchPattern } | Select-Object -First 1
                    if ($TargetAsset) {
                        $DownloadLink = $TargetAsset.browser_download_url
                        $LinkFileName = $TargetAsset.name
                    }
                }
            }
            else {
                # Force strictly to string to prevent array injection
                if (-not ([string]$AbsUrl = Get-AbsoluteUri -DownloadUrl $DownloadUrl)) { return }
                $DownloadUrl = $AbsUrl
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
                    $DownloadLink = $Null
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
            
            if ($SoftwareItem.SoftwareName -eq "platelschik_eaes" -and $DownloadLink) {
                $DownloadLink = $DownloadLink.Replace(" ", "%20" )
            }

            # Download Logic
            if ($LinkFileName -and ($CurrentFileName -ne $LinkFileName -or (Test-FileUpdateRequired -LocalFilePath "$FilePath\$CurrentFileName" -DownloadUrl $DownloadLink))) {
                Write-Warning "+$LinkFileName"
                if (-not (Test-Path $FilePath)) { $Null = New-Item $FilePath -ItemType Directory }
                
                try {
                    Start-DownloadWithProgress -Uri $DownloadLink -OutFile (Join-Path $FilePath $LinkFileName) -SoftwareName $SoftwareItem.SoftwareName
                }
                catch {
                    Write-Warning "$CurrentFileName - Error: $($_.Exception.Message)"
                    $LinkFileName = $Null # Reset so it doesn't try to install a failed download
                }
                
                if ($LinkFileName -and $CurrentFileName -and ($CurrentFileName -ne $LinkFileName)) {
                    Write-Warning "-$CurrentFileName deleted!"
                    Remove-Item -Path (Join-Path $FilePath $CurrentFileName) -Force -Recurse -ErrorAction SilentlyContinue
                }
            }
        }

        # Install Phase (With protection against empty filenames)
        if (-not $DownloadOnly) {
            if ([string]::IsNullOrWhiteSpace($LinkFileName)) {
                Write-Warning "Skipping installation of '$($SoftwareItem.SoftwareName)': File not found locally and download failed/unavailable."
            }
            else {
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
    # Disable default progress bars for all native commands
    $ProgressPreference = "SilentlyContinue"
    # OS-level resolution of special folders
    $CommonPrograms = [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonPrograms)
    $CommonDesktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonDesktopDirectory)
    $ProgramFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
    $ProgramFilesX86 = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86)
    $LocalAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    $TempPath = [System.IO.Path]::GetTempPath()

    if ($DownloadOnly -and $InstallOnly) {
        Write-Warning "ok, do nothing!"
        exit
    }
    
    $CheckDomain = if ($env:DOWNSTALL_CHECK_DOMAIN) { $env:DOWNSTALL_CHECK_DOMAIN } else { "google.com" }
    while (-not $InstallOnly -and -not ([System.Net.Dns]::GetHostAddresses($CheckDomain))) {
        $InstallOnly = $true
        Write-Warning "Internet is missing, -InstallOnly activated"
    }
    
    $Files = "soft.json", "soft+.json"
    $SoftwareList = foreach ($FileName in $Files) {
        $FilePath = Join-Path $PSScriptRoot $FileName
        if (Test-Path $FilePath) {
            Get-Content -Path $FilePath -Raw | ConvertFrom-Json
        }
        elseif ($FileName -eq "soft.json") { 
            Write-Warning "$FileName is missing" 
        }
    }
    
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
