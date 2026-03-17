#Requires -RunAsAdministrator
#Requires -Version 5.1   # Windows 10 LTSC 

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [Switch]$downloadOnly,
  
  [Parameter(Mandatory = $false)]
  [Switch]$installOnly
)
DynamicParam {
  $ParameterName = 'install'
  $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
  $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
  $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
  $ParameterAttribute.Mandatory = $false
  $ParameterAttribute.Position = 1
  $AttributeCollection.Add($ParameterAttribute)
  if (Test-Path "$PSScriptRoot\soft.json") {
    $arrSet = (Get-Content -Path "$PSScriptRoot\soft.json" | ConvertFrom-JSON).name
  }
  if (Test-Path "$PSScriptRoot\soft+.json") {
    $arrSet += (Get-Content -Path "$PSScriptRoot\soft+.json" | ConvertFrom-JSON).name
  }
  $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
  $AttributeCollection.Add($ValidateSetAttribute)
  $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string[]], $AttributeCollection)
  $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
  return $RuntimeParameterDictionary
}
begin {
  [string[]]$install = $PsBoundParameters[$ParameterName]
}

process {
  $Host.UI.RawUI.WindowTitle = "Downstall script 4 Windows $([char]0x00A9) poljik 2019-2026"

  # enable TLSv1.2 for compatibility with older clients
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
  # increase speed of downloading
  $progresspreference = "SilentlyContinue"

  if ($downloadOnly -and $installOnly) {
    Write-Warning "ok, do nothing!"
    exit
  }
  
  # check internet
  while (-not ($installOnly) -and -not (Test-Connection -Count 1 -computer 8.8.8.8 -quiet)) {
    $installOnly = $true
    Write-Warning -Message "Internet is missing, --installOnly activated"
  }

  $soft = @()
  if (Test-Path "$PSScriptRoot\soft.json") {
    $soft = Get-Content -Path "$PSScriptRoot\soft.json" | ConvertFrom-JSON
  }
  else {
    Write-Warning "Soft.json is missing"
  }

  if (Test-Path "$PSScriptRoot\soft+.json") {
    $soft += Get-Content -Path "$PSScriptRoot\soft+.json" | ConvertFrom-JSON
  }

  if (-not $install) {
    $downloadOnly = $true
    $install = $soft.Name
  }

  # get windows version
  function Get-OSVersion {
    if (-not $IsLinux) {
      $build = [System.Environment]::OSVersion.Version.Build
      $productName = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
      if ($productName -match 'Server') {
        if ($build -ge 26100) { return "2025" }
        elseif ($build -ge 20348) { return "2022" }
        elseif ($build -ge 17763) { return "2019" }
        elseif ($build -ge 14393) { return "2016" }
        elseif ($build -eq 9600) { return "2012" } # 2012 R2
        elseif ($build -eq 9200) { return "2012" }
        elseif ($build -ge 7600) { return "2008" } # 2008 R2
      }
      else {
        if ($build -ge 22000) { return "11" }
        elseif ($build -ge 10240) { return "10" }
        elseif ($build -eq 9600) { return "8.1" }
        elseif ($build -eq 9200) { return "8" }
        elseif ($build -ge 7600) { return "7" }
      }
    }
  }

  $winVersion = Get-OSVersion

  # get bit OS
  if ([System.Environment]::Is64BitOperatingSystem) {
    $script:bitOS = 64
  }
  else {
    $script:bitOS = 86
  }

  function Find-Dir {
    param (
      [String]$pattern, [String]$dir
    )
    $file = Get-ChildItem -Path "$PSScriptRoot" -Include $pattern -Recurse -Force -ErrorAction Ignore
    if ($file.Count) {
      $fileName = $file[0].Name
      if ($file[0].Attributes -eq 'Directory') {
        $filePath = $file[0].FullName
      }
      else {
        $filePath = $file[0].DirectoryName
      }
    }
    else {
      $dirs = Get-ChildItem -Path "$PSScriptRoot" -Include $dir -Recurse -Force  -ErrorAction Ignore
      if ($dirs.Count) {
        $filePath = $dirs[0].FullName
      }
      else {
        $filePath = "$PSScriptRoot" + "/" + $dir.Replace("*", " ")
      }
    }
    return $filePath, $fileName
  }
  function Compare-Sizes {
    if ($fileName -eq $linkName) {
      $oldSize = (Get-ChildItem -Path $filePath\$fileName).Length
      if ($PSVersionTable.PSVersion.Major -ge 6) {
        try {
          $newSize = (Invoke-WebRequest $link -Headers @{Referer = $link } -Method Head -SkipCertificateCheck).Headers.'Content-Length' 
        }
        catch {
          $newSize = (Invoke-WebRequest $link -Headers @{Referer = $link } -Method Head).Headers.'Content-Length'
        }
      }
      else {
        $newSize = (Invoke-WebRequest $link -Headers @{Referer = $link } -Method Head).Headers.'Content-Length'        
      }
      if ($null -eq $newSize) {
        Write-Host "The new size is unknown, skip download $linkName"
        return $false
      }
      else {
        if ($newSize.GetType().Name -eq "String[]") {
          $newSize = [Int]$newSize[0]
        }
        else {
          $newSize = [Int]$newSize
        }
        if (($oldSize -eq $newSize) -or ($newSize -eq 0)) {
          Write-Host "Skip download $linkName"
          return $false
        } 
        # different sizes
        return $true
      }
    }
  }
  function Get-AbsoluteUri {
    if ($PSVersionTable.PSVersion.Major -ge 6) {
      try {
        $url0 = Invoke-WebRequest -Uri $url -Method Head -SkipCertificateCheck
      }
      catch {
        Write-Warning "$fileName - Error: $($_.Exception.Message)"
        try {
          $url0 = Invoke-WebRequest -Uri $url
        }
        catch {
          Write-Warning "$fileName - Error: $($_.Exception.Message)"
        }
      }
    }
    else {
      try {
        $url0 = Invoke-WebRequest -Uri $url
      }  
      catch {
        Write-Warning "$fileName - Error: $($_.Exception.Message)"
      }
    }
    $url = $url0.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
    if (-not $url) {
      $url = $url0.BaseResponse.ResponseUri.AbsoluteUri
    }
    return $url
  }
  function Install-Prog {
    Write-Warning "Install $name"
    # custom installation
    switch -Wildcard ($name) {
      "imagine_plugin_*" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          if (Test-Path $Env:LOCALAPPDATA\Imagine\Plugin) {
            Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:LOCALAPPDATA\Imagine\Plugin -WindowStyle Hidden
          }
        }
      }

      "mas" {
        Write-Host "You have Windows $winVersion"
        if ($winVersion -ge 10) {
          Write-Host "KMS38 windows activation and Office Ohook activation"
          & ([ScriptBlock]::Create((Invoke-RestMethod https://massgrave.dev/get))) /KMS38 /Ohook
        }
        else {
          Write-Host "KMS windows activation only"
          & ([ScriptBlock]::Create((Invoke-RestMethod https://massgrave.dev/get))) /KMS-WindowsOffice /KMS-RenewalTask
        }
      }

      "office365" {
        $setupPath = $filePath | Split-Path -Parent | Split-Path -Parent
        Start-Process -FilePath $setupPath\setup.exe -ArgumentList "/configure `"$setupPath\config.xml`"" -Wait
        # Remove-Item -Path $setupPath\config.xml
      }

      "office*" {
        if (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk") {
          Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" -Destination $Env:PUBLIC\Desktop -Force
        }
        if (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk") {
          Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" -Destination $Env:PUBLIC\Desktop -Force
        }
        if (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk") {
          Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk" -Destination $Env:PUBLIC\Desktop -Force
        }
      }

      "ventoy_linux" {
        Write-Warning "$name is just for Linux :)"
      }

      "platelschik_eaes_rate" {
        $pathPlatelschik = "${Env:ProgramFiles(x86)}\МНС\Плательщик ЕАЭС*\description"
        if (Test-Path ($pathPlatelschik)) {
          $fileSetup = Get-ChildItem -Path $pathPlatelschik -Force
        }
        elseif (($fileSetup = Get-ChildItem -Path D:\* -Include $patternSetup -Recurse -Force)) {
          $fileSetup = $fileSetup | Split-Path -Parent
        }
        else {
          Write-Warning "Can't find platelschik_eaes directory"
          break
        }
        $fileSetup.ToString()
        Remove-Item $fileSetup\reduced_rate_nds.stbl -Force -ErrorAction Ignore
        Remove-Item $fileSetup\$patternSetup -Force
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", `"$fileSetup`" -WindowStyle Hidden
        }
      }

      "rdpwrap_ini" {
        $pathPlatelschik = "$Env:ProgramFiles\RDP Wrapper\"
        if (Test-Path ($pathPlatelschik)) {
          $fileSetup = $pathPlatelschik
        }
        elseif (($fileSetup = Get-ChildItem -Path D:\*  -Directory -Recurse -Force | Where-Object { $_.name -match "RDP Wrapper" })) {
          $fileSetup = $fileSetup.FullName
          $fileSetup.ToString()
        }
        else {
          Write-Warning "Can't find RDPWrap directory"
          break
        }
        Stop-Service TermService -Force
        Remove-Item $fileSetup\$pattern -Force
        Write-Warning "Move new rdpwrap.ini to $fileSetup"
        Move-Item -Path $filePath\$linkName -Destination $fileSetup -Force
        Start-Service TermService
      }

      # main installation
      default {
        Clear-Variable pass -ErrorAction SilentlyContinue
        if (Test-Path $PSScriptRoot/downstall+.ps1) {
          $pass = Invoke-Expression -Command "$PSScriptRoot/downstall+.ps1"
        }
        if (-not $pass) {
          if (($linkName -like "*.zip") -or ($linkName -like "*.7z") -or ($linkName -like "*.rar") -or ($linkName -like "*.gz")) {
            if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
              # unpack one portable exe to desktop
              if ($oneExe2Desktop) {
                if ($patternSetup) {
                  Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "e -o+", `"$filePath\$linkName`", $patternSetup, $Env:PUBLIC\Desktop -WindowStyle Hidden -Wait
                }
                else {
                  Write-Warning "You need add PatternSetup like '*$name.exe' to Json file"
                }
              }
              # unpack to temp and start setup with(out) args
              elseif ($patternSetup) {
                New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
                if ($RarArg) {
                  Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+ $RarArg", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
                }
                else {
                  Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
                }
                $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Recurse -Force
                if ($arg) {
                  Start-Process -FilePath $fileSetup -ArgumentList $arg -Wait # -WindowStyle Hidden    
                }
                else {
                  Start-Process -FilePath $fileSetup -Wait
                }
                if ($name -eq "avest") {
                  $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include "setupAvCSPBel*.exe" -Recurse -Force
                  Start-Process -FilePath $fileSetup -ArgumentList "/verysilent /devices=avToken,avPass,iKey" -Wait
                }
                Remove-Item $Env:Temp\temp1 -Recurse -Force
              }
              # just unpack on desktop
              else {
                New-Item -ItemType Directory -Force -path $Env:PUBLIC\Desktop\$name > $Null
                Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:PUBLIC\Desktop\$name -WindowStyle Hidden -Wait
              }
            }
            else {
              Write-Warning "You need install WinRar first!"
            }
          }
          elseif ($oneExe2Desktop) {
            if (Test-Path $filePath\$linkName) {
              Copy-Item -Path $filePath\$linkName -Destination $Env:PUBLIC\Desktop -Force
            }
            else {
              Write-Warning "You must download $name first!"
            }
          }
          # start setup with(out) args
          else {
            if ($arg) {
              Start-Process -FilePath $filePath\$linkName -ArgumentList $arg -Wait # -WindowStyle Hidden
            }
            else {
              Start-Process -FilePath $filePath\$linkName
            }
          }
        }

        # do after installation
        switch ($name) {
          "true_image" {
            # disable services Acronis Nonstop Backup Service, Acronis Sync Agent Service
            Set-Service afcdpsrv -StartupType Disabled -PassThru -Confirm:$false -ErrorAction SilentlyContinue | Stop-Service -PassThru > $Null
            Set-Service syncagentsrv -StartupType Disabled -PassThru -Confirm:$false -ErrorAction SilentlyContinue | Stop-Service -PassThru > $Null
          }

          "total_commander" {
            if ($bitOS -eq 64) {
              $copy_to = ${Env:Programfiles(x86)}
            }
            else {
              $copy_to = $Env:Programfiles
            }
            $copy_to += "\Total Commander"
    
            if ($file2copy = Get-ChildItem -Path $filePath -Include wincmd.key -Recurse -Force -ErrorAction Ignore) {
              Copy-Item -Path $file2copy -Destination $copy_to -Force
            }
    
            if ($file2copy = Get-ChildItem -Path $filePath -Include TOTALCMD*.EXE -Recurse -Force -ErrorAction Ignore) {
              Copy-Item -Path $file2copy -Destination $copy_to -Force
            }
          }

          "far" {
            if ($file2copy = Get-ChildItem -Path $filePath -Include Panel.Esc.lua -Recurse -Force -ErrorAction Ignore) {
              "$Env:APPDATA\Far Manager\Profile\Macros\scripts\" | ForEach-Object {
                New-Item -Path $_ -ItemType Directory -Force > $Null
                Copy-Item -Path $file2copy -Destination $_ -Force
              }
            }
            if ($file2copy = Get-ChildItem -Path $filePath -Include generalconfig.db -Recurse -Force -ErrorAction Ignore) {
              "$Env:APPDATA\Far Manager\Profile\" | ForEach-Object {
                New-Item -Path $_ -ItemType Directory -Force > $Null
                Copy-Item -Path $file2copy -Destination $_ -Force
              }
            }
            $bit = ""
            if ($bitOS -eq 64) {
              $bit = " (x$bitOS)"
            }
            Do {
              Start-Sleep -Seconds 1
            } While (-not (Get-ChildItem -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Far Manager 3$bit\Far Manager 3$bit.lnk" -ErrorAction Ignore))
            Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Far Manager 3$bit\Far Manager 3$bit.lnk" -Destination $Env:PUBLIC\Desktop -Force
          }

          "adobe_reader" {
            Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Acrobat.lnk" -Destination $Env:PUBLIC\Desktop -Force
            # delete button-banner "Try Acrobat Pro DC"
            "HKLM:\Software\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" | ForEach-Object {
              New-Item -Path $_ -ItemType Directory -Force > $Null
              Set-ItemProperty -Path $_ -Name bAcroSuppressUpsell -Value 1 -Force
            }
          }

          "viber" {
            Do {
              Start-Sleep -Seconds 1
            } While (Get-Process | Where-Object { $_.Path -like "*vibersetup*" })
            Write-Host "Kill Viber autostart"
            Get-Process | Where-Object { $_.Path -like "*Viber*" } | Stop-Process -Force -processname { $_.ProcessName }
          }

          "winrar" {
            if ((Test-Path $Env:Programfiles\WinRAR\Winrar.exe) -and `
              ($file2copy = Get-ChildItem -Path $filePath -Include "rarreg.key" -Recurse -Force -ErrorAction Ignore)) {
              Copy-Item $file2copy $Env:Programfiles\WinRAR
              Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\WinRAR\WinRAR.lnk" -Destination $Env:PUBLIC\Desktop -Force
            }
          }

          "edeclaration" {
            Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\EDeclaration\Запустить EDeclaration.lnk" -Destination $Env:PUBLIC\Desktop -Force
            $path2file = Get-ChildItem -Path $Env:PUBLIC\Desktop -Include "Запустить EDeclaration.lnk" -Recurse -Force -ErrorAction Ignore
            Rename-Item -Path $path2file -NewName "EDeclaration.lnk" -Force
          }

          "platelschik_eaes" {
            Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\МНС\Плательщик ЕАЭС*\МНС\Плательщик ЕАЭС*.lnk" -Destination $Env:PUBLIC\Desktop -Force
            $path2file = Get-ChildItem -Path $Env:PUBLIC\Desktop -Include "Плательщик ЕАЭС*.lnk" -Recurse -Force -ErrorAction Ignore
            Rename-Item -Path $path2file -NewName "Плательщик ЕАЭС.lnk" -Force
          }
        }
      }
    }
  }
  function Get-File {
    param (
      [String]$pattern, [String]$dir, [String]$name, [String]$url, [String]$arg, [string]$RarArg, [string]$patternSetup, [System.Boolean]$oneExe2Desktop
    )
    Clear-Variable fileName -ErrorAction SilentlyContinue
    $filePath, $fileName = Find-Dir $pattern $dir

    if ($name -eq "office365") { 
      Get-Office365
      $filePath, $fileName = Find-Dir $pattern $dir
    }

    if ($url -and (-not $installOnly)) {
      # for yandex disk links
      if (($url -like '*.yandex.*') -or ($url -like '*yadi.sk*')) {
        $searchString = "/d/"
        $beginIndex = $url.IndexOf($searchString) + $searchString.Length
        $endIndex = $url.IndexOf("/", $beginIndex)
        # check for path
        if ($endIndex -eq -1) {
          # no path
          $ud = [System.Web.HttpUtility]::UrlEncode($url)
          try {
            $link = Invoke-RestMethod "https://cloud-api.yandex.net/v1/disk/public/resources?public_key=$ud&limit=200"
          }
          catch {
            Write-Warning "$fileName - Error: $($_.Exception.Message)"
          }
        }
        else {
          $ud = [System.Web.HttpUtility]::UrlEncode($url.Substring(0, $endIndex))
          $path = $url.Substring($endIndex, $url.Length - $endIndex)
          $link = Invoke-RestMethod "https://cloud-api.yandex.net/v1/disk/public/resources?public_key=$ud&path=$path&limit=200"
        }
        if ($link.type -eq "file") {
          $link = (Invoke-RestMethod "https://cloud-api.yandex.net:443/v1/disk/public/resources/download?public_key=$ud").href
          $searchString = "&filename="
          $beginIndex = $link.IndexOf($searchString) + $searchString.Length
          $endIndex = $link.IndexOf("&", $beginIndex)
          $linkName = $link.Substring($beginIndex, ($endIndex - $beginIndex))
        }
        else {
          $link = $link._embedded.items | Where-Object { $_.name -like $pattern } | Sort-Object -Property created
          if ($link) {
            $linkName = $link[-1].name
            $link = $link[-1].file
          }
        }
      }

      # for google drive links
      elseif ($url -like "*drive.google.com/file/d*") {
        try {
          $result = Invoke-WebRequest $url -UseBasicParsing -MaximumRedirection 0 -ErrorAction SilentlyContinue
        }
        catch {
          Write-Warning "$fileName - Error: $($_.Exception.Message)"
        }
        [Array]$links = $result.links.href | Where-Object { $_ -like "*drive.google.com/file/d*" }
        $linkID = $links[0] | Split-Path -Parent | Split-Path -Leaf

        # get filename
        $searchString = "meta property=""og:title"" content="""
        $beginIndex = $result.Content.IndexOf($searchString) + $searchString.Length
        $endIndex = $result.Content.IndexOf("""", $beginIndex)
        $linkName = $result.Content.Substring($beginIndex, ($endIndex - $beginIndex))
        $link = "https://drive.google.com/uc?export=download&id=$linkID"

        # # download Warning html into _tmp.txt
        # try {
        #   Invoke-WebRequest -Uri $link -OutFile "_tmp.txt" -SessionVariable googleDriveSession -ErrorAction SilentlyContinue
        # } catch {
        #   Write-Error $_
        # }
        # # get at code from _tmp.txt
        # $string = Select-String -Path "_tmp.txt" -Pattern 'name="at" value="'
        # $begin = $string.LastIndexOf('name="at" value="') + 17
        # $end = $string.LastIndexOf('"></form></div>')
        # $atCode = "$string.Substring($begin, $end)"
        # Remove-Item "_tmp.txt" -Force
        # $link = "https://drive.google.com/uc?export=download&confirm=t&id=$linkID&at=$atCode"
      }
      else {
        # github links
        $url = (Get-AbsoluteUri) -replace "/tag/", "/expanded_assets/"
        
        # for other links
        $link = $url
        if ($PSVersionTable.PSVersion.Major -ge 6) {
          try {
            $result = Invoke-WebRequest -Uri $url -UseBasicParsing -SkipCertificateCheck -DisableKeepAlive
          }
          catch {
            Write-Warning "$fileName - Error: $($_.Exception.Message)"
            try {
              $result = Invoke-WebRequest -Uri $url -UseBasicParsing -DisableKeepAlive
            }
            catch {
              Write-Warning "$fileName - Error: $($_.Exception.Message)"
            }
          }
        }
        else {
          try {
            $result = Invoke-WebRequest -Uri $url -UseBasicParsing -DisableKeepAlive
          }
          catch {
            Write-Warning "$fileName - Error: $($_.Exception.Message)"
          }
        }

        # imagine exception
        if ($name -eq "imagine") {
          $pattern = "*download.php?arch=x64&unicode=1&full=0&setup=1*"
        }

        [Array]$links = $result.links.href | where-Object { $_ -like $pattern }
        
        # if there are more then one link
        if ($links.Count) {
          $link = $links[0]
          
          # vvoddpu exception
          if ($name -eq "vvoddpu") {
            if ($links[0] -like "*0.zip") {
              $link = $links[1]
            }
          }

          # victoria exception
          if ($name -eq "victoria") {
            $begin = $links[1].LastIndexOf("https:")
            $link = $links[1].Substring($begin, ($links[1].Length - $begin))
          }
          
          if (-not ($link.Contains("http"))) {
            if ($link.Contains("/")) {
              if ($link[0] -eq "/") {
                # if $link starts with / , $url ends without /
                $link = $url.Substring(0, ($url.IndexOf("/", 8))) + $link
              }
              else {
                # if $link not starts with / , $url ends with /
                $link = $url.Substring(0, ($url.IndexOf("/", 8) + 1)) + $link
              }
            }
            else {
              $link = $url + $link
            }
          }
        }
        elseif ($link -like '*//github.com/*') {
          Clear-Variable -Name link
        }
        
        if ($link) {
          $linkName = $link | Split-Path -Leaf
          if ($linkName.Contains("?")) {
            $linkName = $linkName.Substring(0, $linkName.IndexOf("?"))
          }
          # replace spaces to %20
          $linkName = $linkName.ToString() -replace "%20", " "
        }
        
        # imagine exception
        if ($name -eq "imagine") {
          $begin = $links[0].LastIndexOf("version=") + 8
          $version = $links[0].Substring($begin, ($links[0].Length - $begin))
          $linkName = "Imagine_" + $version + "_x64_Unicode.exe"
          # $filePath, $fileName = Find-Dir $linkName $dir
        }
        
        # cloudflare warp exception
        if ($name -eq "cloudflare_warp") {
          $linkName = "cloudflare_warp_" + $linkName + ".msi"
        }
      }

      # for platelschik_eaes convert cyrillic symbols in link
      if ($name -eq "platelschik_eaes") {
        $link = [uri]::EscapeUriString($link)
      }

      # if new filename or filesize
      if (($linkName) -and (($fileName -ne $linkName) -or (Compare-Sizes))) {

        Write-Warning "+$linkName"
        # create directory if not
        if (-not (Test-Path $filePath)) {
          New-Item $filePath -ItemType Directory > $Null
        }

        # download new file
        try {
          Invoke-WebRequest -Uri $link -OutFile $filePath\$linkName -Headers @{Referer = $link }
        }
        catch {
          Write-Warning "$fileName - Error: $($_.Exception.Message)"
          return
        }

        # if download new filename, remove old file
        if (($fileName) -and ($fileName -ne $linkName)) {
          Write-Warning "-$fileName deleted!"
          Remove-Item -Path $filePath\$fileName -Force -Recurse
        }
      }
    }
    else {
      $linkName = $fileName
    }

    if (-not $downloadOnly) {
      Install-Prog
    }
  }
  function Get-Office365 {
    # original is here https://github.com/farag2/Office/blob/master/Download.ps1
    [xml]$Config = @"
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
    $setupPath = $filePath | Split-Path -Parent | Split-Path -Parent
    if (-not (Test-Path $setupPath\setup.exe)) {
      "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" | ForEach-Object {
        New-Item -Path $_ -ItemType Directory -Force > $Null
        Set-ItemProperty -Path $_ -Name CountryCode -Value "std::wstring|US" -Force
      }
      
      if (-not (Test-Path $filePath)) {
        New-Item $filePath -ItemType Directory > $Null
      }
      Push-Location
      Set-Location $filePath

      $result = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing
      $url = ($result.Links | Where-Object { $_.outerHTML -match "officedeploymenttoll" }).href
      Invoke-WebRequest -Uri $url -OutFile officedeploymenttool.exe -UseBasicParsing -Verbose

      Start-Process -FilePath officedeploymenttool.exe -ArgumentList "/quiet /extract:officedeploymenttool" -Wait

      Move-Item -Path officedeploymenttool\setup.exe -Destination ./ -Force
      Start-Sleep -Seconds 1
      Remove-Item -Path officedeploymenttool* -Recurse -Force
      $config.Save("$filePath\config.xml")
      Start-Process -FilePath setup.exe -ArgumentList "/download `"config.xml`"" -Wait
      Pop-Location
    }
  }
  function Get-lastVersionIrfanView {
    # get last version number of IrfanView
    $html = (Invoke-WebRequest -Uri https://www.irfanview.com/).Content
    $searchStr = "<span>Version "
    $begin = $html.IndexOf($searchStr)
    $html = $html.Substring($begin + $searchStr.Length, $html.Length - $begin - $searchStr.Length)
    $end = $html.IndexOf("</span>")
    return $html.Substring(0, $end).Replace(".", "").Replace(" ", "")
  }

  if (-not $installOnly) {
    $irfanviewVersion = Get-lastVersionIrfanView
  }

  $install | ForEach-Object {
    if (-not $installOnly) {
      switch ($_) {
        "irfanview" {
          foreach ($item in $soft) {
            if ($item.Name -eq $_) {
              $item.Url = $item.Url.Replace('#irfanviewVersion', $irfanviewVersion)
            }
          }
          break
        }

        "irfanview_plugins" {
          foreach ($item in $soft) {
            if ($item.Name -eq $_) {
              $item.Url = $item.Url.Replace('#irfanviewVersion', $irfanviewVersion)
            }
          }
          break
        }

        "notepad++" {
          # get notepad++ url
          $Parameters = @{
            Uri             = 'https://notepad-plus-plus.org/downloads/'
            UseBasicParsing = $true
          }
          $notepadppUrl = ((Invoke-WebRequest @Parameters).links.href |
            Where-Object { $_ -like $Parameters.Uri + '*' })[0]
          foreach ($item in $soft) {
            if ($item.Name -eq $_) {
              $item.Url = $item.Url.Replace('#notepadppUrl', $notepadppUrl)
            }
          }
          break
        }

        "opera" {
          # get last opera x64 link
          $Parameters = @{
            Uri             = 'https://get.geo.opera.com/pub/opera/desktop/'
            UseBasicParsing = $true
          }
          $operaUrl = $Parameters.Uri + (((Invoke-WebRequest @Parameters).links.href | Where-Object { $_ -like "1??.*" })[-1]) + 'win/'
          foreach ($item in $soft) {
            if ($item.Name -eq $_) {
              $item.Url = $item.Url.Replace('#operaUrl', $operaUrl)
            }
          }
          break
        }

        "softethervpn" {
          # get softether vpn client x64 link
          $Uri = 'http://www.softether-download.com'
          $Parameters = @{
            Uri             = $Uri + '/files/softether/'
            UseBasicParsing = $true
          }
          $softethervpnUrl = $Uri + ((Invoke-WebRequest @Parameters).links.href |
            Where-Object { $_ -like '*v*-rtm-*-tree*' })[-1] + 'Windows/SoftEther_VPN_Client/'
          foreach ($item in $soft) {
            if ($item.Name -eq $_) {
              $item.Url = $item.Url.Replace('#softethervpnUrl', $softethervpnUrl)
            }
          }
          break
        }

        "softethervpn_server" {
          # get softether vpn server x64 link
          $Uri = 'http://www.softether-download.com'
          $Parameters = @{
            Uri             = $Uri + '/files/softether/'
            UseBasicParsing = $true
          }
          $softethervpnServerUrl = $Uri + ((Invoke-WebRequest @Parameters).links.href |
            Where-Object { $_ -like '*v*-rtm-*-tree*' })[-1] + 'Windows/SoftEther_VPN_Server_and_VPN_Bridge/'
          foreach ($item in $soft) {
            if ($item.Name -eq $_) {
              $item.Url = $item.Url.Replace('#softethervpnServerUrl', $softethervpnServerUrl)
            }
          }
          break
        }
      }
    }

    foreach ($prog in $soft) {
      #$prog = ($soft | Where-Object {$_.Name -eq $element})
      if ($_ -eq $prog.Name) {
        if ($prog.OneExe2Desktop) {
          Get-File $prog.Pattern $prog.Dir $prog.Name $prog.Url $prog.Arg $prog.RarArg $prog.PatternSetup $prog.OneExe2Desktop
        }
        else {
          Get-File $prog.Pattern $prog.Dir $prog.Name $prog.Url $prog.Arg $prog.RarArg $prog.PatternSetup
        }
        break
      }
    }
  }
}