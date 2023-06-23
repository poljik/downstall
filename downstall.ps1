[CmdletBinding()]
param()
DynamicParam {
	$ParameterName = 'install'
	$RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
	$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
	$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
	$ParameterAttribute.Mandatory = $false
	$ParameterAttribute.Position = 1
	$AttributeCollection.Add($ParameterAttribute)
	$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
  $arrSet = (Get-Content -Path "$path/soft.json" | ConvertFrom-JSON).name
	$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
	$AttributeCollection.Add($ValidateSetAttribute)
	$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string[]], $AttributeCollection)
	$RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
	return $RuntimeParameterDictionary
}
begin {
	[string[]]$install = $PsBoundParameters[$ParameterName]
}

end {
  $Host.UI.RawUI.WindowTitle = "Downstall script 4 Windows $([char]0x00A9) poljik 2019-2023"
  #Requires -RunAsAdministrator
  #Requires -Version 5.1   # Windows 10 LTSC 

  # enable TLSv1.2 for compatibility with older clients
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
  # increase speed of downloading
  $progresspreference = "SilentlyContinue"

  $soft = Get-Content -Path "$path/soft.json" | ConvertFrom-JSON

  # get windows version
  if (!($IsLinux)) {
    $OSVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName).ProductName
    switch -Wildcard ($OSVersion) {
      'Windows 7*' {$global:win = 7; break}
      'Windows 8*' {$global:win = 8; break}
      'Windows 10*' {$global:win = 10; break}
      'Windows 11*' {$global:win = 11; break}
      'Windows Server 2012*' {$global:win = 2012; break}
      'Windows Server 2016*' {$global:win = 2016; break}
      'Windows Server 2019*' {$global:win = 2019; break}
      'Windows Server 2022*' {$global:win = 2022; break}
    }
  } else {
    $path = $path | Split-Path -Parent | Split-Path -Parent -ErrorAction Ignore
  }

  # get bit OS
  if ([System.Environment]::Is64BitOperatingSystem) {
    $global:bitOS = 64
  } else {
    $global:bitOS = 86
  }

  function Find-Dir {
    param (
      [System.Object]$where, [String]$pattern, [String]$dir
    )
    $file = Get-ChildItem -Path $where -Include $pattern -Recurse -Force -ErrorAction Ignore
    if ($file.Count) {
      $filePath = $file[0].DirectoryName
      $fileName = $file[0].Name
    }
    else {
      $dirs = Get-ChildItem -Path $where -Include $dir -Recurse -Force  -ErrorAction Ignore
      if ($dirs.Count) {
        $filePath = $dirs[0].FullName
      }
      else {
        $filePath = $where + "/" + $dir.Replace("*"," ")
      }
    }
    return $filePath, $fileName
  }
  function Compare-Sizes {
    if ($fileName -eq $linkName) {
      $oldSize = (Get-ChildItem -Path $filePath\$fileName).Length
      $newSize = (Invoke-WebRequest $link -Headers @{Referer = $link} -Method HEAD).Headers.'Content-Length'
      if ($newSize.GetType().FullName -eq "System.String[]") {
        $newSize = [Int]$newSize[0]
      } else {
        $newSize = [Int]$newSize
      }
      if ($oldSize -eq $newSize) {
        Write-Host "Skip download",$linkName
        return $false
      } 
      return $true
    }
  }
  function Get-AbsoluteUri {
    try {
      $url0 = Invoke-WebRequest -Method Head -Uri $url
    } catch {
      try {
        $url0 = Invoke-WebRequest -Uri $url
      } catch {
        Write-Error $_
      }
    }
    $url = $url0.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
    if (!$url) {
      $url = $url0.BaseResponse.ResponseUri.AbsoluteUri
    }
    return $url
  }

  function Install-Prog {
    Write-Host "Install", $linkName, $arg
    switch -Wildcard ($name) {
      "winscan2pdf" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:PUBLIC\Desktop -WindowStyle Hidden
        }
        break
      }
      "edgeblock" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "e -o+", `"$filePath\$linkName`", "EdgeBlock\EdgeBlock_x64.exe", $Env:PUBLIC\Desktop -WindowStyle Hidden
        }
        break
      }
      "speedtest" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "e -o+", `"$filePath\$linkName`", "speedtest.exe", $Env:PUBLIC\Desktop -WindowStyle Hidden
        }
        break
      }
      "syncthing" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "e -o+", `"$filePath\$linkName`", "*syncthing.exe" , $Env:PUBLIC\Desktop -WindowStyle Hidden
        }
        break
      }
      "imagine_plugin_*" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:LOCALAPPDATA\Imagine\Plugin -WindowStyle Hidden
        }
        break
      }
      "vvoddpu" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          $patternSetup = "vvoddpu-setup*.exe"
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "e -o+", `"$filePath\$linkName`", $patternSetup, $Env:Temp\temp1 -WindowStyle Hidden -Wait
          $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Force
          Start-Process -FilePath $fileSetup -ArgumentList $arg -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
        }
        break
      }
      "avest" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          $patternSetup = "AvPKISetup*.exe"
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Recurse -Force
          Start-Process -FilePath $fileSetup -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
        }
        break
      }
      "sophia_app" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          $patternSetup = "SophiApp.exe"
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Recurse -Force
          Start-Process -FilePath $fileSetup -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
        }
        break
      }
      "drivers_store_explorer" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          $patternSetup = "rapr.exe"
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Recurse -Force
          Start-Process -FilePath $fileSetup -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
        }
        break
      }
      "ventoy" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          $patternSetup = "Ventoy2Disk.exe"
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Recurse -Force
          Start-Process -FilePath $fileSetup -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
        }
        break
      }
      "victoria" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          $patternSetup = "victoria.exe"
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Recurse -Force
          Start-Process -FilePath $fileSetup -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
        }
        break
      }
      "adobe_reader" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          Start-Process -FilePath $Env:Temp\temp1\acroread.msi -ArgumentList $arg -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
          Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Acrobat Reader.lnk" -Destination $Env:PUBLIC\Desktop -Force
          # delete button-banner "Try Acrobat Pro DC"
          "HKLM:\Software\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" | ForEach-Object {
            New-Item -Path $_ -ItemType Directory -Force > $Null
            Set-ItemProperty -Path $_ -Name bAcroSuppressUpsell -Value 1 -Force
          }
        }
        # original Acrobat Reader
        # "AcroRdrDC*.exe" + "/sAll /msi /norestart /quiet ALLUSERS=1 EULA_ACCEPT=YES"
        # "msiexec.exe /update " + "AcroRdrDC*.msp" + "/quiet /norestart"
        break
      }
      "true_image" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+ -prepack.me", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          Start-Process -FilePath $Env:Temp\temp1\*.exe -ArgumentList $arg -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
          # disable services Acronis Nonstop Backup Service, Acronis Sync Agent Service
          Set-Service afcdpsrv -StartupType Disabled -PassThru -Confirm:$false -ErrorAction SilentlyContinue | Stop-Service -PassThru > $Null
          Set-Service syncagentsrv -StartupType Disabled -PassThru -Confirm:$false -ErrorAction SilentlyContinue | Stop-Service -PassThru > $Null
        }
        break
      }
      "mas" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          Write-Host "You have Windows $win"
          if ($win -eq 10) {
            & ([ScriptBlock]::Create((Invoke-RestMethod https://massgrave.dev/get))) /KMS38 /KMS-Office /KMS-RenewalTask
          }
          if (!($win -eq 10)) {
            & ([ScriptBlock]::Create((Invoke-RestMethod https://massgrave.dev/get))) /KMS-WindowsOffice /KMS-RenewalTask
          }
        }
        break
      }
      "total_commander" {
        if (Test-Path $Env:Programfiles\WinRAR\Winrar.exe) {
          New-Item -ItemType Directory -Force -path $Env:Temp\temp1 > $Null
          $patternSetup = "total commander*litepack*.exe"
          Start-Process -FilePath $Env:Programfiles\WinRAR\Winrar.exe -ArgumentList "x -o+", `"$filePath\$linkName`", $Env:Temp\temp1 -WindowStyle Hidden -Wait
          $fileSetup = Get-ChildItem -Path $Env:Temp\temp1\* -Include $patternSetup -Recurse -Force
          Start-Process -FilePath $fileSetup -ArgumentList $arg  -Wait
          Remove-Item $Env:Temp\temp1 -Recurse -Force
        }

        if ($bitOS -eq 64) {
          $copy_to = ${Env:Programfiles(x86)}
        } else {
          $copy_to = $Env:Programfiles
        }
        $copy_to += "\Total Commander"

        if ($file2copy = Get-ChildItem -Path $filePath -Include wincmd.key -Recurse -Force -ErrorAction Ignore) {
          Copy-Item -Path $file2copy -Destination $copy_to -Force
        }

        if ($file2copy = Get-ChildItem -Path $filePath -Include TOTALCMD*.EXE -Recurse -Force -ErrorAction Ignore) {
          Copy-Item -Path $file2copy -Destination $copy_to -Force
        }
        break
      }
      "office_2016" {
        Start-Process -FilePath $filePath\$linkName -ArgumentList $arg -Wait
        if (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk") {
          Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" -Destination $Env:PUBLIC\Desktop -Force
        }
        if (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk") {
          Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" -Destination $Env:PUBLIC\Desktop -Force
        }
        if (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk") {
            Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk" -Destination $Env:PUBLIC\Desktop -Force
        }
        break
      }
      default {
        if($arg) {
          Start-Process -FilePath $filePath\$linkName -ArgumentList $arg #-Wait # -WindowStyle Hidden
        } else {
          Start-Process -FilePath $filePath\$linkName
        }
        switch ($name) {
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
              Start-Sleep -s 1
            } While (!(Get-ChildItem -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Far Manager 3$bit\Far Manager 3$bit.lnk" -ErrorAction Ignore))
            Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\Far Manager 3$bit\Far Manager 3$bit.lnk" -Destination $Env:PUBLIC\Desktop -Force
            break
          }
          "viber" {
            Do {
              Start-Sleep -s 1
            } While (Get-Process | Where-Object { $_.Path -like "*vibersetup*" })
            Write-Host "Kill Viber autostart"
            Get-Process | Where-Object { $_.Path -like "*Viber*" } | Stop-Process -Force -processname { $_.ProcessName }
            break
          }
          "winrar" {
            if ($file2copy = Get-ChildItem -Path $filePath -Include "rarreg.key" -Recurse -Force -ErrorAction Ignore){
              Copy-Item $file2copy $Env:Programfiles\WinRAR
              Copy-Item -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\WinRAR\WinRAR.lnk" -Destination $Env:PUBLIC\Desktop -Force
            }
            break
          }
          "skype" {
            Write-Host "Kill Skype autostart"
            Get-Process | Where-Object { $_.Path -like "*Skype*" } | Stop-Process -Force -processname { $_.ProcessName }
            break
          }
        }
      }
    }
  }

  function Get-File {
    param (
      [System.Object]$where, [String]$pattern, [String]$dir, [String]$name, [String]$url, [String]$arg
    )
    Clear-Variable -Name fileName -Force -ErrorAction SilentlyContinue
    $filePath, $fileName = Find-Dir $where $pattern $dir
    if ($url) {
      # for yandex disk links
      if (($url -like '*.yandex.*') -or ($url -like '*yadi.sk*')) {
        $ud = [System.Web.HttpUtility]::UrlEncode($url)
        $link = Invoke-RestMethod "https://cloud-api.yandex.net/v1/disk/public/resources?public_key=$ud&limit=200"
        if ($link.type -eq "file") {
          $link = (Invoke-RestMethod "https://cloud-api.yandex.net:443/v1/disk/public/resources/download?public_key=$ud").href
          $namePattern = "&filename="
          $beginName = $link.IndexOf($namePattern) + $namePattern.Length
          $linkName = $link.Substring($beginName, ($link.Length - $beginName))
          $linkName = $linkName.Substring(0, $linkName.IndexOf("&"))
        } else {
          $link = $link._embedded.items | Where-Object { $_.name -like $pattern}
          $linkName = $link[-1].name
          $link = $link[-1].file
        }
      } elseif ($url -like "*drive.google.com/file/d*") {
        try {
          $result = Invoke-WebRequest $url -UseBasicParsing -MaximumRedirection 0 -ea silentlycontinue
        } catch {
          Write-Error $_
        }
        [Array]$links = $result.links.href | Where-Object { $_ -like "*drive.google.com/file/d*" }
        $link = $links[0]
        $linkname = $link | Split-Path -Parent
        $linkID = $linkname | Split-Path -Leaf
        # download the Virus Warning into _tmp.txt
        try {
          Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&id=$linkID" -OutFile "_tmp.txt" -SessionVariable googleDriveSession -ErrorAction SilentlyContinue
        } catch {
          Write-Error $_
        }
        # get confirmation code from _tmp.txt
        $searchString = Select-String -Path "_tmp.txt" -Pattern "confirm="
        $searchString -match "confirm=(?<content>.*)&amp;id=" > $Null
        try {
          $confirmCode = $matches["content"]          
        } catch {}
        # get name of file on google drive
        $searchString = (Select-String -Path "_tmp.txt" -Pattern "a href=").ToString()
        $beginIndex = $searchString.IndexOf("/open?id=" + $linkID)
        $nameString = $searchString.Substring($beginIndex, ($searchString.Length - $beginIndex))
        $beginIndex = $nameString.IndexOf(">") + 1
        $endIndex = $nameString.IndexOf("<")
        $linkname = $nameString.Substring($beginIndex, ($endIndex - $beginIndex))
        Remove-Item "_tmp.txt" -Force
        $link = "https://drive.google.com/uc?export=download&confirm=${confirmCode}&id=$linkID"
      } else {
        # for github links
        $url = (Get-AbsoluteUri) -replace "/tag/", "/expanded_assets/"
        $link = $url
        try {
          $result = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction SilentlyContinue
        } catch {
          Write-Error $_
        }
        [Array]$links = $result.links.href | where-Object { $_ -like $pattern }
        if ($links.Count) {
          $link = $links[0]
          # victoria exception
          if ($name -eq "victoria") {
            $begin = $links[1].LastIndexOf("https:")
            $link = $links[1].Substring($begin, ($links[1].Length - $begin))
          }
          if (!($link.Contains("http"))) {
            if ($link.Contains("/") -and !($name -like "imagine*")) {
              $link = $url.Substring(0,($url.IndexOf("/",8)+1)) + $link
            }
            else {
              # imagine exception
              $link = $url + $link
            }
          }
        }
        $linkName = $link | Split-Path -Leaf
      }
      $linkName = $linkName.ToString() -replace "%20", " "
      if (!($fileName -eq $linkName) -or (Compare-Sizes)) {
        Write-Host "+$linkName"
        if (!(Test-Path $filePath)) {
          New-Item $filePath -ItemType Directory > $Null
        }
        try {
          Invoke-WebRequest -Uri $link -OutFile $filePath\$linkName -Headers @{Referer = $link}
        } catch {
          Write-Error $_
          return
        }
        if (($fileName) -and !($fileName -eq $linkName)) {
          Write-Host "-$fileName deleted!"
          Remove-Item -Path $filePath\$fileName -Force -Recurse
        }
      }
    } else {
      $linkName = $fileName
    }
    if ($install | Where-Object { $_ -like $name}) {
      Install-Prog
    }
  }
  function Test-Names {
    param ( $install, $name )
    if ($install | Where-Object {$_ -like $name}) {
      return $true
    }
    return $false
  }

  # get last version number of IrfanView
  $html = (Invoke-WebRequest -Uri https://www.irfanview.com/checkversion.php).Content
  $searchStr = "Current IrfanView version is: <b>"
  $begin = $html.IndexOf($searchStr)
  $html = $html.Substring($begin + $searchStr.Length, $html.Length - $begin - $searchStr.Length)
  $end = $html.IndexOf("</b>")
  $html = $html.Substring(0, $end)
  $irfanviewVersion = $html.Replace(".","").Replace(" ","")

  # get notepad++ url
  $Parameters = @{
    Uri = 'https://notepad-plus-plus.org/downloads/'
    UseBasicParsing = $true
  }
  $notepadppUrl = ((Invoke-WebRequest @Parameters).links.href |
    Where-Object { $_ -like $Parameters.Uri + '*'})[0]

  # get last opera x64 link
  $Parameters = @{
    Uri = 'https://get.geo.opera.com/pub/opera/desktop/'
    UseBasicParsing = $true
  }
  $operaUrl = $Parameters.Uri + ((Invoke-WebRequest @Parameters).links.href[-1]) + 'win/'

  # get softether vpn client x64 link
  $Uri = 'http://www.softether-download.com'
  $Parameters = @{
    Uri = $Uri+ '/files/softether/'
    UseBasicParsing = $true
  }
  $softethervpnUrl = $Uri + ((Invoke-WebRequest @Parameters).links.href |
    Where-Object { $_ -like '*v*-rtm-*-tree*' })[-1] + 'Windows/SoftEther_VPN_Client/'
  # get softether vpn server x64 link
  $softethervpnServerUrl = $Uri + ((Invoke-WebRequest @Parameters).links.href |
    Where-Object { $_ -like '*v*-rtm-*-tree*' })[-1] + 'Windows/SoftEther_VPN_Server_and_VPN_Bridge/'

  $soft | ForEach-Object {
    $_.Url = $_.Url.Replace('$irfanviewVersion', $irfanviewVersion)
    $_.Url = $_.Url.Replace('$notepadppUrl', $notepadppUrl)
    $_.Url = $_.Url.Replace('$operaUrl', $operaUrl)
    $_.Url = $_.Url.Replace('$softethervpnUrl', $softethervpnUrl)
    $_.Url = $_.Url.Replace('$softethervpnServerUrl', $softethervpnServerUrl)
    if (($install) -and (Test-Names $install $_.Name)) {
      Get-File $path $_.Pattern $_.Dir $_.Name $_.Url $_.Arg
    } elseif (!($install)){
      Get-File $path $_.Pattern $_.Dir $_.Name $_.Url
    }
  }
  #"Url": "https://github.com/massgravel/Microsoft-Activation-Scripts/releases/latest"
}