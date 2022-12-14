$Host.UI.RawUI.WindowTitle = "Downstall script 4 Windows $([char]0x00A9) poljik 2019-2022"
#Requires -RunAsAdministrator
#Requires -Version 5.1   # Windows 10 LTSC 

# enable TLSv1.2 for compatibility with older clients
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
# increase speed of downloading
$progresspreference = "SilentlyContinue"

$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
function Find-Dir {
  param (
    [System.Object]$Where, [String]$What, [String]$WhatDir, [String]$Type
  )
  $file = Get-ChildItem $Where -Include $What -Recurse -Force
  if ($file.Count) {
    if ($file.Count -ge 2) {
      $filepath = $file[0].DirectoryName
      $filename = $file[0].Name
    }
    else {
      $filepath = $file.DirectoryName
      $filename = $file.Name
    }
  }
  else {
    $dir = Get-ChildItem $Where -Include $WhatDir -Recurse -Force
    if ($dir.Count) {
      if ($dir.Count -ge 2) {
        $filepath = $dir[0].FullName
      }
      else {
        $filepath = $dir.FullName
      }
    }
    else {
      $filepath = $Where + "/" + $Type.Replace("*"," ") + "/" + $WhatDir.Replace("*"," ")
      if (!(Test-Path $filepath)) {
        New-Item $filepath -ItemType Directory > $Null
      }
    }
  }
  return $filepath, $filename
}
function Get-File {
  param (
    [System.Object]$Where, [String]$Arr
  )
  $What, $WhatDir, $Site, $Type, $AnywayDownload = $Arr.Split()
  $filename = ''
  $filepath, $filename = Find-Dir $Where $What $WhatDir $Type
  $result = Invoke-WebRequest -Uri $Site -UseBasicParsing -ErrorAction SilentlyContinue
  $links = $result.links.href | Where-Object { $_ -like $What }
  if ($links.Count) {
    if ($links.Count -ge 2) {
      if ($WhatDir -eq "victoria") {
        $begin = $links[1].LastIndexOf("https:")
        $link = $links[1].Substring($begin, ($links[1].Length - $begin))
      }
      else {
        $link = $links[0]
      }
    }
    else {
      $link = $links
    }
    if (!($link.Contains("http"))) {
      if ($link.Contains("/") -and !($WhatDir -eq 'imagine')) {
        $link = $Site.Substring(0,($Site.IndexOf("/",8)+1)) + $link
      }
      else {
        $link = $Site + "/" + $link
      }
    }
    $linkname = ($link | Split-Path -Leaf).ToString()
    if (!($filename -eq $linkname) -or ($AnywayDownload)) {
      Write-Host "  "$WhatDir.Replace("*"," ")"  " -ForegroundColor DarkBlue
      Write-Host "+downloading to $filepath/" -ForegroundColor Green
      Write-Host "+$linkname" -ForegroundColor Yellow
      try {
        Invoke-WebRequest -Uri $link -OutFile $filepath\$linkname
      } catch {
        Write-Error $_
      }
        if ($filename -and !$AnywayDownload) {
        Write-Host "-$filename deleted!" -ForegroundColor Red
        Remove-Item -Path $filepath\$filename -Force -Recurse
      }
    }
  }
}
function Get-FileFromLink {
  param (
    [System.Object]$Where, [String]$Arr
  )
  $What, $WhatDir, $Site, $Type = $Arr.Split()
  $filename = ''
  $filepath, $filename = Find-Dir $Where $What $WhatDir $Type
  $link = (Invoke-WebRequest -Uri $Site).BaseResponse.RequestMessage.RequestUri.absoluteuri
  $linkname = ($link | Split-Path -Leaf) -replace "%20", " "
  if (!($filename -eq $linkname)) {
    Write-Host "  "$WhatDir.Replace("*"," ")"  " -ForegroundColor DarkBlue
    Write-Host "+downloading to $filepath/" -ForegroundColor Green
    Write-Host "+$linkname" -ForegroundColor Yellow
    try {
      Invoke-WebRequest -Uri $link -OutFile $filepath\$linkname
    } catch {
      Write-Error $_
    }
    if ($filename) {
      Write-Host "-$filename deleted!" -ForegroundColor Red
      Remove-Item -Path $filepath\$filename -Force -Recurse
    }
  }
}

@(
  "*firefox*setup*.msi firefox https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=ru internet"
  "*thunderbird*setup*.msi thunderbird https://download.mozilla.org/?product=thunderbird-msi-latest-ssl&os=win64&lang=ru internet"
  "*aimp*.exe aimp https://www.aimp.ru/?do=download.file&id=4 media"
  "*skype*.msi skype https://go.skype.com/msi-download internet"
  "*windowsadmincenter*.msi windowsadmincenter https://aka.ms/wacdownload soft"
  "anydesk.exe anydesk https://download.anydesk.com/AnyDesk.exe internet"
  #"TeamViewerQS.exe teamviewer https://download.teamviewer.com/download/TeamViewerQS.exe internet"
  "*chrome*64.msi chrome https://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise64.msi internet"
  "rdpwrap.ini rdpwrap https://raw.githubusercontent.com/affinityv/INI-RDPWRAP/master/rdpwrap.ini soft"
  "*winscan2pdf_portable.zip winscan2pdf https://softwareok.com/Download/WinScan2PDF_Portable.zip soft"
  "*ZeroTier*.msi zerotier https://download.zerotier.com/dist/ZeroTier%20One.msi internet"
) | ForEach-Object {
  Get-FileFromLink $path $_.Split()
}

# get notepad++ url
$Parameters = @{
  Uri = 'https://notepad-plus-plus.org/downloads/'
  UseBasicParsing = $true
}
$nppUrl = ((Invoke-WebRequest @Parameters).links.href |
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
$urlSoftether = $Uri + ((Invoke-WebRequest @Parameters).links.href |
  Where-Object { $_ -like '*v*-rtm-*-tree*' })[-1] + 'Windows/SoftEther_VPN_Client/'
# get softether vpn server x64 link
$urlSoftetherServer = $Uri + ((Invoke-WebRequest @Parameters).links.href |
  Where-Object { $_ -like '*v*-rtm-*-tree*' })[-1] + 'Windows/SoftEther_VPN_Server_and_VPN_Bridge/'

@(
  "*far*.x64.*.msi far*manager https://www.farmanager.com/download.php?l=ru file*managers"
  "*vc_redist*x86*.exe zerotier https://docs.microsoft.com/ru-RU/cpp/windows/latest-supported-vc-redist?view=msvc-170 internet"
  "*vc_redist*x64*.exe zerotier https://docs.microsoft.com/ru-RU/cpp/windows/latest-supported-vc-redist?view=msvc-170 internet"
  "*vvoddpu*.zip vvoddpu https://ssf.gov.by/ru/po-fonda-ru/ by"
  "*EDeclaration*(Include_JRE).exe edeclaration https://lkfl.portal.nalog.gov.by/web/guest/arm_soft by"
  "*ARM-Platelschik-EAES*.exe platelschik*eaes https://www.nalog.gov.by/software/#2 by"
  "*AvPKISetup*bel*.zip edeclaration https://nces.by/pki/info/software/#Программное%20обеспечение by"
  "*wireguard-amd64*.msi wireguard https://download.wireguard.com/windows-client internet"
  "*trjsetup*.exe trojan*remover https://www.simplysup.com/tremover/download_full.html antivirus $true"
  "*viber*setup*.exe viber https://www.viber.com/ru/download/ internet $true"
  "*winrar-x64*ru.exe winrar https://www.rarlab.com/download.htm soft"
  "*syncthing*windows-amd64*.zip syncthing https://syncthing.net/downloads/ internet"
  "*Sophia.Script.for.Windows.10.LTSC*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  "*Sophia.Script.for.Windows.10.PowerShell.7*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  "*Sophia.Script.for.Windows.11.PowerShell.7*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  "*Sophia.Script.Wrapper.*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  # "*MAS_*_Password_1234.7z activation*10 https://github.com/massgravel/Microsoft-Activation-Scripts/releases/latest soft/activation"
  # "*DriverStoreExplorer*.zip drivers*store*explorer https://github.com/lostindark/DriverStoreExplorer/releases/latest drivers"
  # "*ventoy*linux.tar.gz ventoy https://github.com/ventoy/Ventoy/releases/latest soft"
  "*RDPWInst*.msi rdpwrap https://github.com/stascorp/rdpwrap/releases/latest soft"
  "*RDPWrap*.zip rdpwrap https://github.com/stascorp/rdpwrap/releases/latest soft"
  "*install.exe onionfruit https://github.com/dragonfruitnetwork/onionfruit/releases/latest internet $true"
  "*Imagine_*_x64_Unicode.exe imagine https://www.nyam.pe.kr/dev/imagine/ media"
  "*PotPlayerSetup64*.exe pot*player http://potplayer.daum.net/ media $true"
  "*7z*-x64.msi 7zip https://www.7-zip.org/download.html soft"
  "*Victoria*.zip victoria https://hdd.by/victoria/ portable"
  "*npp.*.Installer.x64.exe notepad++ $nppUrl soft"
  "*opera*setup_x64.exe opera $operaUrl internet"
  "*softether-vpnclient-v*-rtm-*-windows-x86_x64-intel.exe softethervpn $urlSoftether internet"
  "*softether-vpnserver_vpnbridge-v*-rtm-*-windows-x86_x64-intel.exe softethervpn*server $urlSoftetherServer internet"
) | ForEach-Object {
  Get-File $path $_.Split()
}

# x32 soft
Write-host "   x32 soft" -BackgroundColor Blue
if (!(Test-Path $path/x32)) {
  New-Item $path/x32 -ItemType Directory > $Null
}
@(
  "*firefox*setup*.msi firefox https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win&lang=ru internet"
  "*thunderbird*setup*.msi thunderbird https://download.mozilla.org/?product=thunderbird-msi-latest-ssl&os=win&lang=ru internet"
  "*chrome*.msi chrome https://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise.msi internet"
) | ForEach-Object {
  Get-FileFromLink $path/x32 $_.Split()
}

# get last opera x32 link
if (!$operaUrl) {
  $Parameters = @{
    Uri = 'https://get.geo.opera.com/pub/opera/desktop/'
    UseBasicParsing = $true
  }
  $operaUrl = $Parameters.Uri + ((Invoke-WebRequest @Parameters).links.href[-1]) + 'win/'
}
@(
  "*far*.x86.*.msi far*manager https://www.farmanager.com/download.php?l=ru file*managers"
  "*winrar-x32*ru.exe winrar https://www.rarlab.com/download.htm soft"
  "*opera*setup.exe opera $operaUrl internet"
) | ForEach-Object {
  Get-File $path/x32 $_.Split()
}