$Host.UI.RawUI.WindowTitle = "Downstall script 4 Windows $([char]0x00A9) poljik 2019-2023"
#Requires -RunAsAdministrator
#Requires -Version 5.1   # Windows 10 LTSC 

# enable TLSv1.2 for compatibility with older clients
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
# increase speed of downloading
$progresspreference = "SilentlyContinue"

$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
function Find-Dir {
  param (
    [System.Object]$where, [String]$what, [String]$whatDir, [String]$type
  )
  $file = Get-ChildItem $where -Include $what -Recurse -Force
  if ($file.Count) {
    $filePath = $file[0].DirectoryName
    $fileName = $file[0].Name
  }
  else {
    $dir = Get-ChildItem $where -Include $whatDir -Recurse -Force
    if ($dir.Count) {
      $filePath = $dir[0].FullName
    }
    else {
      $filePath = $where + "/" + $type.Replace("*"," ") + "/" + $whatDir.Replace("*"," ")
    }
  }
  return $filePath, $fileName
}
function Compare-FileSizes {
  if ($fileName -eq $linkName) {
    $oldFileSize = (Get-ChildItem $filePath\$fileName).Length
    $newFileSize = (Invoke-WebRequest $link -Headers @{Referer = $link} -Method HEAD).Headers.'Content-Length'
    if ($newFileSize.GetType().FullName -eq "System.String[]") {
      $newFileSize = [Int]$newFileSize[0]
    } else {
      $newFileSize = [Int]$newFileSize
    }
    if ($oldFileSize -eq $newFileSize) {
      Write-Host "The size is the same, skip"
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
function Get-File {
  param (
    [System.Object]$where, [String]$arr
  )
  $what, $whatDir, $url, $type = $arr.Split()
  Clear-Variable -Name fileName -Force -ErrorAction SilentlyContinue
  $filePath, $fileName = Find-Dir $where $what $whatDir $type
  $url = (Get-AbsoluteUri) -replace "/tag/", "/expanded_assets/"
  $link = $url
  try {
    $result = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction SilentlyContinue
  } catch {
    Write-Error $_
  }
  [Array]$links = $result.links.href | where-Object { $_ -like $what }
  if ($links.Count) {
    $link = $links[0]
    if ($whatDir -eq "victoria") {
      $begin = $links[1].LastIndexOf("https:")
      $link = $links[1].Substring($begin, ($links[1].Length - $begin))
    }
    if (!($link.Contains("http"))) {
      if ($link.Contains("/") -and !($whatDir -eq 'imagine')) {
        $link = $url.Substring(0,($url.IndexOf("/",8)+1)) + $link
      }
      else {
        $link = $url + "/" + $link
      }
    }
    $linkName = ($link | Split-Path -Leaf).ToString()
  }
  $linkName = ($link | Split-Path -Leaf).ToString()
  $linkName = ($link | Split-Path -Leaf) -replace "%20", " "
  if (!($fileName -eq $linkName) -or (Compare-FileSizes)) {
    Write-Host "  "$whatDir.Replace("*"," ")"  "
    # Write-Host "downloading to $filePath/"
    Write-Host "+$linkName"
    if (!(Test-Path $filePath)) {
      New-Item $filePath -ItemType Directory > $Null
    }
    try {
      Invoke-WebRequest -Uri $link -OutFile $filePath\$linkName -Headers @{Referer = $link}
    } catch {
      Write-Error $_
    }
      if (($fileName) -and !($fileName -eq $linkName)) {
      Write-Host "-$fileName deleted!"
      Remove-Item -Path $filePath\$fileName -Force -Recurse
    }
  }
}

# get last version number of IrfanView
$html = (Invoke-WebRequest -Uri https://www.irfanview.com/checkversion.php).Content
$searchStr = "Current IrfanView version is: <b>"
$begin = $html.IndexOf($searchStr)
$html = $html.Substring($begin + $searchStr.Length, $html.Length - $begin - $searchStr.Length)
$end = $html.IndexOf("</b>")
$html = $html.Substring(0, $end)
$irfanVersion = $html.Replace(".","").Replace(" ","")

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
  "*iview*x64_setup.exe irfanview https://www.irfanview.info/files/iview"+$irfanVersion+"_x64_setup.exe media"
  "*iview*_plugins_x64.zip irfanview https://www.irfanview.info/files/iview"+$irfanVersion+"_plugins_x64.zip media"
  "irfanview_lang_russian.exe irfanview https://www.irfanview.net/lang/irfanview_lang_russian.exe media"
  "Cloudflare_WARP_Release-x64.msi cloudflare_warp https://1111-releases.cloudflareclient.com/windows/Cloudflare_WARP_Release-x64.msi internet"
  "EdgeBlock.zip edgeblock https://www.sordum.org/files/downloads.php?st-edge-block soft"
  "reduced_rate_nds.zip platelschik*eaes https://www.nalog.gov.by/upload/reduced_rate_nds.zip by"
  "*firefox*setup*.msi firefox https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=ru internet"
  "*thunderbird*setup*.msi thunderbird https://download.mozilla.org/?product=thunderbird-msi-latest-ssl&os=win64&lang=ru internet"
  "*aimp*.exe aimp https://www.aimp.ru/?do=download.file&id=29 media"
  "*skype*.msi skype https://go.skype.com/msi-download internet"
  "*windowsadmincenter*.msi windowsadmincenter https://aka.ms/wacdownload soft"
  "anydesk.exe anydesk https://download.anydesk.com/AnyDesk.exe internet"
  #"TeamViewerQS.exe teamviewer https://download.teamviewer.com/download/TeamViewerQS.exe internet"
  "*chrome*64.msi chrome https://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise64.msi internet"
  "rdpwrap.ini rdpwrap https://raw.githubusercontent.com/affinityv/INI-RDPWRAP/master/rdpwrap.ini soft"
  "*winscan2pdf_portable.zip winscan2pdf https://softwareok.com/Download/WinScan2PDF_Portable.zip soft"
  "*ZeroTier*.msi zerotier https://download.zerotier.com/dist/ZeroTier%20One.msi internet"
  "*ookla-speedtest*win64.zip speedtest https://www.speedtest.net/apps/cli internet"
  "*Imagine_Plugin_Archive*x64*.zip imagine https://www.nyam.pe.kr/dev/imagine media"
  "*Imagine_Plugin_DCRaw*x64*.zip imagine https://www.nyam.pe.kr/dev/imagine media"
  "*Imagine_Plugin_JPEG2000*x64*.zip imagine https://www.nyam.pe.kr/dev/imagine media"
  "*VeraCrypt*Setup*.exe veracrypt https://www.veracrypt.fr/en/Downloads.html soft"
  "*far*.x64.*.msi far*manager https://www.farmanager.com/download.php?l=ru file*managers"
  "*vc_redist*x86*.exe zerotier https://docs.microsoft.com/ru-RU/cpp/windows/latest-supported-vc-redist?view=msvc-170 internet"
  "*vc_redist*x64*.exe zerotier https://docs.microsoft.com/ru-RU/cpp/windows/latest-supported-vc-redist?view=msvc-170 internet"
  "*vvoddpu*.zip vvoddpu https://ssf.gov.by/ru/po-fonda-ru/ by"
  "*EDeclaration*(Include_JRE).exe edeclaration https://lkfl.portal.nalog.gov.by/web/guest/arm_soft by"
  "*АРМ*Плательщик*ЕАЭС*.exe platelschik*eaes https://www.nalog.gov.by/software/arm-payer/ by"
  "*AvPKISetup*bel*.zip edeclaration https://www.avest.by/crypto/csp.htm by"
  "*wireguard-amd64*.msi wireguard https://download.wireguard.com/windows-client internet"
  "*trjsetup*.exe trojan*remover https://www.simplysup.com/tremover/download_full.html antivirus"
  "*viber*setup*.exe viber https://www.viber.com/ru/download/ internet"
  "*winrar-x64*ru.exe winrar https://www.rarlab.com/download.htm soft"
  "*syncthing*windows-amd64*.zip syncthing https://syncthing.net/downloads/ internet"
  "*Sophia.Script.for.Windows.10.LTSC*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  "*Sophia.Script.for.Windows.10.PowerShell.7*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  "*Sophia.Script.for.Windows.11.PowerShell.7*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  "*Sophia.Script.Wrapper.*.zip sophia*script https://github.com/farag2/Sophia-Script-for-Windows/releases/latest tweaks"
  "*SophiApp.zip sophia*script https://github.com/Sophia-Community/SophiApp/releases/latest tweaks"
  "*MAS_*_Password_1234.7z activation*10 https://github.com/massgravel/Microsoft-Activation-Scripts/releases/latest soft/activation"
  "*DriverStoreExplorer*.zip drivers*store*explorer https://github.com/lostindark/DriverStoreExplorer/releases/latest drivers"
  "*ventoy*linux.tar.gz ventoy https://github.com/ventoy/Ventoy/releases/latest soft"
  "*ventoy*windows.zip ventoy https://github.com/ventoy/Ventoy/releases/latest soft"
  "*RDPWInst*.msi rdpwrap https://github.com/stascorp/rdpwrap/releases/latest soft"
  "*RDPWrap*.zip rdpwrap https://github.com/stascorp/rdpwrap/releases/latest soft"
  "*install.exe onionfruit https://github.com/dragonfruitnetwork/onionfruit/releases/latest internet"
  "*Imagine_*_x64_Unicode.exe imagine https://www.nyam.pe.kr/dev/imagine/ media"
  "*PotPlayerSetup64*.exe pot*player http://potplayer.daum.net/ media"
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
Write-host "   x32 soft"
if (!(Test-Path $path/x32)) {
  New-Item $path/x32 -ItemType Directory > $Null
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
  "*firefox*setup*.msi firefox https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win&lang=ru internet"
  "*thunderbird*setup*.msi thunderbird https://download.mozilla.org/?product=thunderbird-msi-latest-ssl&os=win&lang=ru internet"
  "*chrome*.msi chrome https://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise.msi internet"
  "*aimp*.exe aimp https://www.aimp.ru/?do=download.file&id=4 media"
  "*far*.x86.*.msi far*manager https://www.farmanager.com/download.php?l=ru file*managers"
  "*winrar-x32*ru.exe winrar https://www.rarlab.com/download.htm soft"
  "*opera*setup.exe opera $operaUrl internet"
) | ForEach-Object {
  Get-File $path/x32 $_.Split()
}