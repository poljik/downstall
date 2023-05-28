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
    [System.Object]$where, [String]$pattern, [String]$dir, [String]$name
  )
  $file = Get-ChildItem $where -Include $pattern -Recurse -Force
  if ($file.Count) {
    $filePath = $file[0].DirectoryName
    $fileName = $file[0].Name
  }
  else {
    $dirs = Get-ChildItem $where -Include $name -Recurse -Force
    if ($dirs.Count) {
      $filePath = $dirs[0].FullName
    }
    else {
      $filePath = $where + "/" + $dir.Replace("*"," ") + "/" + $name.Replace("*"," ")
    }
  }
  return $filePath, $fileName
}
function Compare-Sizes {
  if ($fileName -eq $linkName) {
    $oldSize = (Get-ChildItem $filePath\$fileName).Length
    $newSize = (Invoke-WebRequest $link -Headers @{Referer = $link} -Method HEAD).Headers.'Content-Length'
    if ($newSize.GetType().FullName -eq "System.String[]") {
      $newSize = [Int]$newSize[0]
    } else {
      $newSize = [Int]$newSize
    }
    if ($oldSize -eq $newSize) {
      Write-Host $pattern.Replace("*"," "),"- skip"
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
    [System.Object]$where, [String]$pattern, [String]$dir, [String]$name, [String]$url
  )
  Clear-Variable -Name fileName -Force -ErrorAction SilentlyContinue
  $filePath, $fileName = Find-Dir $where $pattern $dir $name
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
      if ($link.Contains("/") -and !($name -eq 'imagine')) {
        $link = $url.Substring(0,($url.IndexOf("/",8)+1)) + $link
      }
      else {
        # imagine exception
        $link = $url + "/" + $link
      }
    }
  }
  $linkName = ($link | Split-Path -Leaf).ToString()
  $linkName = ($link | Split-Path -Leaf) -replace "%20", " "
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

$soft = (Get-Content -Path "./soft.json" | ConvertFrom-JSON).soft
$soft | ForEach-Object {
  $_.Url = $_.Url.Replace('$irfanviewVersion', $irfanviewVersion)
  $_.Url = $_.Url.Replace('$notepadppUrl', $notepadppUrl)
  $_.Url = $_.Url.Replace('$operaUrl', $operaUrl)
  $_.Url = $_.Url.Replace('$softethervpnUrl', $softethervpnUrl)
  $_.Url = $_.Url.Replace('$softethervpnServerUrl', $softethervpnServerUrl)
  Get-File $path $_.Pattern $_.Dir $_.Name $_.Url
}

# x32 soft
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

$soft = (Get-Content -Path "./softX32.json" | ConvertFrom-JSON).soft
$soft | ForEach-Object {
  $_.Url = $_.Url.Replace('$operaUrl', $operaUrl)
  Get-File $path/x32 $_.Pattern $_.Dir $_.Name $_.Url
}