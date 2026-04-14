param(
  [string]$BaseUrl = 'https://dietgogo7.github.io/cast/',
  [string]$ChannelTitle = 'dietgogo 의 podcast',
  [string]$ChannelDescription = '',
  [string]$ChannelDescriptionFile = '',
  [string]$ImageUrl = 'https://image.yes24.com/sysimage/mv3/com/ico_ai02.svg',
  [string]$Author = '',
  [string]$Language = 'ko-KR',
  [string]$Explicit = 'no'
)

function Escape-Xml([string]$s){
  if (-not $s) { return '' }
  $s = $s.Replace('&','&amp;')
  $s = $s.Replace('<','&lt;')
  $s = $s.Replace('>','&gt;')
  $s = $s.Replace('"','&quot;')
  $s = $s.Replace("'","&apos;")
  return $s
}

# Create Windows Media Player COM only when needed (catch if not available)
try { $wmp = New-Object -ComObject WMPlayer.OCX } catch { $wmp = $null }

# If a description file is provided, read it as UTF8 to support Korean safely
if ($ChannelDescriptionFile -and (Test-Path $ChannelDescriptionFile)) {
  try {
    $ChannelDescription = Get-Content -Path $ChannelDescriptionFile -Raw -Encoding UTF8
  } catch {
    # ignore and fall back to provided ChannelDescription
  }
}

$files = Get-ChildItem -File -Path 'd:\podcast' | Where-Object { @('.mp3','.m4a','.mp4','.wav','.aac') -contains $_.Extension.ToLower() } | Sort-Object -Property LastWriteTime -Descending
$last = (Get-Date).ToString('r')
$out = @()
$out += '<?xml version="1.0" encoding="utf-8"?>'
$out += '<rss version="2.0" xmlns:itunes="http://www.itunes.com/dtds/podcast-1.0.dtd">'
$out += '  <channel>'
$out += "    <title>$(Escape-Xml $ChannelTitle)</title>"
$out += "    <link>$(Escape-Xml $BaseUrl)</link>"
if ($ChannelDescription -ne '') { $out += "    <description>$(Escape-Xml $ChannelDescription)</description>" } else { $out += "    <description>Auto-generated podcast feed</description>" }
$out += "    <language>$(Escape-Xml $Language)</language>"
$out += "    <lastBuildDate>$last</lastBuildDate>"
$out += "    <itunes:explicit>$(Escape-Xml $Explicit)</itunes:explicit>"
if ($Author -ne '') { $out += "    <itunes:author>$(Escape-Xml $Author)</itunes:author>" }
# Add both RSS image and iTunes image (use double quotes for attributes)
$out += "    <image>"
$out += "      <url>$(Escape-Xml $ImageUrl)</url>"
$out += "      <title>$(Escape-Xml $ChannelTitle)</title>"
$out += "      <link>$(Escape-Xml $BaseUrl)</link>"
$out += "    </image>"
$out += '    <itunes:image href="' + (Escape-Xml $ImageUrl) + '" />'

foreach ($f in $files) {
  $path = $f.FullName
  $durSec = 0
  if ($wmp) {
    try { $m = $wmp.newMedia($path); $durSec = [math]::Round($m.duration) } catch { $durSec = 0 }
  }
  if ($durSec -gt 0) {
    $ts = [TimeSpan]::FromSeconds($durSec)
    $duration = if ($ts.Hours -gt 0) { "{0}:{1:D2}:{2:D2}" -f $ts.Hours,$ts.Minutes,$ts.Seconds } else { "{0}:{1:D2}" -f $ts.Minutes,$ts.Seconds }
  } else { $duration = '' }
  $title = Escape-Xml $f.BaseName
  $pub = (Get-Date $f.LastWriteTime).ToString('r')
  $enclosureUrl = ($BaseUrl.TrimEnd('/') + '/' + $f.Name)
  $out += '  <item>'
  $out += "    <title>$title</title>"
  $out += '    <enclosure url="' + (Escape-Xml $enclosureUrl) + '" length="' + $f.Length + '" type="audio/mpeg" />'
  $out += "    <guid>$(Escape-Xml $enclosureUrl)</guid>"
  $out += "    <pubDate>$pub</pubDate>"
  if ($duration -ne '') { $out += "    <itunes:duration>$duration</itunes:duration>" }
  $out += "    <description>File size: $($f.Length) bytes</description>"
  $out += '  </item>'
}

$out += '  </channel>'
$out += '</rss>'

# Write with BOM (UTF-8)
$pathOut = 'd:\podcast\feed.xml'
$text = $out -join "`n"
$pre = [System.Text.Encoding]::UTF8.GetPreamble()
$body = [System.Text.Encoding]::UTF8.GetBytes($text)
$all = New-Object byte[] ($pre.Length + $body.Length)
[Array]::Copy($pre,0,$all,0,$pre.Length)
[Array]::Copy($body,0,$all,$pre.Length,$body.Length)
[System.IO.File]::WriteAllBytes($pathOut,$all)
Write-Output "Wrote iTunes-compatible feed.xml with $($files.Count) items (UTF-8 BOM)"
