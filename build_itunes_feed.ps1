param(
  [string]$BaseUrl = 'https://dietgogo7.github.io/cast/',
  [string]$ChannelTitle = 'dietgogo 의 podcast',
  [string]$ChannelDescription = '',
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

$wmp = New-Object -ComObject WMPlayer.OCX
$files = Get-ChildItem -File -Path 'd:\podcast' | Where-Object { @('.mp3','.m4a','.mp4','.wav','.aac') -contains $_.Extension.ToLower() } | Sort-Object -Property LastWriteTime -Descending
$last = (Get-Date).ToString('r')
$out = @()
$out += '<?xml version="1.0" encoding="utf-8"?>'
$out += '<rss version="2.0" xmlns:itunes="http://www.itunes.com/dtds/podcast-1.0.dtd">'
$out += '  <channel>'
$out += "    <title>$(Escape-Xml $ChannelTitle)</title>"
$out += "    <link>$BaseUrl</link>"
if ($ChannelDescription -ne '') { $out += "    <description>$(Escape-Xml $ChannelDescription)</description>" } else { $out += "    <description>자동 생성된 팟캐스트 피드</description>" }
$out += "    <language>$Language</language>"
$out += "    <lastBuildDate>$last</lastBuildDate>"
$out += "    <itunes:explicit>$Explicit</itunes:explicit>"
if ($Author -ne '') { $out += "    <itunes:author>$(Escape-Xml $Author)</itunes:author>" }
$out += "    <itunes:image href='$ImageUrl' />"

foreach ($f in $files) {
  $path = $f.FullName
  try {
    $m = $wmp.newMedia($path)
    $durSec = [math]::Round($m.duration)
  } catch {
    $durSec = 0
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
  $out += "    <enclosure url='$enclosureUrl' length='$($f.Length)' type='audio/mpeg' />"
  $out += "    <guid>$enclosureUrl</guid>"
  $out += "    <pubDate>$pub</pubDate>"
  if ($duration -ne '') { $out += "    <itunes:duration>$duration</itunes:duration>" }
  $out += "    <description>파일 크기: $($f.Length) bytes</description>"
  $out += '  </item>'
}

$out += '  </channel>'
$out += '</rss>'

# Write with BOM
$pathOut = 'd:\podcast\feed.xml'
$text = $out -join "`n"
$pre = [System.Text.Encoding]::UTF8.GetPreamble()
$body = [System.Text.Encoding]::UTF8.GetBytes($text)
$all = New-Object byte[] ($pre.Length + $body.Length)
[Array]::Copy($pre,0,$all,0,$pre.Length)
[Array]::Copy($body,0,$all,$pre.Length,$body.Length)
[System.IO.File]::WriteAllBytes($pathOut,$all)
Write-Output "Wrote iTunes-compatible feed.xml with $($files.Count) items"
