param([string]$BasePath = 'd:\podcast', [string]$ChannelTitle = 'Podcast', [string]$ChannelDescription = '', [string]$ChannelDescriptionFile = '')

# Find media files
$files = Get-ChildItem -File -Path $BasePath | Where-Object { @('.mp3','.m4a','.mp4','.wav','.aac') -contains $_.Extension.ToLower() } | Sort-Object LastWriteTime -Descending
$last = (Get-Date).ToString('r')

# If a description file is provided, read it as UTF8
if ($ChannelDescriptionFile -and (Test-Path $ChannelDescriptionFile)) {
  try { $ChannelDescription = Get-Content -Path $ChannelDescriptionFile -Raw -Encoding UTF8 } catch { }
}

if (-not $ChannelDescription) { $ChannelDescription = 'Auto-generated podcast feed' }

$out = @()
$out += '<?xml version="1.0" encoding="utf-8"?>'
$out += '<rss version="2.0">'
$out += '  <channel>'
$out += "    <title>$ChannelTitle</title>"
$out += "    <link>./</link>"
$out += "    <description>$ChannelDescription</description>"
$out += "    <lastBuildDate>$last</lastBuildDate>"

foreach ($f in $files) {
  $type = switch ($f.Extension.ToLower()) { '.mp3' {'audio/mpeg'} '.m4a' {'audio/mp4'} '.mp4' {'video/mp4'} '.wav' {'audio/wav'} '.aac' {'audio/aac'} default {'application/octet-stream'} }
  $pub = (Get-Date $f.LastWriteTime).ToString('r')
  $title = $f.BaseName
  $out += '  <item>'
  $out += "    <title>$title</title>"
  $out += '    <enclosure url="' + $f.Name + '" length="' + $f.Length + '" type="' + $type + '" />'
  $out += "    <guid>$($f.Name)</guid>"
  $out += "    <pubDate>$pub</pubDate>"
  # include same-name .md as description if present (read as UTF8)
  $mdPath = Join-Path (Split-Path $f.FullName -Parent) ($f.BaseName + '.md')
  if (Test-Path $mdPath) {
    try {
      $mdRaw = Get-Content -Path $mdPath -Raw -Encoding UTF8
      $cdata = '<![CDATA[' + $mdRaw + ']]>'
      $out += '    <description>' + $cdata + '</description>'
    } catch {
      # ignore
    }
  }
  $out += '  </item>'
}

$out += '  </channel>'
$out += '</rss>'

# Write with UTF-8 BOM to avoid Korean garbling
$pathOut = Join-Path $BasePath 'feed.xml'
$text = $out -join "`n"
$pre = [System.Text.Encoding]::UTF8.GetPreamble()
$body = [System.Text.Encoding]::UTF8.GetBytes($text)
$all = New-Object byte[] ($pre.Length + $body.Length)
[Array]::Copy($pre,0,$all,0,$pre.Length)
[Array]::Copy($body,0,$all,$pre.Length,$body.Length)
[System.IO.File]::WriteAllBytes($pathOut,$all)
Write-Output "Rebuilt feed.xml with $($files.Count) items (UTF-8 BOM)"
