$files = Get-ChildItem -File -Path 'd:\podcast' | Where-Object { @('.mp3','.m4a','.mp4','.wav','.aac') -contains $_.Extension.ToLower() } | Sort-Object LastWriteTime -Descending
$last = (Get-Date).ToString('r')
$out = @()
$out += '<?xml version="1.0" encoding="utf-8"?>'
$out += '<rss version="2.0">'
$out += '  <channel>'
$out += '    <title>Podcast</title>'
$out += '    <link>./</link>'
$out += '    <description>자동 생성된 팟캐스트 피드</description>'
$out += "    <lastBuildDate>$last</lastBuildDate>"
foreach ($f in $files) {
  $type = switch ($f.Extension.ToLower()) { '.mp3' {'audio/mpeg'} '.m4a' {'audio/mp4'} '.mp4' {'video/mp4'} '.wav' {'audio/wav'} '.aac' {'audio/aac'} default {'application/octet-stream'} }
  $pub = (Get-Date $f.LastWriteTime).ToString('r')
  $out += '  <item>'
  $out += "    <title>$($f.Name)</title>"
  $out += "    <enclosure url='$($f.Name)' length='$($f.Length)' type='$type' />"
  $out += "    <guid>$($f.Name)</guid>"
  $out += "    <pubDate>$pub</pubDate>"
  $out += "    <description>파일 크기: $($f.Length) bytes</description>"
  $out += '  </item>'
}
$out += '  </channel>'
$out += '</rss>'
$out | Set-Content -Path 'd:\podcast\feed.xml' -Encoding utf8
Write-Output "Rebuilt feed.xml with $($files.Count) items"
