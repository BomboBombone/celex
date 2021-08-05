$celexPath = $args[0]
$linkPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Celex.lnk"
$targetPath = "$celexPath/Celex.py"
$link = (New-Object -ComObject WScript.Shell).CreateShortcut( $linkpath )
$link.TargetPath = $targetPath
$link.IconLocation = "$celexPath/icon.ico, 0"
$link.Save()