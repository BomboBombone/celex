$linkPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Celex.lnk"
$targetPath = "C:/Program Files/Celex/Celex.py"
$link = (New-Object -ComObject WScript.Shell).CreateShortcut( $linkpath )
$link.TargetPath = $targetPath
$link.IconLocation = "C:/Program Files/Celex/icon.ico, 0"
$link.Save()