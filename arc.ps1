# Kerio Connect mail archivation

$SOURCE_DIR="C:\Program Files (x86)\Kerio\MailServer\store\mail\domain.ru"
$ARC_BASE_DIR="C:\mail_arc"
$FILTER = "*.eml"
$EXCLUDE = "Contacts|lavna\.mtu|abonotdel" # regexp, e.g. "str1|str2"

if ($args.Length -ne 1)
{
	Write-Host "Usage:   .\arc.ps1 MM/YYYY"
	Write-Host "Example: .\arc.ps1 8/2015 - moves mails created until 01.09.2015"
	exit
}
$ARC_PERIOD=$args[0]
$date=(Get-Date -DATE "1/$ARC_PERIOD 0:0").AddMonths(1)
$dpath=Get-Date -DATE "1/$ARC_PERIOD 0:0" -format yMM
New-Item -type d "$ARC_BASE_DIR\log" -force | Out-Null
$log_file="$ARC_BASE_DIR\log\$dpath.log"
$arc_dir="$ARC_BASE_DIR\archive\$dpath"

'Script started at ' + (Get-Date -format G).ToString() | Out-File -Filepath $log_file -Append
get-childitem –Path $SOURCE_DIR -recurse -include $FILTER | 
Where-object {$_.FullName -NotMatch $EXCLUDE} |
Where-Object {$_.CreationTime –lt $date} |
%{ 
	$dest = $_.FullName -replace [regex]::escape($SOURCE_DIR), $arc_dir
	$_.CreationTime.ToString() + ' | ' + $_.FullName + ' -> ' + $dest | Write-Output	
	New-Item -type f $dest -force | Out-Null
	Move-Item $_.FullName -destination $dest -force	
	(Get-ChildItem $dest).CreationTime = $_.CreationTime 
} | Out-File -Filepath $log_file -Append
'Script completed at ' + (Get-Date -format G).ToString() | Out-File -Filepath $log_file -Append

