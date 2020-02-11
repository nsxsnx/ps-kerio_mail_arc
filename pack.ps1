# 7z packing of Kerio Connect mail archive

$ARC_BASE_DIR="C:\mail_arc"
$7Z_EXE="C:\Program Files\7-Zip\7z.exe"
$7Z_PARMS="a -mx=7 -sdel"

$date=Get-Date -format yMMdd_hhmmss
$log_file="$ARC_BASE_DIR\log\pack_$date.log"
$arc_dir="$ARC_BASE_DIR\archive\$dpath"
New-Item -type d "$ARC_BASE_DIR\log" -force | Out-Null

'Script started at ' + (Get-Date -format G).ToString() | Out-File -Filepath $log_file -Append

# loop through "archive\":
get-childitem –Path $arc_dir -Directory |
%{
	# loop through "archive\YYMM\":
	get-childitem –Path $_.FullName -Directory |
	%{
		$src = $_.FullName
		$arc = $_.FullName + '.7z'
		$cmd = '"' + $7Z_EXE + '" ' + "$7Z_PARMS $arc $src"
		Write-Host $_.FullName
		Write-Output "================================================================"
		Write-Output $cmd
		Write-Output "================================================================"
		iex "& $cmd"
	}
	
} | Out-File -Filepath $log_file -Append

'Script completed at ' + (Get-Date -format G).ToString() | Out-File -Filepath $log_file -Append