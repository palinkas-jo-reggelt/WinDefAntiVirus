<#

.SYNOPSIS
	WinDefAntiVirus for hMailServer

.DESCRIPTION
	Powershell script to use Windows Defender as custom virus scanner in hMailServer. 
	
	Script runs Windows Defender single-file scan on messages passed to it from hMailServer.
	
	Script mitigates false positives and offers detailed logging.

.FUNCTIONALITY
	Enter into hMailServer Admin Console > Settings > Anti-virus > External virus scanner > Scanner executable:
		Powershell -File "C:\path\to\script\WinDefAntiVirus.ps1" "%FILE%"
		
	Enter "2" for Return value 

.PARAMETER FileToScan
	The file to scan passed from hMailServer while calling this script.
	
.NOTES
	Discussion topic at hMailServer forum:
	https://hmailserver.com/forum/viewtopic.php?f=9&t=40635

#>

Param(
	[String]$FileToScan
)

<###   VARIABLES   ###>

$LogFolder = "C:\hMailServer\Logs"               # Location of hMailServer Log folder
$ScanTries = 3                                   # Number of times to try scanning file before giving up

<###   FUNCTIONS   ###>

Function Log($Msg) {
	$LogFile = "$LogFolder\WinDefAntiVirus.log"
	If (-Not(Test-Path $LogFile)) {New-Item $LogFile -Force}
	Write-Output "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')) : $Msg" | Out-File $LogFile -Append -Encoding ASCII
}

Function Execute-Command ($CommandPath, $CommandArguments) {
	# https://stackoverflow.com/questions/8761888/capturing-standard-out-and-error-with-start-process
	$pinfo = New-Object System.Diagnostics.ProcessStartInfo
	$pinfo.FileName = $CommandPath
	$pinfo.RedirectStandardError = $True
	$pinfo.RedirectStandardOutput = $True
	$pinfo.UseShellExecute = $False
	$pinfo.Arguments = $CommandArguments
	$p = New-Object System.Diagnostics.Process
	$p.StartInfo = $pinfo
	$p.Start() | Out-Null
	$p.WaitForExit()
	[pscustomobject]@{
		StdOut = $p.StandardOutput.ReadToEnd()
		StdErr = $p.StandardError.ReadToEnd()
		ExitCode = $p.ExitCode
	}
}

Function Ordinal ($Integer) {
	$Return = Switch ($Integer) {
		1 {"first"; Break}
		2 {"second"; Break}
		3 {"third"; Break}
		4 {"fourth"; Break}
		Default {$Integer; Break}
	}
	Return $Return 
}

<###   START SCRIPT   ###>

If ([String]::IsNullOrEmpty($FileToScan)) {
	Log "[ERROR] : No file argument presented : Quitting"
	Exit 0
}

$IterateScan = 0

If (Test-Path $FileToScan) {
	Do {
		Try {
			$WinDef = Execute-Command -CommandPath "$Env:ProgramW6432\Windows Defender\MpCmdRun.exe" -CommandArguments "-Scan -ScanType 3 -File ""$FileToScan"" -DisableRemediation"
		}
		Catch {
			Log "[ERROR] : $FileToScan : Error running Windows Defender command : $($Error[0])"
			Exit 0
		}

		If ($WinDef.ExitCode -eq 0) {
			If ($IterateScan -gt 0) {
				Log "[CLEAN] : $FileToScan : Clean scan on $(Ordinal ($IterateScan + 1)) scan : Exit code $($WinDef.ExitCode)"
			}
			Exit 0
		}

		If (-not([String]::IsNullOrEmpty($WinDef.StdErr))) {
			Log "[ERROR] : $FileToScan : Error on $(Ordinal ($IterateScan + 1)) scan : $($WinDef.StdErr)"
		}

		If (-not([String]::IsNullOrEmpty($WinDef.ExitCode))) {
			$VirusName = ($WinDef.StdOut | Select-String -Pattern "(?<=\sVirus:).*").Matches.Value 
			$VirusName = $VirusName -Replace "[\n\r]+",""
			If (-not([String]::IsNullOrEmpty($VirusName))) {
				Log "[VIRUS] : $FileToScan : VIRUS FOUND! $VirusName : Found on $(Ordinal ($IterateScan + 1)) scan : Exit code $($WinDef.ExitCode)"
				Exit $WinDef.ExitCode
			} Else {
				If ($IterateScan -lt ($ScanTries - 1)) {
					Log "[VIRUS] : $FileToScan : Probable error on $(Ordinal ($IterateScan + 1)) scan : Exit code $($WinDef.ExitCode) : Trying again"
				} Else {
					Log "[VIRUS] : $FileToScan : Probable error on $(Ordinal ($IterateScan + 1)) scan : Exit code $($WinDef.ExitCode) : Giving up : Exit as clean"
					Exit 0
				}
			}
		} Else {
			Log "[ERROR] : $FileToScan : No exit code on $(Ordinal ($IterateScan + 1)) scan : Quitting"
			Exit 0
		}

		Start-Sleep -Seconds 1
		$IterateScan++
	} Until ($IterateScan -eq $ScanTries)

} Else {
	Log "[NOFND] : $FileToScan : File could not be found : Quitting"
	Exit 0
}

Log "[ERROR] : $FileToScan : You should never see this message : Notify administrator"
Exit 0
