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
	Write-Output "$((Get-Date).ToString('yy-MM-dd HH:mm:ss.fff')) : $Msg" | Out-File $LogFile -Append -Encoding ASCII
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

Function Confirm-FileNotLocked ($FileToScan) {
	# https://stackoverflow.com/questions/24992681/powershell-check-if-a-file-is-locked
	Try {
		$x = [System.IO.File]::Open($FileToScan, 'Open', 'Read')
		$x.Close()
		$x.Dispose()
		Return $True
	}
	Catch [System.Management.Automation.MethodException] {
		Return $False
	}
}

<###   START SCRIPT   ###>

If ([String]::IsNullOrEmpty($FileToScan)) {
	Log "[ERROR] : No message file argument presented : Quitting"
	Exit 0
}

$IterateFileLocked = 0
$IterateScan = 0

If (Test-Path $FileToScan) {
	Do {
		If (Confirm-FileNotLocked $FileToScan) {
			Do {
				Try {
					$WinDef = Execute-Command -CommandPath "$Env:ProgramW6432\Windows Defender\MpCmdRun.exe" -CommandArguments "-Scan -ScanType 3 -File ""$FileToScan"" -DisableRemediation"
				}
				Catch {
					Log "[ERROR] : $FileToScan : Error running Windows Defender command: $($Error[0])"
					Exit 0
				}

				If ($WinDef.ExitCode -eq 0) {
					If ($IterateScan -gt 0) {
						Log "[CLEAN] : $FileToScan : Clean message on scan # $($IterateScan +1) - Exit code $($WinDef.ExitCode)"
					}
					Exit 0
				}

				If (-not([String]::IsNullOrEmpty($WinDef.StdErr))) {
					Log "[ERROR] : $FileToScan : Error on scan # $($IterateScan + 1) - $($WinDef.StdErr)"
				}

				If (-not([String]::IsNullOrEmpty($WinDef.ExitCode))) {
					Log "[VIRUS] : $FileToScan : Exit code $($WinDef.ExitCode) on scan # $($IterateScan + 1)"
				} Else {
					Log "[ERROR] : $FileToScan : No exit code on scan # $($IterateScan + 1) - Quitting"
					Exit 0
				}

				Start-Sleep -Seconds 1
				$IterateScan++

			} Until ($IterateScan -eq $ScanTries)

			$VirusName = ($WinDef.StdOut | Select-String -Pattern "(?<=\sVirus:).*").Matches.Value 
			$VirusName = $VirusName -Replace "[\n\r|\r\n|\r|\n]",""
			Log "[VIRUS] : $FileToScan : VIRUS FOUND! $VirusName"
			Exit $WinDef.ExitCode
		}

		If ($IterateFileLocked -lt ($ScanTries - 1)) {
			Log "[FLOCK] : $FileToScan : File LOCKED on try # $($IterateFileLocked + 1)! Trying again"
		} Else {
			Log "[FLOCK] : $FileToScan : File LOCKED on last try! Quitting"
			Exit 0
		}

		Start-Sleep -Seconds 1
		$IterateFileLocked++

	} Until ($IterateFileLocked -eq $ScanTries)

} Else {
	Log "[NOFND] : $FileToScan : File could not be found! Quitting"
	Exit 0
}