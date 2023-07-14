# hMailServer WinDefAntiVirus
 Use Windows Defender Antivirus as external scanner for hMailServer
 
 Discussion thread: https://hmailserver.com/forum/viewtopic.php?f=9&t=40635

 This script will mitigate false positives caused by using Windows Defender as external scanner by re-scanning files that test positive. It also logs scan activity that does not immediately come back clean.

# Instructions
 Fill in the variables at the top of the script

 Enter into hMailServer Admin Console > Settings > Anti-virus > External virus scanner > Scanner executable:

 Powershell -File "C:\path\to\script\WinDefAntiVirus.ps1" "%FILE%"
		
 Enter "2" for Return value 
 