#requires -version 4.0
#requires –runasadministrator

#
#v-nagaur: email me for a newer version, or just get it from public directory of my workstation:
# \\corp-dsktp--051\PublicShare_eyewink_\Powershell\ESCIEDisable

#INPUT: None; just execute it as admin on your VM

# to enable powershell script execution
#
# please execute the following in a powershell window(run->"powershell"-> <enter>)
#
# Set-ExecutionPolicy RemoteSigned
#



#How to use it:
#1) run powershell as administrator
#2) drag this scipt in to the window,
#3) hit enter,
#4) pray. and if you truly have administrative privileges, 
#   your OS is somewhere near "Server 2012 R2", and your heart is pure, good things will happen.
#################################################################################################


function Disable-IEESC
{
#registery keys that we need to set
#specific to Windows Server 2012 R2
$AdminKey = “HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}”
$UserKey = “HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}”
    try { #to disable 
        Set-ItemProperty -Path $AdminKey -Name “IsInstalled” -Value 0
        Set-ItemProperty -Path $UserKey -Name “IsInstalled” -Value 0
        Stop-Process -Name Explorer
        Write-Host “IE Enhanced Security Configuration (ESC) has been disabled.” -ForegroundColor Green
        }
    catch{
        Write-Host “Failed to disable IE Enhanced Security Configuration (ESC).” -ForegroundColor Red
        
        }
        
    
}


$old_ErrorActionPreference = $ErrorActionPreference
#let's make sure we don't go "silently in to the night!" //enabling stop behaviour on error
$ErrorActionPreference = 'Stop'

#the sweet spot, 
Disable-IEESC
#, where all the magic happens

#finally, back to how things were. :|
$ErrorActionPreference = $old_ErrorActionPreference 



