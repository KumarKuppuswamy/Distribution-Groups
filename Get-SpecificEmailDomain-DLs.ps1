##============================================================================== 
##==============================================================================
<# NAME...........: Get-SpecificEmailDomains-DLs.ps1
## DESCRIPTION....: The script will work through all the organizations Distribution Groups and remove any SMTP 
addresses for the SMTP Domains List
##
## AUTHOR.........: Lee Stephenson
## VERSION........: 2.0
## Created........: 16/01/2022            (UK Format)
## Last Update....:                       (UK Format)
##
## Email..........: lee@nexlynk.com
##
##
##
##                                       Disclaimer
##
## This script is provided AS IS without warranty of any kind
##
## Please make sure you have read through the script prior to using it
## and ensure that no damage can be made to your Developement or
## Production environment.
##
## USE THIS SCRIPT AT YOUR OWN RISK.
##
## Nexlynk Solutions accepts no liability for any damage caused by the use of this script.

##==============================================================================
## Notes
##==============================================================================

#>

##==============================================================================
## Parameters
##===========================================================================

## None

##==============================================================================
## Functions
##===========================================================================

## None

##==============================================================================
## Welcome
##==============================================================================

Write-Host "Removing Innovisk SMTP Domains from Distribution Groups" -ForegroundColor Cyan
Write-Host " "

##==============================================================================
## Credentails
##==============================================================================

##==============================================================================
## Date
##==============================================================================

$FileDate = Get-Date -Format dd-MM-yyyy

##==============================================================================
## Modules
##==============================================================================

Import-Module ActiveDirectory

##==============================================================================
## Static Variables
##==============================================================================

$UPN = "Put your UPN here"

##==============================================================================
## File Information
##==============================================================================

Write-Host "Preparing Files" -ForeGroundColor Magenta

# Date for LogFile
$Date = Get-Date -Format yyyy-MM-dd

$SMTPDomainsFile = "C:\WTW Exchange\Files\SMTP Domains.txt"
$CheckFile = "C:\WTW Exchange\Files\DistributionGroups.csv"
$OutputFile = "C:\WTW Exchange\Files\SpecifiEmailDomain-DLs.txt"
$OutputFileCSV = "C:\WTW Exchange\Files\SpecifiEmailDomain-DLs.csv"

If (Test-Path $OutputFile) {
    Clear-Content $OutputFile
} Else {
    New-Item $OutputFile -Type File | Out-Null
}

if (Test-Path $OutputFileCSV) {
    Remove-Item $OutputFileCSV
}

Start-Sleep -s 3
Add-Content $OutputFile "Name,Alias,Location,Domain,Email,NewPrimarySMTPAddress"

##----------------------------------------------------------
## Get Mailboxes
##----------------------------------------------------------

Write-Host "Getting Distribution Groups, Please be patient....." -ForeGroundColor Magenta
Write-Host ""

$DLs = Import-CSV $CheckFile

##-------------.--------------------------------------------
## Process Mailboxes
##----------------------------------------------------------

$DLCount = 0

$SMTPDomains = Get-Content $SMTPDomainsFile


ForEach ($DL in $DLs) {
        
    # Refresh Connection to Exchange Online to stop script failing
    $DLCount ++

    If ($DLCount -eq 500) {
         Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false
        $DLCount = 1
    }

    $Name = $DL.Name
    $Name = $Name -Replace ",",""   
    $Alias = $DL.Alias
    $Location = $DL.Location

    If ($Location -eq "O365") {
        Write-Host "Checking email addressess for - $Name" -ForeGroundColor Cyan
        $EAs =Get-DistributionGroup $Alias -ErrorAction SilentlyContinue | Select-Object -ExpandProperty EmailAddresses
    }
    If ($Location -eq "OnPrem") {
        Write-Host "Checking email addressess for - $Name" -ForeGroundColor Magenta
        $EAs =Get-OPDistributionGroup $Alias -ErrorAction SilentlyContinue | Select-Object -ExpandProperty EmailAddresses
    }
            
    ForEach($Ea in $EAs) {
        $Found = 0
        If($EA -Like "smtp:*") {
            $Email = $EA -Replace "SMTP:",""
            $Prefix = $Email.Split("@")[0]
            $Domain = $EMail.Split("@")[1]
            $Sufix = $Domain.split(".")[0]

             ForEach ($SMTPDomain in $SMTPDomains) {
                If($Domain -eq $SMTPDomain) {
                    Write-Host "$Domain" -ForegroundColor Magenta

                    Add-Content $OutputFile "$Name,$Alias,$Location,$Domain,$Email"
                    $Found = 1
                }
            }
        }
    }
    Write-Host " "
} 

Write-host "-----------"
Write-Host " "

##----------------------------------------------------------
## Convert to CSV
##----------------------------------------------------------

Rename-Item $OutputFile $OutputFileCSV

##----------------------------------------------------------
## Finish
##----------------------------------------------------------

Write-Host " "
Write-Host "Finished"
Write-Host " "

<#----------------------------------------------------------
##----------------------------------------------------------
##----------------------------------------------------------
##----------------------------------------------------------
##----------------------------------------------------------
## Version Control
##----------------------------------------------------------

Version         1.0
Date            07/01/2019
Author          Lee Stephenson
Notes           First working script

Version         1.1
Date            12/01/2019
Author          Lee Stephenson
Notes           Added Logging

Version         1.2
Date            18/01/2019
Author          Lee Stephenson
Notes           increased output to screen for error checking

Version         2.0
Date            21/01/2019
Author          Lee Stephenson
Notes           Added logic to limit number of changes per

#>
