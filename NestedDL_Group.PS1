
#Get nested distribution group member
Write-Host "Getting nested distribution groups and its members ...."
$OutputCsv2=".\NestedDistributionGroupMembersSummaryList_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"
$OutputCsv1=".\NestedDistributionGroupMembersDetailInfo_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"
$DistributionGroupCount = 0
$Global:GroupWithNestedGroupCount = 0
$Global:NestedLevel = 0

#Functions for ExportCSV file
Function ExportCSV1
{ 
  if($Global:NestedDistributionGroup.Count -ne 0)
  {
   $Global:GroupWithNestedGroupCount++
   $DistributionGroupSummaryDetails = @{'DL Name'= $DistributionGroupName; 'No Of Nested Groups'= $Global:NestedDistributionGroup.Count;'Total Members' = $Global:TotalGroupMembers.Count; 'Nested Distribution Group'=$Global:NestedDistributionGroup -join ', ' ;'Group Members'= if($Global:TotalGroupMembers.count -eq 0){"-"}else{ $Global:TotalGroupMembers -join', '}; 'Group Members ID'= if($Global:TotalGroupMembersUPN.count -eq 0){"-"}else{ $Global:TotalGroupMembersUPN -join ', '}}
   $ExportGroupdetails1 = New-Object PSObject -Property $DistributionGroupSummaryDetails
   $ExportGroupdetails1 | Select-Object 'DL Name','No Of Nested Groups','Total Members','Nested Distribution Group','Group Members','Group Members ID' | Export-Csv -path $OutputCsv1 -NoType -Append
  }
}
function ExportCSV2
{
  $DistributionGroupDetailInfo = @{'DL Name'= $DistributionGroupName; 'Nested DL Name'= $_.DisplayName; 'No of Members in Nested Group'= $NestedDistribtionGroupMembers.Count ; 'Nested Group Members' = if($NestedDistribtionGroupMembers.count -eq 0){"-"} else{$NestedDistribtionGroupMembers.DisplayName -join ', '}; 'Nested Group Members ID' = if($NestedDistribtionGroupMembersUPN.count -eq 0){"-"}else{$NestedDistribtionGroupMembersUPN -join ', '}}
  $ExportGroupdetails2   = New-Object PSObject -Property $DistributionGroupDetailInfo
  $ExportGroupdetails2 | Select-Object 'DL Name','Nested DL Name','No of Members in Nested Group','Nested Group Members','Nested Group Members ID' | Export-Csv -path $OutputCsv2 -NoType -Append
}


#Function for getting nested distribution group and its members
function NestedDistribution ($NestedDistribtionGroups) 
{
 $Global:NestedLevel++
 $DisplayName = $_.DisplayName
 $NestedDistribtionGroups | ForEach-Object {
   if($_.RecipientType -eq "MailUniversalDistributionGroup" -or $_.RecipientType -eq "MailUniversalSecurityGroup")
   { 
    $NestedDistribtionGroupMembers = @()
    $NestedDistribtionGroupMemberUPN = @()
    $NestedDistribtionGroupMembersUPN = @()
    $Global:NestedDistributionGroup += $_.DisplayName
    $NestedDistribtionGroupMembers += ( Get-DistributionGroupMember -Identity $_.PrimarySMTPAddress -ResultSize unlimited )
    $NestedDistribtionGroupMember = ($NestedDistribtionGroupMembers | Where{$_.RecipientType -ne "MailUniversalDistributionGroup" -and $_.RecipientType -ne "MailUniversalSecurityGroup"}).Displayname
    $NestedDistribtionGroupMembers| foreach {

     if($_.PrimarySmtpAddress -ne "")
     {  
      if($_.RecipientType -ne "MailUniversalDistributionGroup" -and $_.RecipientType -ne "MailUniversalSecurityGroup")
      {
       $NestedDistribtionGroupMemberUPN += $_.PrimarySmtpAddress
      } 
       $NestedDistribtionGroupMembersUPN += $_.PrimarySmtpAddress
     }
     if($_.PrimarySmtpAddress -eq "")
     {
       if($_.RecipientType -ne "MailUniversalDistributionGroup" -and $_.RecipientType -ne "MailUniversalSecurityGroup")
       {
        $NestedDistribtionGroupMemberUPN += $_.WindowsLiveID
       }
       $NestedDistribtionGroupMembersUPN += $_.WindowsLiveID
     }
    }
    if($NestedDistribtionGroupMember.Count -ne 0)
    {
     $Global:TotalGroupMembers += $NestedDistribtionGroupMember
     $Global:TotalGroupMembersUPN += $NestedDistribtionGroupMemberUPN
    }                                                            
    #Export output with detailed info.
    if($Global:NestedLevel -eq 1)
    {
     ExportCSV2
    }

    #Avoid infinite loop
    if($DisplayName -notin $NestedDistribtionGroupMembers.DisplayName)
    { 
     NestedDistribution -NestedDistribtionGroups $NestedDistribtionGroupMembers | Where{$_.RecipientType -eq "MailUniversalDistributionGroup" -or $_.RecipientType -eq "MailUniversalSecurityGroup"}
    }
   }
 }
 $Global:NestedLevel--
}


Get-DistributionGroup "WTW CRB All Colleagues" | ForEach-Object {
 
 $Global:TotalGroupMembers = @()
 $Global:TotalGroupMembersUPN = @()
 $Global:NestedDistributionGroup = @()
 $DistributionGroupName = $_.DisplayName
 $DistributionGroupCount++
 Write-Progress -Activity "Processed DistributionGroup Count : $DistributionGroupCount" "Currently Processing Distribution Group : $DistributionGroupName"
 $GroupMember = Get-DistributionGroupMember -Identity $_.PrimarySMTPaddress -ResultSize unlimited
 NestedDistribution -NestedDistribtionGroups $GroupMember | Where{$_.RecipientType -eq "MailUniversalDistributionGroup" -or $_.RecipientType -eq "MailUniversalSecurityGroup"}
 $Global:NestedDistributionGroup = ($Global:NestedDistributionGroup -split ',' | Select-Object -Unique )
 $GroupMember = $GroupMember | Where{$_.RecipientType -ne "MailUniversalDistributionGroup" -and $_.RecipientType -ne "MailUniversalSecurityGroup"}
 if($GroupMember.count -ne 0)
 {
  $Global:TotalGroupMembers = (($Global:TotalGroupMembers + ($GroupMember).DisplayName) -split ',' | Select-Object -Unique)
   
  $GroupMember | foreach {
   if($_.PrimarySmtpAddress -ne "")
   {
     $Global:TotalGroupMembersUPN += $_.PrimarySmtpAddress
   }
   if($_.PrimarySmtpAddress -eq "")
   {
     $Global:TotalGroupMembersUPN += $_.WindowsLiveID
   }
  }
  $Global:TotalGroupMembersUPN = (($Global:TotalGroupMembersUPN) -split ',' | Select-Object -Unique)
 }
 else 
 {
  $Global:TotalGroupMembers = $Global:TotalGroupMembers -split ',' | Select-Object -Unique
  $Global:TotalGroupMembersUPN = $Global:TotalGroupMembersUPN -split ',' | Select-Object -Unique
 }
 #Export output with summary details.
 ExportCSV1
 
}



#Open output file after execution 
if($DistributionGroupCount -eq 0)
{
 Write-Host "No distribution group found in this organization"
}
else
{
 Write-Host "$DistributionGroupCount Distribution group found in this organization"
 if($Global:GroupWithNestedGroupCount -ne 0)
 {
  Write-Host "$Global:GroupWithNestedGroupCount Distribution group found with nested group"
  if((Test-Path -Path $OutputCsv1) -eq "True" -and (Test-Path -Path $OutputCsv2) -eq "True") 
  {
   Write-Host "The output files are available in the current working directory"
   Write-Host "The Summary report name  : $OutputCsv2" -ForegroundColor Green 
   Write-Host "The Detailed report name : $OutputCsv1" -ForegroundColor Green
   $Prompt = New-Object -ComObject wscript.shell    
   $UserInput = $Prompt.popup("Do you want to open output files?",` 0,"Open Output Files",4)    
   if ($UserInput -eq 6)    
   {    
    Invoke-Item "$OutputCSV1" 
    Invoke-Item "$OutputCSV2"   
   }  
  } 
 }
 else
 {
  Write-Host "But no distribution group found with nested group"
 }
}


