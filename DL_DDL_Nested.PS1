[array]$All = $null 
function Get-DistributionGroupMembersRecursive
	{
	    [CmdletBinding()]
	    Param (
	    # Id of Distribution Group
	    [Parameter(ValueFromPipeline=$true)]
	    [string]
	    $PrimarySmtpAddress
	    )
	

	    

	    # Process Distribution Group
	    process 
	    {
	        
			##Write-Host "Processing $PrimarySmtpAddress" -ForegroundColor Yellow
			##$Primarysmtpaddress 
			
				if($Primarysmtpaddress -like "*DDL*")
				{	
					$dist = Get-DynamicDistributionGroup -Identity $PrimarySmtpAddress -ErrorAction SilentlyContinue
					$name = $dist.DisplayName
					Write-Host "DynamiciDistribution Group - $name "
					$DistributionGroupMembers = Get-DynamicDistributionGroupMember -Identity $PrimarySmtpAddress -ResultSize Unlimited
						# Add UserMailBox Recipients Members to AllMembers
						$AllMembers = $DistributionGroupMembers | Where-Object{$_.RecipientType -eq 'UserMailBox'} | Select PrimarySMTPAddress,Customattribute6,CountryorRegion,Office
																	
						
				}Else
					{
					    $dist = Get-DistributionGroup -Identity $PrimarySmtpAddress -ErrorAction SilentlyContinue	
						$name = $dist.DisplayName
						Write-Host "Distribution Group - $name "
						# Get All Distribution Group Members
						$DistributionGroupMembers = Get-DistributionGroupMember -Identity $PrimarySmtpAddress -ResultSize Unlimited
						# Add UserMailBox Recipients Members to AllMembers
						$AllMembers = $DistributionGroupMembers | Where-Object{$_.RecipientType -eq 'UserMailBox'} | Select PrimarySMTPAddress,Customattribute6,CountryorRegion,Office
						         
					}
					    # Check for Nested Distribution Groups
						if($DistributionGroupMembers | Where-Object{$_.RecipientType -like '*Group*'})
						{
						$DistributionGroupMembers | Where-Object{$_.RecipientType -like '*Group*'} | ForEach-Object {Get-DistributionGroupMembersRecursive -PrimarySmtpAddress $_.PrimarySmtpAddress}
						}
				

	        
	    }
	

	    # Output Results
	    end 
	    {
			
			Return $AllMembers
			
	    }
	}
	