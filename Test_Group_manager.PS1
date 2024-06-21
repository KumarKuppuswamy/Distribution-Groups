[array]$out = $null

$groups = Get-Distributiongroup -ResultSize Unlimited | select -First 5
ForEach($group in $groups)
{
	$members = Get-Distributiongroupmember -identity "$group" -ResultSize Unlimited -EA Silentlycontinue
	$count = $members.count
	$Man = Get-DistributionGroup -identity "$group" -ResultSize Unlimited | Select-object @{label="ManagedBy";expression={[string]($_.managedby | Get-user | Select-Object -ExpandProperty samaccountname) -join '|'}}
	
	$obj1 = New-object PSobject
		$obj1 | Add-Member -MemberType NoteProperty -Name "GroupType" -Value $group.GroupType
		$obj1 | Add-Member -MemberType NoteProperty -Name "SamAccountName" -Value $group.SamAccountName
		$obj1 | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $group.DistinguishedName
		$obj1 | Add-Member -MemberType NoteProperty -Name "PrimarySMTPaddress" -Value $group.PrimarySMTPaddress
		$obj1 | Add-Member -MemberType NoteProperty -Name "Count" -Value $count
		$obj1 | Add-Member -MemberType NoteProperty -Name "Created" -Value $group.WhenCreated
		$obj1 | Add-Member -MemberType NoteProperty -Name "Owner" -Value $man.Managedby
		
	    $out += $obj1
	}	 

$out 




