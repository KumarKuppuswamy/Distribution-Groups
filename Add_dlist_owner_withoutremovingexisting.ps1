$dls = get-content "C:\Users\Kuppuswamyku\Desktop\wtw_wr.txt"
ForEach($dl in $dls)
{
	$dists = Get-DistributionGroup -identity "$dl"
	
	if($dists.RecipientTypeDetails -eq "MailUniversalDistributionGroup")
	{
	Set-DistributionGroup "$dl" -ManagedBy @{Add='charlotte.carruthers@wtwco.com'}	
	}
	Elseif($dists.RecipientTypeDetails -eq "MailUniversalSecurityGroup")
	{
	Set-DistributionGroup "$dl" -ManagedBy @{Add='charlotte.carruthers@wtwco.com'} -BypassSecurityGroupManagerCheck
	}
	Else
	{
	Write-Host "Check Manually"
	}
}

