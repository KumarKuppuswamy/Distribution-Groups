

[array]$Actobj = $null
[array]$out = $null
[array]$d = $null


$email = Read-Host "Enter the Email address :"
if(!$email)
{
Write-Host "No Information"
}Else
	{
			$mbx = Get-Recipient "$email" 
				if($mbx.RecipientTypeDetails -eq "UserMailbox")
				{
					$mbs = Get-Mailbox $mbx.PrimarySMTPAddress 
					$upn = $mbs.UserPrincipalName
					$hid = $mbs.HiddenFromAddressListsEnabled
					$addr = $mbs.PrimarySMTPAddress
					
						if($upn -like "*@willistowerswatson.com*")
							{
							$accs = Get-ADUser -Filter {UserPrincipalName -eq $upn} -Properties * -Server INT.DIR.WILLIS.COM | select Password*,Enabled,userAccountControl,userprincipalname,@{name="Accounthistory";expression={$_.Accountnamehistory -join ";"}},extensionAttribute6,Office,Mail
							
							
								$Actobj = New-Object PSObject
									$Actobj | Add-Member NoteProperty -Name "PrimarySMTPAddress" -Value $addr
									$Actobj | Add-Member NoteProperty -Name "Account Enabled Status" -Value $accs.Enabled
									$Actobj | Add-Member NoteProperty -Name "PasswordExpired" -Value $accs.PasswordExpired
									$Actobj | Add-Member NoteProperty -Name "PasswordLastSet" -Value $accs.PasswordLastSet
									$Actobj | Add-Member NoteProperty -Name "PasswordNeverExpires" -Value $accs.PasswordNeverExpires
									$Actobj | Add-Member NoteProperty -Name "userAccountControl" -Value $accs.userAccountControl
									$Actobj | Add-Member NoteProperty -Name "userprincipalname" -Value $accs.userprincipalname
									$Actobj | Add-Member NoteProperty -Name "AccountNameHistory" -Value $accs.Accounthistory
									$Actobj | Add-Member NoteProperty -Name "extensionAttribute6" -Value $accs.extensionAttribute6
									$Actobj | Add-Member NoteProperty -Name "Office" -Value $accs.Office
									$Actobj | Add-Member NoteProperty -Name "Mail" -Value $accs.mail	
									$Actobj | Add-Member NoteProperty -Name "HiddenFromAddressLists" -Value $hid
									$Acct += $Actobj				
									
							
							}
							Elseif($upn -like "*@towerswatson.com*")
							{
								$accs = Get-ADUser -Filter {UserPrincipalName -eq $upn} -Properties * -Server INTERNAL.TOWERSWATSON.COM | select Password*,Enabled,userAccountControl,userprincipalname,@{name="Accounthistory";expression={$_.Accountnamehistory -join ";"}},extensionAttribute6,Office,mail
								
								
								$Actobj = New-Object PSObject
									$Actobj | Add-Member NoteProperty -Name "PrimarySMTPAddress" -Value $addr
									$Actobj | Add-Member NoteProperty -Name "Account Enabled Status" -Value $accs.Enabled
									$Actobj | Add-Member NoteProperty -Name "PasswordExpired" -Value $accs.PasswordExpired
									$Actobj | Add-Member NoteProperty -Name "PasswordLastSet" -Value $accs.PasswordLastSet
									$Actobj | Add-Member NoteProperty -Name "PasswordNeverExpires" -Value $accs.PasswordNeverExpires
									$Actobj | Add-Member NoteProperty -Name "userAccountControl" -Value $accs.userAccountControl
									$Actobj | Add-Member NoteProperty -Name "userprincipalname" -Value $accs.userprincipalname
									$Actobj | Add-Member NoteProperty -Name "AccountNameHistory" -Value $accs.Accounthistory
									$Actobj | Add-Member NoteProperty -Name "extensionAttribute6" -Value $accs.extensionAttribute6
									$Actobj | Add-Member NoteProperty -Name "Office" -Value $accs.Office
									$Actobj | Add-Member NoteProperty -Name "Mail" -Value $accs.mail	
									$Actobj | Add-Member NoteProperty -Name "HiddenFromAddressLists" -Value $hid
									$Acct += $Actobj				
									
							
							}
							Else
							{
							 Write-Host "Check Manaully"
							}
							$Acct
							
							Write-Host "Checking DDL for the user - $email"
							
								ForEach($Act in $Acct)
								{
									$fil = $Act.Office
								
									$ddls = Get-DynamicDistributionGroup -ResultSize Unlimited -Filter "RecipientFilter -like '*$fil*'"
										
										ForEach($ddl in $ddls)
										{
											$primary = $ddl.PrimarySMTPAddress
													
													$ddlmember = Get-DynamicDistributionGroupmember $primary -ResultSize Unlimited | where {$_.PrimarySMTPAddress -eq "$email"}
													
													If(!$ddlmember)
													{
													}
													Else	
													{
													
														$obj = New-Object PSobject
														$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $ddl.DisplayName
														$obj | Add-Member -MemberType NoteProperty -Name "PrimarySMTPaddress" -Value $ddl.PrimarySMTPaddress
														##$obj | Add-Member -MemberType NoteProperty -Name "RecipientFilter" -Value $ddl.RecipientFilter
														$out += $obj
													}
										}
								}
							$out
						
						    Write-Host "Distribution Group details - $email"
								$dn = $mbs.DistinguishedName
								$Filter = "Members -like '$dn'"
								$dists = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter
								ForEach($dist in $dists)
								{
									
									$obj1 = New-Object PSobject
									$obj1 | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $dist.DisplayName
									$obj1 | Add-Member -MemberType NoteProperty -Name "Name" -Value $dist.Name
									$obj1 | Add-Member -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $dist.PrimarySMTPAddress
									$d += $obj1
								}
								$d | FT
	
				}Else
					{
						Write-Host "Shared Mailbox"
					}
	}
