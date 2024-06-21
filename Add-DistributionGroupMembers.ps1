


$Members = import-csv "C:\WTW Exchange\Files\KumarMembers.csv"

ForEach ($Member in $Members) {
    $Identity = $Member.Dname
    $User = $Member.Users

    Write-Host "Adding $user to $Identity"
    Add-DistributionGroupMember -Identity $Identity -Member $User
    Exit
}
