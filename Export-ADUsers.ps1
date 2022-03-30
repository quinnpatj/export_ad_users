
<# Export AD Users - Patrick Quinn, 2021

==================================================== ABOUT: ===================================================

   Script to audit all users in AD, then export three spreadsheets to Documents:
   1. All AD users, both active and inactive
   2. All ACTIVE AD users in the Non-Employees OU (Contractors)
   3. All ACTIVE AD users in the Employees OU

===============================================================================================================
#>

$date = Get-Date -Format "yyyyMMdd"

$path = $env:USERPROFILE + "\Documents\PowerShell_Output"
If (!(Test-Path $path)) {
        
    New-Item -ItemType Directory -Force -Path $path
}

# All users

$users = Get-AdUser -Filter *  | Sort-Object SamAccountName | Select-Object -ExpandProperty SamAccountName

$count = $users.Count
$i = 1

$users | foreach {

    Clear-Variable -Name user, username, name, mail, title, dept, loc, manager, mName, mEmail, contractor, active | Out-Null

    $user = Get-ADUser $_ -Properties *
    $username = $user.SamAccountName
    $name = $user.Name
    $mail = $user.mail
    $title = $user.title
    $dept = $user.Department
    $loc = $user.Office
    $manager = Get-ADUser $user.Manager -Properties *
    $mName = $manager.Name
    $mEmail = $manager.mail
    $contractor = $user.AccountExpirationDate
    $active = $user.Enabled
    
    [pscustomobject]@{
        Username =  ($username -join ', ');
        Name =  ($name -join ', ');
        Email = ($mail -join ', ');
        Title = ($title -join ', ');
        Department = ($dept -join ', ');
        Office = ($loc -join ', ');
        Manager = ($mName -join ', ');
        ManagerEmail = ($mEmail -join ', ');
        ContractExpiration = ($contractor -join ', ');
        Enabled = $active
    } | Export-Csv -Path  "$path\AD_AllUsers-$date.csv" -Append -NoTypeInformation

    Write-Progress -Activity "Checking all AD users..." -Status "Checking user $i of $count" -PercentComplete (($i / $count) * 100)
    $i++
}


# Active contractors

$sbCont = "OU=Non-employees,OU=People,OU=Accounts,DC=" + $env:USERDOMAIN + ",DC=net"
$users = Get-AdUser -SearchBase $sbCont -Filter 'enabled -eq $true' -Properties * | Sort-Object SamAccountName | Select-Object -ExpandProperty SamAccountName

$count = $users.Count
$i = 1

$users | foreach {

    Clear-Variable -Name user, username, name, mail, title, dept, loc, manager, mName, mEmail, contractor, active | Out-Null

    $user = Get-ADUser $_ -Properties *
    $username = $user.SamAccountName
    $name = $user.Name
    $mail = $user.mail
    $title = $user.title
    $dept = $user.Department
    $loc = $user.Office
    $manager = Get-ADUser $user.Manager -Properties *
    $mName = $manager.Name
    $mEmail = $manager.mail
    $contractor = $user.AccountExpirationDate
    $active = $user.Enabled
    
    [pscustomobject]@{
        Username =  ($username -join ', ');
        Name =  ($name -join ', ');
        Email = ($mail -join ', ');
        Title = ($title -join ', ');
        Department = ($dept -join ', ');
        Office = ($loc -join ', ');
        Manager = ($mName -join ', ');
        ManagerEmail = ($mEmail -join ', ');
        ContractExpiration = ($contractor -join ', ');
        Enabled = $active
    } | Export-Csv -Path  "$path\AD_ActiveContractors-$date.csv" -Append -NoTypeInformation

    Write-Progress -Activity "Checking AD for active contractors..." -Status "Checking user $i of $count" -PercentComplete (($i / $count) * 100)
    $i++
}


# Active employees

$sbEmp = "OU=Employees,OU=People,OU=Accounts,DC=" + $env:USERDOMAIN + ",DC=net"
$users = Get-AdUser -SearchBase $sbEmp -Filter 'enabled -eq $true' -Properties * | Sort-Object SamAccountName | Select-Object -ExpandProperty SamAccountName

$count = $users.Count
$i = 1

$users | foreach {

    Clear-Variable -Name user, username, name, mail, title, dept, loc, manager, mName, mEmail, contractor, active | Out-Null

    $user = Get-ADUser $_ -Properties *
    $username = $user.SamAccountName
    $name = $user.Name
    $mail = $user.mail
    $title = $user.title
    $dept = $user.Department
    $loc = $user.Office
    $manager = Get-ADUser $user.Manager -Properties *
    $mName = $manager.Name
    $mEmail = $manager.mail
    $contractor = $user.AccountExpirationDate
    $active = $user.Enabled
    
    [pscustomobject]@{
        Username =  ($username -join ', ');
        Name =  ($name -join ', ');
        Email = ($mail -join ', ');
        Title = ($title -join ', ');
        Department = ($dept -join ', ');
        Office = ($loc -join ', ');
        Manager = ($mName -join ', ');
        ManagerEmail = ($mEmail -join ', ');
        ContractExpiration = ($contractor -join ', ');
        Enabled = $active
    } | Export-Csv -Path  "$path\AD_ActiveEmployees-$date.csv" -Append -NoTypeInformation

    Write-Progress -Activity "Checking AD for active employees..." -Status "Checking user $i of $count" -PercentComplete (($i / $count) * 100)
    $i++
}

Write-Host
Write-Host "Results output to $path"
Start-Process $path