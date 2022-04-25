Connect-MicrosoftTeams

$groupid = 'fb2a5d75-bf78-4540-a8e2-8da96f0161ba'

$turma = '19abd'
$arquivo = $turma + '.csv'

$labs = $turma + '-lab'
$produtor = $labs + '-produtor'
$consumidor = $labs + '-consumidor'
$monitoracao = $labs + '-monitoracao'
$webhook = $labs + '-webhook'

Import-Csv -Path $arquivo | foreach{Add-TeamUser -GroupId $groupid -user $_.email -role 'Member'}

New-TeamChannel -GroupId $groupid -DisplayName $labs -MembershipType Private
New-TeamChannel -GroupId $groupid -DisplayName $produtor -MembershipType Private
New-TeamChannel -GroupId $groupid -DisplayName $consumidor -MembershipType Private
New-TeamChannel -GroupId $groupid -DisplayName $monitoracao -MembershipType Private
New-TeamChannel -GroupId $groupid -DisplayName $webhook -MembershipType Private

Get-TeamChannel -GroupId $groupid

Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $labs -user $_.email}
Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $produtor -user $_.email}
Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $consumidor -user $_.email}
Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $monitoracao -user $_.email}
Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $webhook -user $_.email}
