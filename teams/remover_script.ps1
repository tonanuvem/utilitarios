Connect-MicrosoftTeams


# Verificar grouid com o comando abaixo:
#Get-Team -User pf1178@fiap.com.br

# grouid
#fb2a5d75-bf78-4540-a8e2-8da96f0161ba Hands-On - Prof... 
#8cf1d42e-933f-4b80-9131-bb37825418c5 2022b - Hands-On...
#62c739a1-e669-4fb7-aa95-b994ed843e16 2023 - Hands-On P...

$groupid = 'fb2a5d75-bf78-4540-a8e2-8da96f0161ba'

$turma = '42SCJ'
$arquivo = $turma + '.csv'

$labs = $turma + '-lab'
$produtor = $labs + '-produtor'
$consumidor = $labs + '-consumidor'
$monitoracao = $labs + '-monitoracao'
$webhook = $labs + '-webhook'

Write-Host 'Importando arquivo CSV para excluir os usuarios'

Import-Csv -Path $arquivo | foreach{Remove-TeamUser -GroupId $groupid -user $_.email -role 'Member'}

Write-Host 'Removendo canais'

Remove-TeamChannel -GroupId $groupid -DisplayName $labs
Remove-TeamChannel -GroupId $groupid -DisplayName $produtor
Remove-TeamChannel -GroupId $groupid -DisplayName $consumidor
Remove-TeamChannel -GroupId $groupid -DisplayName $monitoracao
Remove-TeamChannel -GroupId $groupid -DisplayName $webhook

Get-TeamChannel -GroupId $groupid

Write-Host 'FIM !!'
