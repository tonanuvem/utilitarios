Write-Host 'Autenticando no Teams'

Connect-MicrosoftTeams

# Verificar grouid com o comando abaixo:
#Get-Team -User pf1178@fiap.com.br

# 2022a - Hands-On - Prof. Andre
#$groupid = 'fb2a5d75-bf78-4540-a8e2-8da96f0161ba'

# 2022b - Hands-On - Prof. Andre
#$groupid = '8cf1d42e-933f-4b80-9131-bb37825418c5'

# 2023 - Hands-On - Prof. Andre
$groupid = '62c739a1-e669-4fb7-aa95-b994ed843e16'

$turma = '1SCJR'
$arquivo = $turma + '.csv'

$labs = $turma + '-lab'
# $produtor = $labs + '-produtor'
# $consumidor = $labs + '-consumidor'
# $monitoracao = $labs + '-monitoracao'
# $webhook = $labs + '-webhook'

Write-Host 'Importando arquivo CSV para adicionar os usuarios'

Import-Csv -Path $arquivo | foreach{Add-TeamUser -GroupId $groupid -user $_.email -role 'Member'}

Write-Host 'Criando canais'

New-TeamChannel -GroupId $groupid -DisplayName $labs -MembershipType Private
# New-TeamChannel -GroupId $groupid -DisplayName $produtor -MembershipType Private
# New-TeamChannel -GroupId $groupid -DisplayName $consumidor -MembershipType Private
# New-TeamChannel -GroupId $groupid -DisplayName $monitoracao -MembershipType Private
# New-TeamChannel -GroupId $groupid -DisplayName $webhook -MembershipType Private

# Get-TeamChannel -GroupId $groupid

Write-Host 'Adicionando usuarios aos canais'

Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $labs -user $_.email}
# Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $produtor -user $_.email}
# Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $consumidor -user $_.email}
# Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $monitoracao -user $_.email}
# Import-Csv -Path $arquivo | foreach{Add-TeamChannelUser -GroupId $groupid -DisplayName $webhook -user $_.email}

Write-Host 'FIM !!'
