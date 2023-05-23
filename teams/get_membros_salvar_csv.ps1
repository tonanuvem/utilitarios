Connect-MicrosoftTeams


# Verificar grouid com o comando abaixo:
#Get-Team -User pf1178@fiap.com.br

#groupid
#58a898fb-9b9c-47f5-ae26-f3efdb901f83 2305SHIFT - The W
$groupid = '58a898fb-9b9c-47f5-ae26-f3efdb901f83'
$nomeArquivo = '11TWDK.csv'

#Get team
$Team = Get-Team -GroupId $groupid
#Get team users
$TeamUsers = Get-TeamUser -GroupId $groupid
 
#Iterate users one by one and add to the result array
$Result = @()
ForEach ($TeamUser in $TeamUsers)
{
    #Add user info to the result array
    $Result += New-Object PSObject -property $([ordered]@{
        #TeamName = $Team.DisplayName
        #TeamVisibility = $Team.Visibility
        #UserName = $TeamUser.Name
        #UserPrincipalName = $TeamUser.User
        email = $TeamUser.User
        role = $TeamUser.Role
        #GroupId = $Team.GroupId
        #UserId = $Team.UserId
    })
}

Write-Host 'Exportando IDs dos usuarios do canal'

#Export the result to CSV file
$Result | Export-CSV $nomeArquivo -NoTypeInformation -Encoding UTF8

Write-Host 'FIM !!'
