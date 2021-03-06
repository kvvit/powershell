Invoke-SqlCmd
cls
$startdate = Get-Date
Write-Host 'Начало в' $startdate
$day = Get-Date -UFormat "%d.%m.%Y"
$serversql = 'SRVEDUCATION'
$ts = Read-Host 'Введите дату...'
$n = 0
Set-Location SQLSERVER:\SQL
$arrays = Invoke-Sqlcmd -Query "SELECT  [Факультет] as company,[Уровень] as department FROM [Priem2018].[dbo].[AbitsForLogin] GROUP BY [Факультет],[Уровень] HAVING COUNT(*) > 1;" -ServerInstance $serversql
foreach($array in $arrays)
{
	$fcompany = $array.company
	$fdepartment = $array.department
	$a = Invoke-Sqlcmd -Query "SELECT [Фамилия] as sn,[Имя] as givenName,[Отчество] as middleName,[Дата рожд] as bd,[Факультет] as company,[Уровень] as department,[AbitsForLogin].[Login],[AbitsForLogin].[Email],[AbitsForLogin].[Password],[TimeStamp] FROM [Priem2018].[dbo].[AbitsForLogin] INNER JOIN [Priem2018].[ed].[ADUserData] ON [AbitsForLogin].[AbiturientId] = [ADUserData].[AbiturientId] WHERE [Факультет] = '$fcompany' And [Уровень] = '$fdepartment' And ([AbitsForLogin].[Login] IS NOT NULL OR [AbitsForLogin].[Login] != '') And [TimeStamp] > '$ts' ;" -ServerInstance $serversql
	if($a)
	{
		Invoke-Sqlcmd -Query "SELECT [Фамилия] as sn,[Имя] as givenName,[Отчество] as middleName,[Дата рожд] as bd,[Факультет] as company,[Уровень] as department,[AbitsForLogin].[Login],[AbitsForLogin].[Email],[AbitsForLogin].[Password],[TimeStamp] FROM [Priem2018].[dbo].[AbitsForLogin] INNER JOIN [Priem2018].[ed].[ADUserData] ON [AbitsForLogin].[AbiturientId] = [ADUserData].[AbiturientId] WHERE [Факультет] = '$fcompany' And [Уровень] = '$fdepartment' And ([AbitsForLogin].[Login] IS NOT NULL OR [AbitsForLogin].[Login] != '') And [TimeStamp] > '$ts' ORDER BY [Фамилия] ;" -ServerInstance $serversql | Export-Csv -Path E:\Scripts\POWERSHELL\Priem2018\Sergey\"$day"_"$fcompany"_"$fdepartment".csv -Append -Delimiter ";" -Encoding "UTF8" -NoTypeInformation
	}
	#Write-Host 'Факультет' $array.company 'уровень' $array.department
	$n++
}
#Invoke-Sqlcmd -Query "SELECT [Фамилия] as sn,[Имя] as givenName,[Отчество] as middleName,[Дата рожд] as bd,[Факультет] as company,[Уровень] as department,[Login],[Email],[Password] FROM [Priem2018].[dbo].[AbitsForLogin] ;" -ServerInstance $serversql | Export-Csv -Path E:\Scripts\POWERSHELL\Priem2018\"$day"_export.csv -Append -Delimiter ";" -Encoding "UTF8" -NoTypeInformation
Write-Host 'Totally' $n
cd E: