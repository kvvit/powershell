#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.17
# Created on:   02.07.2013 15:09
# Created by:   valek
# Organization: SPbU
# Filename:  создание учетной записи компьютера в домене ad.pu.ru
#========================================================================
cls
[int]$m = 0
[int]$l = 0
[int]$j = 0
$y = $null
$p = $null
$OU = $null
$number = $null
$ad = 'created by new script'
Import-Module activedirectory
cd AD:
[int]$number = Read-Host 'Введите, сколько учетных записей Вы хотите создать (от 1 до 30)...' 
if (($number -eq $null) -or ($number -eq '') -or ($number -match '\D+'))
{
	Write-Host "Вы не ввели число учетных записей, запустите скрипт заново"
	Write-Host "Press any key to continue ..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	Exit
}
elseif ($number -gt "30")
{
	Write-Host "Введено число учетных записей больше 30, ошибка, запустите скрипт заново"
	Write-Host "Press any key to continue ..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	Exit
}
else
{
	$array = Get-ADComputer -searchbase "DC=ad,DC=pu,DC=ru" -properties Name -Filter * | where {$_.name -match "CR[0-9]{5}"}
	$array = $array.Name
	$array = ForEach ($item in $array){$item -replace '\D+'}
	$array = ForEach($item in $array){[int]$item}
	$val = 1
	[System.Collections.ArrayList]$comparray = @()
	[System.Collections.ArrayList]$ArrayList = $array
	[System.Collections.ArrayList]$AGroup = @()
	[System.Collections.ArrayList]$AOUnit = @()
	[System.Collections.ArrayList]$Countou = @()
    $UserName = [Environment]::UserName
	cd AD:
	$ADGroups = Get-ADUser $UserName -Properties *  
	$Groups = $ADGroups.MemberOf
	$usermail = $ADGroups.mail
	ForEach ($Group in $Groups)
		{
			$Group = get-adgroup -Identity $Group -Properties * 
            If (($Group.cn -like "admins_*") -or ($Group.CN -eq "Техническая поддержка УСИТ"))
				{
					$Path = $Group.DistinguishedName
					$Gname = $Group.CN
					$PathComp = $Path -replace "CN=$Gname,"
					$bases = Get-ADOrganizationalUnit -SearchScope OneLevel -searchbase "OU=СПбГУ,DC=ad,DC=pu,DC=ru" -filter {(ou -like "*факультет*") -or (ou -eq "БИО") -or (ou -like "*школа*") -or (ou -like "*гимназия*") -or (ou -like "*колледж*") -or (ou -like "*научн*") -or (ou -like "*издательство*") -or (ou -like "*ректорат*") -or (ou -like "*кафедра*") -or (ou -like "*институт*")} -Properties DistinguishedName, Name
					ForEach ($base in $bases)
					{
						$basereg = "^.*" + $base.DistinguishedName + "$"
						$y = $Group -match $basereg
						If ($y)
						{
							$PathComp = $base.DistinguishedName
							$OUnit = $base.Name
							$AOUnit.add($OUnit) | Out-Null
							$AGroup.add($PathComp) | Out-Null
							$Countou.add($j) | Out-Null
							$j++
						}
					}
				}
        }
	If ($j -gt 1)
	{
		Write-Host "Вы являетесь администратором в следующих подразделениях: "
		While ($l -lt $j)
		{
			Write-Host $AOUnit[$l] " номер: " $l
			$l++
		}
	}
	While ($val -le $number)
	{
	$i=1
	Do {
	    $i++
	    }
	While ($ArrayList -contains $i)
	$ArrayList.add($i) | Out-Null
	$ndmax = "{0:00000}" -f $i
	$ncomp = 'CR' + $ndmax
	$Profile = Get-ChildItem Env:USERPROFILE
	If ($j -gt 1)
		{
            If ($OU -eq $null)
            {
		        $m = Read-Host 'Введите номер (см.выше) подразделения в котором Вы хотите создать УЗ компьютера...' 
				$OU = $AOUnit[$m]
				If ($Countou -contains $m)
				{
					If ($OU)
					{
						Write-Host 'Выбрано подразделение: ' $OU
					}
				}
				Else
				{
					Write-Host "Вы указали неправильный номер OU в котором Вы хотите создать УЗ компьютера..., запустите скрипт заново"
					Write-Host "Press any key to continue ..."
					$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
					Exit
				}
            }
				
			Start-Sleep -s 10
				$nncomp = [bool](Get-ADObject -Filter {cn -eq $ncomp})
				If ($nncomp -ne "True")
					{
						$PathComp = "OU=" + $OU + ",OU=СПбГУ,DC=ad,DC=pu,DC=ru"
						$OUC = "OU=" + $OU
						$p = [adsi]::Exists("LDAP://$PathComp")
						if ($p)
							{
								if ($AGroup -contains $PathComp)
								{
									New-ADComputer -Name $ncomp -SamAccountName $ncomp -Path $PathComp -Enabled $true
									Set-ADComputer $ncomp -Add @{adminDescription="$ad"}
									Write-Host 'имя нового компьютера - '$ncomp 'который создан в' $PathComp ' ' $val
									$val++
									$comparray.add($ncomp) | Out-Null
								}
								else
								{
									Write-Host "У Вас нет привелегий на создание УЗ компьютера в указанном Вами OU"
									Write-Host "Press any key to continue ..."
									$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
									Exit
								}
							}
							else
							{
								Write-Host "Вы ввели неправильный OU в котором Вы хотите создать УЗ компьютера..., запустите скрипт заново"
								Write-Host "Press any key to continue ..."
								$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
								Exit
							}
						}
					}
				If ($j -eq 1)
					{
					Start-Sleep -s 10
					$nncomp = [bool](Get-ADObject -Filter {cn -eq $ncomp})
					If ($nncomp -ne "True")
						{
							$OU = $AOUnit[0]
							$PathComp = $AGroup[0]
							New-ADComputer -Name $ncomp -SamAccountName $ncomp -Path $PathComp -Enabled $true
							Set-ADComputer $ncomp -Add @{adminDescription="$ad"}
							Write-Host 'имя нового компьютера - '$ncomp 'который создан в' $PathComp ' ' $val
							$val++
							$comparray.add($ncomp) | Out-Null							
						}
					}
				If ($j -eq 0)
					{
						Write-Host "Вы не состоите в группах с административными привелегиями для создания УЗ компьютеров в домене ad.pu.ru"
						Write-Host "Press any key to continue ..."
						$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
						Exit
					}
				
				
	}
#Write-Host 'попали на создание письма'
$body = "Вами добавлены в домен ad.pu.ru (подразделение: $OU) следующие компьютеры: 
$comparray "
$PSEmailServer = "mail.spb.edu"
$En = [Text.Encoding]::UTF8
$PC = get-content env:computername
Send-MailMessage -From "CompCreator <$usermail>" -To "$usermail" -Encoding $En -Subject "Создание учетной записи компьютера" -Body $body
		
Write-Host "Press any key to continue ..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
$number = $null
$comparray = $null
$array = $null
$ArrayList = $null