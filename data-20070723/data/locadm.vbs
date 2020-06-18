On Error Resume Next
'###################
'Скрипт предназначен для сменя пароля локального админа
'###################
'Новый пароль для локального админа
strNewPassword = "тут должен быть пароль"
'Возможные имена локальных админов
admNames = Array ("администратор", "administrator", "sadmin")
enableAccountName = "sadmin"
useEnableAccountName = False 'используется если локальная учетка адмна заблокирована и нужно её включить




















































Set objNetwork = CreateObject("Wscript.Network")
strComputer = objNetwork.ComputerName

Set objDomain = GetObject("WinNT://" & strComputer)
objDomain.Filter = Array("User")
	found = false
For Each admName in admNames
	for each objUser In objDomain
		If LCase(objUser.Name) = admName Then
			objUser.SetPassword (strNewPassword)
			found = true
			if err.number <> 0 then
				wscript.echo "ошибка: " & err.number & " " & err.description
			else
				wscript.echo "пароль " & objUser.Name & " установлен"
			end If
			If useEnableAccountName then
				If LCase(objUser.Name) = enableAccountName Then
					objUser.AccountDisabled = false
					objUser.SetInfo
					if err.number <> 0 then
						wscript.echo "ошибка: " & err.number & " " & err.description
					else
						wscript.echo "учетка " & objUser.Name & " разблокирована"
					end if
				end If
			End if
		End If
	next		
Next
If Not found Then
		wscript.echo "имя локального администратора не найдено"
End if
