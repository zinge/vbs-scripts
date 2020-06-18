'Запускать от имени активного админа, см. accMng.bat
SearchADC = "domain.name.ru:389"
SearchCN = "dc=domain,dc=name,dc=ru"






















Function enableAccount(svitchArg)
	strUserName = InputBox("sAMAccountName пользователя:")
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = _
			  "<LDAP://" & SearchADC & "/" & SearchCN & ">;" & _
			  "(&(objectCategory=person)(objectClass=user)" & _
			  "(sAMAccountName=" & strUserName & "));" & _
			  "distinguishedName;subtree"
	Set objRecordSet = objCommand.Execute
	distUserName = objRecordSet.Fields("distinguishedName").Value
	objConnection.Close
	Set objUser = GetObject _
		("LDAP://" & SearchADC & "/" & distUserName)
	If svitchArg then
		wscript.echo "User LDAP://" & SearchADC & "/" & distUserName & " blocked" 
		objUser.AccountDisabled = svitchArg + 2
		objUser.SetInfo
	Else 
		objUser.AccountDisabled = svitchArg
		objUser.IsAccountLocked = svitchArg
		objUser.SetInfo
		wscript.echo "User LDAP://" & SearchADC & "/" & distUserName & " enabled"
	End If
End Function

boxVar = MsgBox ("Заблокировать аккаунт:		YES" & Chr(13) & _
	  Chr(10) & "Активировать аккаунт:		NO" & Chr(13) & _
	  Chr(10) & "Ничиво ни делать:		Chancel" _
	  , 3, "User block|unblock")
Select Case boxVar
   case 6      enableAccount(True)
   case 7      enableAccount(False)
   case 2     WScript.Echo "Ok, go away"
End Select