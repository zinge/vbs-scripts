' Сервак на котором ищем и порт
SearchADC = "domen.name.ru:389"
' Место в котором ищем
' ДНС имя домена
domainDNSName = "domen.name.ru"
' Временная OU для новых пользователей
TranzitOU = "Transit"
'Учетные данные для подключения в домен
strDomain = "DOMEN"
strPassword = "пароль"
strUser = "пользователь"



Dim TransTable
Set TransTable = CreateObject ("Scripting.Dictionary")

TransTable.Add "а", "a"
TransTable.Add "б", "b"
TransTable.Add "в", "v"
TransTable.Add "г", "g"
TransTable.Add "д", "d"
TransTable.Add "е", "e"
TransTable.Add "ё", "e"
TransTable.Add "ж", "j"
TransTable.Add "з", "z"
TransTable.Add "и", "i"
TransTable.Add "й", "y"
TransTable.Add "к", "k"
TransTable.Add "л", "l"
TransTable.Add "м", "m"
TransTable.Add "н", "n"
TransTable.Add "о", "o"
TransTable.Add "п", "p"
TransTable.Add "р", "r"
TransTable.Add "с", "s"
TransTable.Add "т", "t"
TransTable.Add "у", "u"
TransTable.Add "ф", "f"
TransTable.Add "х", "h"
TransTable.Add "ц", "z"
TransTable.Add "ч", "ch"
TransTable.Add "ш", "sh"
TransTable.Add "щ", "sch"
TransTable.Add "ъ", ""
TransTable.Add "ы", "y"
TransTable.Add "ь", ""
TransTable.Add "э", "e"
TransTable.Add "ю", "yu"
TransTable.Add "я", "ya"

TransTable.Add " ", " "

TransTable.Add "А", "A"
TransTable.Add "Б", "B"
TransTable.Add "В", "V"
TransTable.Add "Г", "G"
TransTable.Add "Д", "D"
TransTable.Add "Е", "E"
TransTable.Add "Ё", "E"
TransTable.Add "Ж", "J"
TransTable.Add "З", "Z"
TransTable.Add "И", "I"
TransTable.Add "Й", "Y"
TransTable.Add "К", "K"
TransTable.Add "Л", "L"
TransTable.Add "М", "M"
TransTable.Add "Н", "N"
TransTable.Add "О", "O"
TransTable.Add "П", "P"
TransTable.Add "Р", "R"
TransTable.Add "С", "S"
TransTable.Add "Т", "T"
TransTable.Add "У", "U"
TransTable.Add "Ф", "F"
TransTable.Add "Х", "H"
TransTable.Add "Ц", "Z"
TransTable.Add "Ч", "Ch"
TransTable.Add "Ш", "Sh"
TransTable.Add "Щ", "Sch"
TransTable.Add "Ъ", ""
TransTable.Add "Ы", "Y"
TransTable.Add "Ь", ""
TransTable.Add "Э", "E"
TransTable.Add "Ю", "Yu"
TransTable.Add "Я", "Ya"


Function isJotVowel (letter)
    isJotVowel = (letter = "я" _
	OR letter = "и" _
	OR letter = "е" _
	OR letter = "ё" _
	OR letter = "ю")
end function


function Translit (RussianString)
  dim LatinString
  length = len (RussianString)

  for i = 1 to length
      ch = mid (RussianString, i, 1)
      if ch = "ь" AND i < length - 1 AND isJotVowel (mid (RussianString, i + 1, 1)) then
        LatinString = LatinString + "j"
      else
        LatinString = LatinString + TransTable.Item (ch)
      end if
  next 
  Translit = LatinString
end function


function FioToSama (fio)
  parts = Split (fio, " ")
  tfio = lcase (parts (0)) + left (parts (1), 1) + left (parts (2), 1)
  FioToSama = Translit (tfio)
end function


Function convertDns2LdapName(strDnsDomain)
convertDns2LdapName = ""
arrDnsDomain = Split(strDnsDomain, ".", -1, 1)
for i=0 to UBound(arrDnsDomain)
    If i = 0 Then
		convertDns2LdapName = convertDns2LdapName & "dc=" & arrDnsDomain(i)
	Else
		convertDns2LdapName = convertDns2LdapName & ",dc=" & arrDnsDomain(i)
	End If
next
End Function
SearchCN = convertDns2LdapName(domainDNSName)
Sub userAdd (fullName, sama)
        Set objOU = GetObject("LDAP://" & SearchADC & "/OU=" & TranzitOU & "," & SearchCN)
        Set objUser = objOU.Create("User", "cn= " + fullName)
        objUser.Put "displayName", fullName
        objUser.Put "givenName", fullName
        objUser.Put "userPrincipalName", sama + "@" + domainDNSName
        objUser.Put "sAMAccountName", sama
        objUser.SetInfo
        objUser.AccountDisabled = False
        objUser.SetInfo
End Sub
Sub enableAccount(tmpSAMAccountName)
	strUserName = tmpSAMAccountName
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
		objUser.AccountDisabled = false
		objUser.IsAccountLocked = false
		objUser.SetInfo
		wscript.echo "User LDAP://" & SearchADC & "/" & distUserName & " enabled"
End Sub
'------------- Выводит инфу об ошипке
Sub errMsg (srtActionName, strActionResult)
    If strActionResult = 0 Then
      wscript.echo srtActionName & " OK"
    Else
      wscript.echo srtActionName & " error" & vbCrLF & "err.Number=" & err.Number & vbCrLF & "err.description=" & err.description
      strError(0) = Hex(Err.Number)
      Set objWMIError = New SWbemLastError
      If TypeName(objWMIError) <> "Nothing" Then
        strError(1) = objWMIError.ParameterInfo
        strError(2) = objWMIError.Description
      End if
    End if
End Sub



'--------------------------------- Запрос пользователя в АД
strUserName = InputBox("Пожалуйста, введите свою Фамилию Имя Отчество:")

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"

Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

objCommand.CommandText = _
          "<LDAP://" & SearchADC & "/" & SearchCN & ">;" & _
          "(&(objectCategory=person)(objectClass=user)" & _
          "(displayName=" & strUserName & "));" & _
          "displayName,sAMAccountName;subtree"

foundp = FALSE

Set objRecordSet = objCommand.Execute
While Not objRecordset.EOF
  zSAMa = objRecordset.Fields("sAMAccountName")
  foundp = TRUE
  objRecordset.MoveNext
Wend

objConnection.Close

If not foundp Then
 zSAMa = FioToSama (strUserName) ' inputBox("Введи samAccountName:")
 userAdd strUserName, zSAMa
 wscript.echo "Пользователь " & strUserName & " | " & zSAMa & " создан успешно."
End If

 enableAccount zSAMa
'--------------------------------- /Запрос пользователя в АД