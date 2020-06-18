xlsFilePath = "filepath\filename"
SearchADC = "xx.xx.xx.xx:389"
domainDNSName = "domain.name.ru"
startOU = "СтартовоеОУ"
cexCellNum = 1 'номер столбца значений "цех"
departmentCellNum = 4 'номер столбца значений "отдел"
telephoneNumberCellNum = 9 'номер столбца значений "телефон"
physicalDeliveryOfficeNameCellNum = 10 'номер столбца значений "комната"
displayNameCellNum = 8 'номер столбца значений "ФИО"
lNameCellNum = 2 'номер столбца значений "город"
lNameCellNumUse = True 'использовать параметр город
streetAddressNameCellNum = 3 'номер столбца значений "улица"
streetAddressNameCellNumUse = True 'использовать параметр улица
titleCellNum = 11 'номер столбца значений "должность"
companyName = "Можно написать полное имя организации" 'или companyName = startOU название организации
controlCellNum = cexCellNum 'контрольный столбец, если ячейка будет пустой, то !!!!!!!импорт данных прикратитца!!!!
inputStartRow = 2 'начинать импорт данных с этой строки
'descriptionCellNum = 5 'номер столбца значений "описание"

'hjdvsavd6462371nfsndfb для быстрого перехода к роботу выполни поиск

Const ADS_PROPERTY_UPDATE = 2 
'#######################################
'#######################################
'#######################################
'#######################################
'#######################################
'#######################################
'#######################################


'поиск в АД User
Function returnUserName(displayName)
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = _
		"<LDAP://" & SearchADC & "/" & SearchCN & ">;" & _
			"(&(objectCategory=Person)(objectClass=user)" & _
			"(displayName="&displayName&"));" & _
			"distinguishedName;subtree"
	Set objRecordSet = objCommand.Execute
	If Not objRecordset.EOF Then
		returnUserName = objRecordSet.Fields("distinguishedName").Value	
		objConnection.Close
	End If
End Function 
'/поиск в АД User
'проверка наличия OU, при отсутствии создает, в обоих случаях вернет distinguishedName
'поиск в АД OU
Function returnOUName(startThisOU, OU, searchCN)
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = _
		"<LDAP://" & SearchADC & "/ou="&startThisOU&"," & SearchCN & ">;" & _
			"(&(objectCategory=OrganizationalUnit)(objectClass=organizationalUnit)" & _
			"(name="&OU&"));" & _
			"distinguishedName;subtree"
			'wscript.echo objCommand.CommandText
	Set objRecordSet = objCommand.Execute
	If Not objRecordset.EOF Then
		returnOUName = objRecordSet.Fields("distinguishedName").Value	
		objConnection.Close
	End If
End Function
'/поиск в АД OU
Function mngOU(upOU, OU, SearchCN)
	distUpOUName = returnOUName(upOU, OU, SearchCN)
		If distUpOUName = "" Then
			Set objDomain = GetObject("LDAP://" & SearchADC & "/ou="&upOU&"," & SearchCN)
			Set objOU = objDomain.Create("organizationalUnit", "ou="&OU)
			objOU.SetInfo
			mngOU = "ou="&OU&",ou="&upOU&"," & SearchCN
		Else
			mngOU = distUpOUName
		End If
End Function

Function myOU(upStage, stage1, stage2)
If stage1 = "" Then
	myTmpOU = "ou="&upStage&","&SearchCN
	If stage2 = "" then
		myOU = myTmpOU
	Else
		myOU = mngOU (upStage, stage2, searchCN)
	End If
Else
	myTmpOU = mngOU (upStage, stage1, searchCN)
	If stage2 = "" then
		myOU = myTmpOU
	Else
		myOU = mngOU (stage1, stage2, "ou="&upStage&","&searchCN)
	End If
End If
End Function
'/проверка наличия OU, при отсутствии создает, в обоих случаях вернет distinguishedName

Sub userAdd (cex, department, title, telephoneNumber, physicalDeliveryOfficeName, displayName, l , streetAddress)
		Set objOU = GetObject("LDAP://" & SearchADC & "/"&myOU(startOU, cex, department))
		Set objUser = objOU.Create("User", "cn= " + displayName)
		objUser.Put "displayName", displayName
		objUser.Put "givenName", displayName
		objUser.Put "userPrincipalName", FioToSama(displayName) + "@" + domainDNSName
		objUser.Put "sAMAccountName", FioToSama(displayName)
		objUser.Put "company", companyName&", "&cex
		If lNameCellNumUse And (l <> "") Then
			objUser.Put "l", l
		End If
		If streetAddressNameCellNumUse And (streetAddress <> "") Then
			objUser.Put "streetAddress", streetAddress
		End If
		If department <> "" Then
			objUser.Put "department", department
		End If
		If title <> "" Then
			objUser.Put "title", title
		End If
		If telephoneNumber <> "" then
			objUser.Put "telephoneNumber", telephoneNumber
		End If
		If physicalDeliveryOfficeName <> "" then
			objUser.Put "physicalDeliveryOfficeName", physicalDeliveryOfficeName & ""
		End if
		objUser.PutEx ADS_PROPERTY_UPDATE, "description", Array(myOU(startOU, cex, department))
		objUser.SetInfo
		objUser.AccountDisabled = True
		objUser.SetInfo
End Sub

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


Function normDisplayName (DisplayName)
	normDisplayName = ""
	arrDisplayName = Split(DisplayName, " ", -1, 1)
	For i = 0 To UBound(arrDisplayName)
		length = Len (arrDisplayName(i))	
		For j = 1 To length
			ch = mid (arrDisplayName(i), j, 1)
			If j = 1 Then
				normDisplayName = normDisplayName & UCase(ch)
			Else
				normDisplayName = normDisplayName & LCase(ch)
			End if	
		Next
		normDisplayName = normDisplayName & " "
	next
End Function

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

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    (xlsFilePath)
wscript.echo "open file = " & xlsFilePath
intRow = inputStartRow
j = 1



'hjdvsavd6462371nfsndfb   конрольная точка, для быстрого поиска робота
Do Until objExcel.Cells(intRow,controlCellNum).Value = ""
	'	wscript.echo "row" & intRow
	cex = Trim(objExcel.Cells(intRow, cexCellNum).Value)
	department = Trim(objExcel.Cells(intRow, departmentCellNum).Value)
	title = Trim(objExcel.Cells(intRow, titleCellNum).Value)
	telephoneNumber = Trim(objExcel.Cells(intRow, telephoneNumberCellNum).Value)
	physicalDeliveryOfficeName = Trim(objExcel.Cells(intRow, physicalDeliveryOfficeNameCellNum).Value)
	'   description = Trim(objExcel.Cells(intRow, descriptionCellNum).Value)
	displayName = Trim(normDisplayName(objExcel.Cells(intRow, displayNameCellNum).Value))
	lName = Trim(normDisplayName(objExcel.Cells(intRow, lNameCellNum).Value))
	streetAddressName = Trim(normDisplayName(objExcel.Cells(intRow, streetAddressNameCellNum).Value))

	'   wscript.echo "dN" & displayName
	'   wscript.echo "c" &cex
	'   wscript.echo "d" &department
	If (displayName <> "" AND displayName <> "Приемная" AND displayName <> "Диспетчер" AND displayName <> "Телемеханика" AND displayName <> "Телемеханика ЦАП" AND displayName <> "Диспетчерская") Then
	'   wscript.echo "z"
		If returnUserName(displayName) = "" then
			wscript.echo "Пользователь "&displayName&" создаетца"
			userAdd cex, department, title, telephoneNumber, physicalDeliveryOfficeName, displayName, lName, streetAddressName
			wscript.echo "Пользователь "&displayName&" создан"
			j = j + 1
		Else 
			wscript.echo "Пользователь "&displayName&" существует"
		End If
	End if
	intRow = intRow + 1
Loop

objExcel.Quit

wscript.echo "Всего создано: " &j
