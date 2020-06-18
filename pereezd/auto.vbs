On Error Resume Next
'==================================================================================
'                     Скрипт автоматичекой миграции                              
'==================================================================================
'----------------------------------------------------------------------------------
'                     План действия
'---------------------------------------------------------------------------------
'0) врубаем логирование
'0.main) проверка версии ОС
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'1) переменные           изменяйте только в этом блоке
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'2) функции
'3) получение данных от пользователя
'4) смена сетевых настроек
'5) поиск и разблокировка, либо создание пользователя
'6) смена пароля локального админа
'7) подготовка действий после перезагрузки
'8) включение компьютера в домен и переименование
'9) перезагрузка компьютера
'############################################################################
'############################################################################
'###################                              0) врубаем логирование
'############################################################################
'############################################################################
logFileName = "domainMigrate.log"
logFileNameDriveName = "c:"
domainMigrateLogMsg = "Перевод в новый домен начат: "& Now & VbCrLf
Set objFSO01 = CreateObject("Scripting.FileSystemObject")
Set f1 = objFSO01.CreateTextFile(logFileNameDriveName&"\"&logFileName, True)
f1.WriteLine domainMigrateLogMsg
f1.Close
Set objFSO01 = Nothing
Sub appendLogMsg (domainMigrateLogMsgTmp) 
   Const ForAppending = 8
   Dim fso, f
   Set fso02 = CreateObject("Scripting.FileSystemObject")
   Set f = fso02.OpenTextFile(logFileNameDriveName&"\"&logFileName, ForAppending, True)
   f.Write Now & "   " & domainMigrateLogMsgTmp & VbCrLf
   f.Close
   Set fso02 = nothing
End Sub
'############################################################################
'############################################################################
'###################                              \0) врубаем логирование
'############################################################################
'############################################################################
appendLogMsg "\0) врубаем логирование"
appendLogMsg "0.main) проверка версии ОС"
'############################################################################
'############################################################################
'###################                              0.main) проверка версии ОС
'############################################################################
'############################################################################
windowsXP= false
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
	If objOperatingSystem.Version = "5.1.2600" Then 
	windowsXP = true
    end if
Next
If windowsXP Then
	appendLogMsg "Windows XP, скрипт в работу ......"
'############################################################################
'############################################################################
'###################                              0.main) проверка версии ОС
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'###################                              Т Е Л О ___ С К Р И П Т А
'############################################################################
'############################################################################

appendLogMsg "1) переменные"
'############################################################################
'############################################################################
'###################                              1) переменные
'############################################################################
'############################################################################

'При смене статики ДНС или ВИНС проставить true, и параметры серваков
keyChangeDNS = true
keyChangeWINS = false

'Укажи статический адрес ВИНС сервера
strPrimaryWinsServer = ""
strSecondaryWinsServer = ""

'Укажи статический адрес ДНС сервера
arrDNSServers = Array("xx.xx.xx.xx", "xx.xx.xx.xx")

'Если учетка в домене заблокирована, и нужно её включить, используйте эту опцию
useEnableAccount = true

' Место в котором ищем
' ДНС имя домена
domainDNSName = "child.domain.net" 'новый домен

' Сервак на котором ищем и порт
SearchADC = domainDNSName&":389" 'новый домен

' Временная OU для новых пользователей
TranzitOU = "Transit" 'в файле README1ST.TXT  пункт ***3***

'Новый пароль для локального админа, перебивает пароль локального администратора, список админов указываешь в массиве admNames
strNewPassword = "Придумай пароль"
'Возможные имена локальных админов
admNames = Array ("администратор", "administrator")

nextScriptFilePath = "\\"&domainDNSName&"\netlogon\Pereezd\profile.vbs" ' файл в соответствующее место не забудь положить, в файле README1ST.TXT  пункт ***8***

'Новый домен
strDomain = domainDNSName 'новый домен
'Пароль пользователя указанного в strUser
strPassword = "P@ssw0rd" 'в файле README1ST.TXT  пункт ***4***
'Пользователь с нов. домене
strUser = "pereezd"	'в файле README1ST.TXT  пункт ***4***

nextReboootAutoLoginUserName = strUser
nextReboootAutoLoginUserNamePass = strPassword

enableDefLogonEngLangRegFilePath = "\\"&domainDNSName&"\NETLOGON\Pereezd\preload.reg" 'в файле README1ST.TXT  пункт ***8***


testDrivePath = "\\xx.xx.xx.xx\gsdg4r7.dhsadre_sdkt53288" 'в файле README1ST.TXT  пункт ***2***
staryiDomenNetbiosName = "OLDOMAINNAME" 'старое нетбиос имя домена

'############################################################################
'############################################################################
'###################                              \1) переменные
'############################################################################
'############################################################################
appendLogMsg "\1) переменные"
'----------- получение текущего имени пользователя
Set objNetwork = CreateObject("Wscript.Network")
AldUser = objNetwork.UserName
Set objNetwork = nothing
appendLogMsg "получение текущего имени пользователя: " & AldUser
'----------- \получение текущего имени пользователя
appendLogMsg "2) функции"
'############################################################################
'############################################################################
'###################                              2) функции
'############################################################################
'############################################################################
'------------- Выводит инфу об ошипке
Function errMsg (srtActionName, strActionResult)
    If strActionResult = 0 Then
      errMsg = srtActionName & " OK"
    Else
      errMsg = srtActionName & " error" & vbCrLF & "err.Number=" & err.Number & vbCrLF & "err.description=" & err.description
	  Err.Clear
    End if
End Function
'------------- \Выводит инфу об ошипке
appendLogMsg "Функция errMsg загружена"

Dim TransTable  ' Массив соответствия
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


Function isJotVowel (letter)   ' замена русских на 2 лат
    isJotVowel = (letter = "я" _
	OR letter = "е" _
	OR letter = "ё" _
	OR letter = "ю")
end function

function Translit (RussianString)  ' преобразование рус в лат
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


function FioToSama (fio)  ' Получение zSAMa
  parts = Split (fio, " ")
  FioToSama = lcase (parts (0)) + left (parts (1), 1) + left (parts (2), 1)
end function

function FioToComp (fio)  ' Получение имени компьютера
  parts = Split (fio, " ")
  FioToComp = lcase (parts (0)) + "-" + left (parts (1), 1) + left (parts (2), 1)
end Function
appendLogMsg "Функции FioToSama и FioToComp загружены"

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
appendLogMsg "Функция convertDns2LdapName загружена, SearchCN = "& SearchCN 

Sub userAdd (fullName, sama)
        Set objOU = GetObject("LDAP://" & SearchADC & "/OU=" & TranzitOU & "," & SearchCN)
        Set objUser = objOU.Create("User", "cn= " + fullName)
        objUser.Put "displayName", fullName
        objUser.Put "givenName", fullName
        objUser.Put "userPrincipalName", sama + "@" + domainDNSName
        objUser.Put "sAMAccountName", sama
        objUser.SetInfo
        objUser.AccountDisabled = False
		objUser.IsAccountLocked = false
        objUser.SetInfo
End Sub
appendLogMsg "Процедура userAdd загружена"

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
		appendLogMsg "User LDAP://" & SearchADC & "/" & distUserName & " enabled"
End Sub
appendLogMsg "Процедура enableAccount загружена"
Function testDrive(dom, user, pass, driveMap, mapPath)
	Dim WshNetwork
	Set WshNetwork = WScript.CreateObject("WScript.Network")
	WshNetwork.MapNetworkDrive driveMap, mapPath,, dom&"\"&user, pass
'wscript.echo "test1"
	appendLogMsg "err.number = "&err.number
'	If Not (err.number <> 0) Then
	If (Not (err.number <> 0)) Or err.number = 424 then
		testDrive = 1
		Err.Clear
		WshNetwork.RemoveNetworkDrive driveMap, "true", "true" 'отключить ранее подключенный сетевой диск
	Else
		'wscript.echo "err.number1 = "&err.number
		If Err.number = -2147024811 Or Err.number = -2147023570 Then
				Err.Clear
				WshNetwork.RemoveNetworkDrive driveMap, "true", "true" 'отключить ранее подключенный сетевой диск
				WshNetwork.MapNetworkDrive driveMap, mapPath,, dom&"\"&user, pass
				'wscript.echo "err.number2 = "&err.number
			If Not (err.number <> 0) then
					testDrive = 1
					Err.Clear
					WshNetwork.RemoveNetworkDrive driveMap, "true", "true" 'отключить ранее подключенный сетевой диск
			End if
		End if
	End If
	Set WshNetwork = Nothing
		'wscript.echo "testDrive = "&testDrive
End Function
appendLogMsg "Функция testDrive загружена"
'--------------------------------- Переименование и включение в домен компутера
sub joinInDomainAndRename(tmpDomainName, tmpUserName, tmpUserPass, newDomainName, strNewCompName)
Set objNetwork = CreateObject("Wscript.Network")
strComputer2 = objNetwork.ComputerName
Set objNetwork = Nothing

Set objWMIService = GetObject("winmgmts:" _
                                & "{impersonationLevel=impersonate}!\\" & strComputer2 & "\root\cimv2")

		Const JOIN_DOMAIN = 1
        Const ACCT_CREATE = 2
        Const ACCT_DELETE = 4
        Const WIN9X_UPGRADE = 16
        Const DOMAIN_JOIN_IF_JOINED = 32
        Const JOIN_UNSECURE = 64
        Const MACHINE_PASSWORD_PASSED = 128
        Const DEFERRED_SPN_SET = 256
        Const INSTALL_INVOCATION = 262144
		WScript.Sleep 500
	Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer2 & "\root\cimv2:Win32_ComputerSystem.Name='" & strComputer2 & "'")
		ReturnValue22 = objComputer.JoinDomainOrWorkGroup(newDomainName, tmpUserPass, tmpDomainName &"\"& tmpUserName, NULL, JOIN_DOMAIN + ACCT_CREATE + DOMAIN_JOIN_IF_JOINED + DEFERRED_SPN_SET)
		If Not ReturnValue22 Then
				appendLogMsg "Команда "& errMsg ("Domain join", ReturnValue22)
				Else
				appendLogMsg "Команда "& errMsg ("Domain join завершилась с ошибкой", ReturnValue22)
		End if
		WScript.Sleep 500
	Set colComputersA = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
		For Each objComputer in colComputersA
		ReturnValue0 = objComputer.Rename(strNewCompName, tmpUserPass, tmpDomainName &"\"& tmpUserName)
		If Not ReturnValue0 Then
				appendLogMsg "Команда "& errMsg ("Computer rename", ReturnValue0)
				Else
				appendLogMsg "Команда "& errMsg ("Computer rename завершилась с ошибкой", ReturnValue0)
		End if
		Next
Set objNetwork = Nothing
Set objWMIService = nothing
end sub

'--------------------------------- /Переименование и включение в домен компутера
appendLogMsg "Функция joinInDomainAndRename загружена"
'############################################################################
'############################################################################
'###################                              \2) функции
'############################################################################
'############################################################################
appendLogMsg "\2) функции"
appendLogMsg "3) получение данных от пользователя"
'############################################################################
'############################################################################
'###################                              3) получение данных от пользователя
'############################################################################
'############################################################################
'----------------------------------------------------------------------------------
'      Уведомление о начале миграции
'----------------------------------------------------------------------------------
'Определяем путь к web-странице с формой
Path = WScript.ScriptFullName
Window1 = Left(Path, InStrRev(Path, "\")) & "attention.htm"
Window2 = Left(Path, InStrRev(Path, "\")) & "forms.htm"
Window3 = Left(Path, InStrRev(Path, "\")) & "pass.htm"

'Загружаем Internet Explorer
Set IEObject = WScript.CreateObject("InternetExplorer.Application")
IEObject.Navigate(Window1)
IEObject.MenuBar = 0
IEObject.ToolBar = 0
IEObject.StatusBar = 0
IEObject.FullScreen=1
IEObject.Top = 0
IEObject.Visible = 1
'Открываем страничку с формой и ждем, пока она загрузится
Do While (IEObject.Document.Body.All.OKClicked.Value="")
 WScript.Sleep(200)
Loop
WScript.Sleep(250)
'----------------------------------------------------------------------------------
'      Сбор сведений о пользователе
'----------------------------------------------------------------------------------


IEObject.Navigate(Window2)
IEObject.Visible = 1

Do While (IEObject.Document.Body.All.OKClicked1.Value="")
 WScript.Sleep(200)
Loop

      fName = Trim(IEObject.Document.UsForm.FirstName.Value) 'Имя: <input type=text name="FirstName">
      lName = Trim(IEObject.Document.UsForm.LastName.Value) 'Фамилия: <input type=text name="LastName">
      tName = Trim(IEObject.Document.UsForm.TrName.Value) 'Отчество: <input type=text name="TrName">
      password = IEObject.Document.UsForm.password.Value
zz03 = testDrive(staryiDomenNetbiosName, AldUser, password, "Y:", testDrivePath)
count = 1
'wscript.echo "count = "&count
	  While zz03 <> 1
		'wscript.echo "testDrive(staryiDomenNetbiosName, AldUser, password, Y:, testDrivePath) = "&testDrive(staryiDomenNetbiosName, AldUser, password, "Y:", testDrivePath)
		IEObject.Navigate(Window3)
		IEObject.Visible = 1

		Do While (IEObject.Document.Body.All.OKClicked2.Value="")
		 WScript.Sleep(200)
		Loop

	      password = IEObject.Document.UsForm0.password0.Value
		  zz03 = 0
		  zz03 = testDrive(staryiDomenNetbiosName, AldUser, password, "Y:", testDrivePath)
		  'wscript.echo "password = "&password
		  count = count + 1
		  'wscript.echo "count = "&count

		  'IEObject.Visible = 0
	  
	  Wend

IEObject.Visible = 0
IEObject.Quit
WScript.Sleep(250)
appendLogMsg "Данные от пользователя получены: " & lName &" "& fName &" "& tName &", пароль скрыт ******"

FIO = lName & " " & fName & " " & tName
appendLogMsg "ФИО : "& FIO
CompName = FioToComp(Translit(FIO))  ' новое имя компьютера
appendLogMsg "Новое имя компа будет: "&CompName

'wscript.quit(1)

'############################################################################
'############################################################################
'###################                              \3) получение данных от пользователя
'############################################################################
'############################################################################
appendLogMsg "\3) получение данных от пользователя"


Set objIE = CreateObject("InternetExplorer.Application")
objIE.Navigate Left(Path, InStrRev(Path, "\")) & "splash1.htm"
objIE.MenuBar = 0
objIE.ToolBar = 0
objIE.StatusBar = 0
objIE.FullScreen=1
objIE.Top = 0
While objIE.Busy
    Wscript.Sleep 200
Wend
objIE.Visible = 1	



appendLogMsg "4) смена сетевых настроек"
'############################################################################
'############################################################################
'###################                              4) смена сетевых настроек
'############################################################################
'############################################################################

Set objNetwork = CreateObject("Wscript.Network")
  strComputer = objNetwork.ComputerName
Set objNetwork = nothing
'--------------------------------- DNS&&WINS change
If ( keyChangeWINS Or keyChangeDNS ) then
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

  Set colNetCards = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
  Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
  For Each objNetCard in colNetCards
    If keyChangeWINS Then
      ' Строки изменения ВИНС сервера
		errResult0 = objNetCard.SetWINSServer(strPrimaryWinsServer, strSecondaryWinsServer)
		If Not errResult0 Then
				appendLogMsg "Команда "& errMsg ("WINS Change", errResult0)
				Else
				appendLogMsg "Команда "& errMsg ("WINS Change завершилась с ошибкой", errResult0)
		End if
	End If
		errResult1 = objNetworkSettings.EnableWINS(false, false)
		If Not errResult1 Then
				appendLogMsg "Команда "& errMsg ("WINS Lmhost Change", errResult1)
				Else
				appendLogMsg "Команда "& errMsg ("WINS Lmhost Change завершилась с ошибкой", errResult1)
		End if
	If keyChangeDNS Then
      'Строки изменения ДНС сервера
		errResult3 = objNetCard.SetDNSServerSearchOrder(arrDNSServers)
		'wscript.echo errResult3
		If Not errResult3 Then
				appendLogMsg "Команда "& errMsg ("DNS Change", errResult3)
				Else
				appendLogMsg "Команда "& errMsg ("DNS Change завершилась с ошибкой", errResult3)
		End if
	End If
  Next
End If
'--------------------------------- /DNS&&WINS change
'############################################################################
'############################################################################
'###################                              \4) смена сетевых настроек
'############################################################################
'############################################################################
appendLogMsg "\4) смена сетевых настроек"
appendLogMsg "5) поиск и разблокировка, либо создание пользователя"
'############################################################################
'############################################################################
'###################                              5) поиск и разблокировка, либо создание пользователя
'############################################################################
'############################################################################
'--------------------------------- Запрос пользователя в АД
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"

Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

objCommand.CommandText = _
          "<LDAP://" & SearchADC & "/" & SearchCN & ">;" & _
          "(&(objectCategory=person)(objectClass=user)" & _
          "(displayName=" & FIO & "));" & _
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
 zSAMa = FioToSama (Translit(FIO)) ' inputBox("Введи samAccountName:")
 userAdd FIO, zSAMa
 appendLogMsg "Пользователь " & FIO & " | " & zSAMa & " создан успешно."
End If

If useEnableAccount then
 enableAccount zSAMa 'логирование внутри процедуры
End if
'--------------------------------- /Запрос пользователя в АД
'############################################################################
'############################################################################
'###################                              \5) поиск и разблокировка, либо создание пользователя
'############################################################################
'############################################################################
appendLogMsg "\5) поиск и разблокировка, либо создание пользователя"
appendLogMsg "6) смена пароля локального админа"
'############################################################################
'############################################################################
'###################                              6) смена пароля локального админа
'############################################################################
'############################################################################
Set objNetwork = CreateObject("Wscript.Network")
strComputer1 = objNetwork.ComputerName
Set objNetwork = Nothing

Set objDomain = GetObject("WinNT://" & strComputer1)
objDomain.Filter = Array("User")
	found = false
For Each admName in admNames
	for each objUser In objDomain
		If LCase(objUser.Name) = admName Then
			objUser.SetPassword (strNewPassword)
			found = true
			if err.number <> 0 then
				appendLogMsg "При установке пароля локального администратора ошибка: " & err.number & " " & err.description
			else
				appendLogMsg "пароль " & objUser.Name & " установлен"
			end if
		End If
	next		
Next
If Not found Then
		appendLogMsg "заданное имя локального администратора не найдено"
End if


'############################################################################
'############################################################################
'###################                              \6) смена пароля локального админа
'############################################################################
'############################################################################
appendLogMsg "\6) смена пароля локального админа"
appendLogMsg "8) включение компьютера в домен и переименование"
'############################################################################
'############################################################################
'###################                              8) включение компьютера в домен и переименование
'############################################################################
'############################################################################
joinInDomainAndRename staryiDomenNetbiosName, AldUser, password, strDomain, CompName
WScript.Sleep 500
'############################################################################
'############################################################################
'###################                              \8) включение компьютера в домен и переименование
'############################################################################
'############################################################################
appendLogMsg "\8) включение компьютера в домен и переименование"
appendLogMsg "7) подготовка действий после перезагрузки"
'############################################################################
'############################################################################
'###################                              7) подготовка действий после перезагрузки
'############################################################################
'############################################################################
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile nextScriptFilePath, "c:\"
set objFSO = Nothing

const HKEY_LOCAL_MACHINE = &H80000002
set WshShell = CreateObject("WScript.Shell")
strComputer = "."
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
strValueName = "DefaultUserName"
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,nextReboootAutoLoginUserName
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"DefaultDomainName",Split(strDomain, ".", -1, 1)(0)
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"AltDefaultDomainName",Split(strDomain, ".", -1, 1)(0)
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"CachePrimaryDomain",Split(strDomain, ".", -1, 1)(0)
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"DefaultPassword",nextReboootAutoLoginUserNamePass
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"AutoAdminLogon","1"
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"ForceAutoLogon","1"
oReg.SetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce","Profile","C:\profile.vbs " & zSAMa & " " & AldUser
WshShell.Run "regedit -s "&enableDefLogonEngLangRegFilePath
'############################################################################
'############################################################################
'###################                              \7) подготовка действий после перезагрузки
'############################################################################
'############################################################################
appendLogMsg "\7) подготовка действий после перезагрузки"
appendLogMsg "9) перезагрузка компьютера"
'############################################################################
'############################################################################
'###################                              9) перезагрузка компьютера
'############################################################################
'############################################################################
WScript.Sleep 500


objIE.Visible = 0
objIE.Quit
WScript.Sleep(250)


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Shutdown)}!\\" & _
        strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    ObjOperatingSystem.Reboot()
Next
'############################################################################
'############################################################################
'###################                              \9) перезагрузка компьютера
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'###################                              Т Е Л О ___ С К Р И П Т А
'############################################################################
'############################################################################
'Set objFSO001 = CreateObject ("Scripting.FileSystemObject")
'Set objFile = objFSO001.OpenTextFile ("c:\compName.txt", 8, True)
'objFile.WriteLine (CompName)
Else
	appendLogMsg "ОС не Windows XP."
End if
'############################################################################
'############################################################################
'###################                              \0.main) проверка версии ОС
'############################################################################
'############################################################################






