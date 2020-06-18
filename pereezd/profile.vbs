' ########################
' ############## ПЕРЕМЕННЫЕ
' ########################
netlogonMapName = "Z:"
domainDNSName = "child.domain.net" 'новый домен
netlogonPath = "\\"&domainDNSName&"\Netlogon" 'папка в старом домене где лежат необходимые нам скрипты
' указываем папку где в старом домене где лежат скрипты
netBiosDomainNameOld = "OLDOMAINNAME" 'старое НетБИОС имя домена, участвует при перемещении профиля
netBiosDomainNameNew = Split(domainDNSName, ".", -1, 1)(0) 'новое НетБИОС имя домена, участвует при перемещении профиля, и при задании имени пользователя по умолчанию
domainGroup = "Пользователи домена" 'соответствующую группу добавляем в локальные группы указанные в массиве admNames
admNames = Array ("администраторы", "administrators")

' ########################
'\ ############## ПЕРЕМЕННЫЕ
' ########################
'############################################################################
'############################################################################
'###################                              0) врубаем логирование
'############################################################################
'############################################################################
logFileName = "domainMigrate.log"
logFileNameDriveName = "c:"
Sub appendLogMsg (domainMigrateLogMsgTmp) 
   Const ForAppending = 8
   Dim fso, f
   Set fso02 = CreateObject("Scripting.FileSystemObject")
   Set f = fso02.OpenTextFile(logFileNameDriveName&"\"&logFileName, ForAppending, True)
   f.Write Now & "   " & domainMigrateLogMsgTmp & VbCrLf
   f.Close
   Set fso02 = nothing
End Sub
appendLogMsg "Вторая часть перевода, начат: "& Now & VbCrLf
'############################################################################
'############################################################################
'###################                              \0) врубаем логирование
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
zSAMa = WScript.Arguments(0) 'в 1 аргументе командрой строке скрипта должно быть указано новое имя пользователя
appendLogMsg "Новое имя пользователя: " & zSAMa
AldName = WScript.Arguments(1) 'во 2 аргументе командрой строке скрипта должно быть указано старое имя пользователя
appendLogMsg "Старое имя пользователя: " & AldName
Set WshNetwork = WScript.CreateObject("WScript.Network")
WScript.Sleep 5000
reTurnMap = WshNetwork.MapNetworkDrive (netlogonMapName, netlogonPath) 'подключаем папку netlogonPath как локальный диск с именем netlogonMapName
		If Not reTurnMap Then
				appendLogMsg "Команда "& errMsg ("Подмонтирования диска", reTurnMap)
				Else
				appendLogMsg "Команда "& errMsg ("Подмонтирования диска завершилась с ошибкой", reTurnMap)
		End If

Set objIE = CreateObject("InternetExplorer.Application")
objIE.Navigate netlogonMapName & "\Pereezd\splash.htm"
objIE.MenuBar = 0
objIE.ToolBar = 0
objIE.StatusBar = 0
objIE.FullScreen=1
objIE.Top = 0
While objIE.Busy
    Wscript.Sleep 200
Wend
objIE.Visible = 1	
WScript.Sleep 5000
WScript.Sleep 5000
set WshShell = CreateObject("WScript.Shell")
retuRnKav0 = WshShell.Run ("%windir%\system32\cmd.exe /c net stop AVP")
retuRn0 = WshShell.Run  (netlogonMapName & "\Pereezd\moveuser.exe " & netBiosDomainNameOld & "\" & AldName & " "&netBiosDomainNameNew&"\" & zSAMa & " /y /k", 2, True) 'с netlogonMapName запускаем скрипт по перемещению профиля пользователя
		If Not retuRn0 Then
				appendLogMsg "Команда "& errMsg ("Переноса профиля", retuRn0)
				Else
				appendLogMsg "Команда "& errMsg ("Переноса профиля завершилась с ошибкой", retuRn0)
		End If
retuRnKav1 = WshShell.Run ("%windir%\system32\cmd.exe /c net start AVP")		
' изменить имя пользователя при входе на нужное нам новое

Set objNetwork1 = CreateObject("Wscript.Network")
strComputer1 = objNetwork1.ComputerName

Set objDomain = GetObject("WinNT://" & strComputer1)
objDomain.Filter = Array("Group")

For Each admName in admNames
	for each objUser In objDomain
		If LCase(objUser.Name) = admName Then
			set WshShell1 = CreateObject("WScript.Shell")
			retuRnKav0 = WshShell1.Run ("%windir%\system32\cmd.exe /c net localgroup "&objUser.Name&" """&netBiosDomainNameNew&"\"&domainGroup&""" /add")
		End If
	next		
Next
Set objDomain = nothing
Set objNetwork1 = Nothing
set WshShell1 = nothing



set WSHShell= WScript.CreateObject("WScript.Shell")
const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
strValueName = "DefaultUserName"
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,zSAMa
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"DefaultDomainName",netBiosDomainNameNew
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"DefaultPassword","0"
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"AutoAdminLogon","0"
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"ForceAutoLogon","0"
appendLogMsg "\ изменить имя пользователя при входе на нужное нам новое"
'\ изменить имя пользователя при входе на нужное нам новое
WshNetwork.RemoveNetworkDrive netlogonMapName, "true", "true" 'отключить ранее подключенный сетевой диск
appendLogMsg "отключить ранее подключенный сетевой диск"
appendLogMsg "Перегружаю..."
'WshShell.Popup "Перегружаю..."
objIE.Visible = 0
objIE.Quit
WScript.Sleep 250
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Shutdown)}!\\" & _
        strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
Const EWX_LOGOFF = 0 
For Each objOperatingSystem in colOperatingSystems
    ObjOperatingSystem.win32shutdown EWX_LOGOFF
Next