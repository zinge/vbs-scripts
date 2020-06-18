On Error Resume Next
'###########################
'###########################
'Новый домен
strDomain = "NetBiosDomainName"
'Пароль пользователя указанного в strUser
strPassword = "пароль"
'Пользователь с нов. домене
strUser = "пользователь"




















































' ------------- считываем SAMA из файла

Set objFSO = CreateObject ("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile ("c:\sama.txt", 1)
zSAMA = objFile.ReadLine
Set objFile = Nothing
Set objFSO = CreateObject ("Scripting.FileSystemObject")
objFSO.DeleteFile ("c:\sama.txt")

if zSAMA = "" then
  zSAMA = "Familija-IO"
end if
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
'-------------
'--------------------------------- Переименование и включение в домен компутера
Function SamaToHost (sama)
  SamaToHost = UCase (Left (sama, 1)) + Mid (sama, 2, Len (sama) - 3) + "-" + UCase (Right (sama, 2))
End Function


Set objNetwork = CreateObject("Wscript.Network")
strComputer = objNetwork.ComputerName

Set objWMIService = GetObject("winmgmts:" _
                                & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
	If objOperatingSystem.Version = "5.1.2600" Then 
	strNewCompName = InputBox("Введите новое имя компа:", "Comp name", SamaToHost (zSAMa))
		Const JOIN_DOMAIN = 1
        Const ACCT_CREATE = 2
        Const ACCT_DELETE = 4
        Const WIN9X_UPGRADE = 16
        Const DOMAIN_JOIN_IF_JOINED = 32
        Const JOIN_UNSECURE = 64
        Const MACHINE_PASSWORD_PASSED = 128
        Const DEFERRED_SPN_SET = 256
        Const INSTALL_INVOCATION = 262144
	Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & strComputer & "'")
		ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomain, strPassword, strDomain & "\" & strUser, NULL, JOIN_DOMAIN + ACCT_CREATE + DOMAIN_JOIN_IF_JOINED + DEFERRED_SPN_SET)
		errMsg "Domain join", ReturnValue
		WScript.Sleep 2500
	Set colComputersA = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
		For Each objComputer in colComputersA
		ReturnValue = objComputer.Rename(strNewCompName, strPassword, strUser)
		errMsg "Computer rename", ReturnValue
		Next
	Else
		wscript.echo "Please rename the computer manually"
	End If 
Next
'--------------------------------- /Переименование и включение в домен компутера
