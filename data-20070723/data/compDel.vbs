'замените domen0000 на свое
SearchADC = "domen0000.name.ru:389"
SearchCN = "dc=domen0000,dc=name,dc=ru"











strComputer = InputBox("Новое глючное имя компа? ")


set objComputer = GetObject("LDAP://" & SearchADC & "/CN=" & strComputer & _
    ",CN=Computers," & SearchCN)
objComputer.DeleteObject (0)