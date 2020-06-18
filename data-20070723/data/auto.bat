@echo off
@rem этот скрипт использует rename-join.vbs (т.е. НетБиосИмяДомена)
@rem EDIT share.ip, admin@domen
set REALNAME=%USERNAME%
set REALDOMAIN=%USERDOMAIN%
@regedit -s \\share.ip\data\preload.reg
@cscript \\share.ip\data\locadm.vbs
@runas /env /user:admin@domen.name.ru /netonly "cmd /k \"cscript \\share.ip\data\user-new.vbs\""
@runas /user:admin@domen.name.ru /netonly "cmd /k \"cscript \\share.ip\data\rename-join.vbs\""
@pause
@\\share.ip\data\shutdown -r -t 00 -f
