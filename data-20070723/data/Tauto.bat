@echo off
@rem этот скрипт использует Trename-join.vbs (т.е. ДнсИмяДомена)
@rem EDIT share.ip, admin@domen
set REALNAME=%USERNAME%
set REALDOMAIN=%USERDOMAIN%
@regedit -s \\share.ip\data\preload.reg
@cscript \\share.ip\data\locadm.vbs
@runas /env /user:admin@domen.name.ru /netonly "cmd /k \"cscript \\share.ip\data\user-new.vbs\""
@runas /user:admin@domen.name.ru /netonly "cmd /k \"cscript \\share.ip\data\Trename-join.vbs\""
@pause
@\\share.ip\data\shutdown -r -t 00 -f
