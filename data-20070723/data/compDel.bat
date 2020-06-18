@echo off
@rem EDIT share.ip, admin@domen
@runas /user:admin@domen.name.ru /netonly "cmd /k \"cscript \\share.ip\data\compDel.vbs\""