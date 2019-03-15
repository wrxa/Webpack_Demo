::Class Name : test.bat
::author  wang.rui
::version 1.00 2019/02/13
::History
::1.00 2019/02/13  FXS)wang.rui			initialize release.

@echo off
echo start BAT
:start
:: start "" "C:\Program Files (x86)\sakura\sakura.exe"
::CScript  test.vbs
CScript  NoticeMailSend.vbs

:: 下行的10代表每100秒钟后循环执行start中的内容
ping -n 100 127.1>NUL
goto start