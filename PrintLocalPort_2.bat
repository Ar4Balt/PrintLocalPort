@echo off
color 71
echo    1    11111 11111 11111
echo   1 1       1 1   1   1  
echo  1   1  11111 1   1   1  
echo  11111      1 1   1   1  
echo 1     1 11111 11111   1  
echo ---
title Добавление сетевых портов принтера

:: Автоматически получение имени компьютера
set mypc=%COMPUTERNAME%

:: ================================
:: Список доступных сетевых портов
:: ================================
set port1=\\%mypc%\KIP7170
set port2=\\%mypc%\OCE-360 PW360 WPD2
set port3=\\%mypc%\OCE-360.AZOT.COM.BY_PW360_WPD2
set port4=\\%mypc%\PRN-OKI-2BCB2D
set port5=\\%mypc%\PRN-EPS-AED7D2
set port6=\\%mypc%\PRN-PAN-49509E
set port7=\\%mypc%\PRN-PAN-4954AE
set port8=\\%mypc%\PRN-PAN-32DC63
set port9=\\%mypc%\PRN-PAN-DF6D5A
set port10=\\%mypc%\PRN-PAN-54BCA6
set port11=\\%mypc%\PRN-PAN-32DBFA
set port12=\\%mypc%\PRN-PAN-54BCAD
set port13=\\%mypc%\PRN-PAN-374302
set port14=\\%mypc%\PRN-XER-9222E4
set port15=\\%mypc%\TCS300_1_TCS300_WPD2
set port16=\\%mypc%\TCS300_2_TCS300 WPD2
set port17=\\%mypc%\HPOA403D
set port18=\\%mypc%\HPOA91AD
set port19=\\%mypc%\HPOA411D
set port20=\\%mypc%\HP0A411D
set port21=\\%mypc%\XEROX-6705

:MENU
cls
echo ===========================================
echo     Управление сетевыми портами принтера
echo ===========================================
echo.
echo ==== Уже существующие порты в системе ====
REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports"
echo ===========================================
echo.
echo Что вы хотите сделать?
echo   1) Добавить порты
echo   2) Удалить порты
echo   3) Выйти
echo.
set /p action=Выберите действие (1-3): 

if "%action%"=="1" goto ADD
if "%action%"=="2" goto DELETE
if "%action%"=="3" exit
goto MENU

:ADD
cls
echo ===========================================
echo   Добавление сетевых портов
echo ===========================================
echo.
echo Список доступных портов:
for /L %%i in (1,1,21) do (
    call echo   %%i^) !port%%i!
)
echo  22) Ввести свой порт вручную
echo.
set /p choice=Введите номера через запятую (например: 1,3,7): 
set choice=%choice: =%

for %%i in (%choice%) do (
    if "%%i"=="22" (
        set /p custom=Введите свой сетевой порт (\\сервер\принтер): 
        call :AddPort "%custom%"
    ) else (
        call set "p=%%port%%i%%"
        call :AddPort "%%p%%"
    )
)

echo.
echo Перезапускаем службу печати...
net stop spooler >nul
net start spooler >nul

pause
goto MENU

:DELETE
cls
echo ===========================================
echo   Удаление сетевых портов
echo ===========================================
echo.
echo Список доступных портов:
for /L %%i in (1,1,21) do (
    call echo   %%i^) !port%%i!
)
echo  22) Указать вручную
echo.
set /p delchoice=Введите номера через запятую (например: 2,5,10): 
set delchoice=%delchoice: =%

for %%i in (%delchoice%) do (
    if "%%i"=="22" (
        set /p customDel=Введите полный путь порта (\\сервер\принтер): 
        call :DelPort "%customDel%"
    ) else (
        call set "p=%%port%%i%%"
        call :DelPort "%%p%%"
    )
)

echo.
echo Перезапускаем службу печати...
net stop spooler >nul
net start spooler >nul

pause
goto MENU

:AddPort
REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports" /v %~1 >nul 2>&1
if %errorlevel%==0 (
    echo Порт уже существует: %~1
) else (
    echo Добавляем порт: %~1
    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports" /v %~1 /t REG_SZ /d "" /f >nul
)
exit /b

:DelPort
REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports" /v %~1 >nul 2>&1
if %errorlevel%==0 (
    echo Удаляем порт: %~1
    REG DELETE "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports" /v "%~1" /f >nul
) else (
    echo Порт не найден: %~1
)
exit /b
