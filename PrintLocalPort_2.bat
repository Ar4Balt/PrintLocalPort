@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
color 71
title Управление TCP/IP (Standard) портами принтеров

:: путь к встроенному скрипту Windows
set SCRIPT=%SystemRoot%\System32\Printing_Admin_Scripts\ru-RU\prnport.vbs
if not exist "%SCRIPT%" (
    echo Не найден %SCRIPT%
    pause
    exit /b
)

:: ================================
:: Список портов (имена)
:: ================================
set port1=KIP7170
set port2=OCE-360 PW360 WPD2
set port3=OCE-360.AZOT.COM.BY_PW360_WPD2
set port4=PRN-OKI-2BCB2D
set port5=PRN-EPS-AED7D2
set port6=PRN-PAN-49509E
set port7=PRN-PAN-4954AE
set port8=PRN-PAN-32DC63
set port9=PRN-PAN-DF6D5A
set port10=PRN-PAN-54BCA6
set port11=PRN-PAN-32DBFA
set port12=PRN-PAN-54BCAD
set port13=PRN-PAN-374302
set port14=PRN-XER-9222E4
set port15=TCS300_1_TCS300_WPD2
set port16=TCS300_2_TCS300 WPD2
set port17=HPOA403D
set port18=HPOA91AD
set port19=HPOA411D
set port20=HP0A411D
set port21=XEROX-6705

:MENU
cls
echo ===========================================
echo      Управление TCP/IP (Standard) портами
echo ===========================================
echo.
echo ==== Существующие TCP/IP порты: ====
cscript //nologo "%SCRIPT%" -l
echo ===========================================
echo.
echo 1) Добавить порт
echo 2) Удалить порт
echo 3) Выход
echo.
set /p action=Выберите действие (1-3): 

if "%action%"=="1" goto ADD
if "%action%"=="2" goto DELETE
if "%action%"=="3" exit
goto MENU

:ADD
cls
echo ===========================================
echo           Добавление TCP/IP портов
echo ===========================================
echo.
echo Список доступных имен портов:
for /L %%i in (1,1,21) do (
    call echo   %%i^) !port%%i!
)
echo   22) Ввести своё имя вручную
echo.
set /p choice=Введите номера через запятую (например: 1,3,7): 
set choice=%choice: =%
set choice=%choice:,= %

for %%i in (%choice%) do (
    if "%%i"=="22" (
        set /p custom=Введите имя хоста/IP: 
        call :AddPort "!custom!"
    ) else (
        call set "p=%%port%%i%%"
        call :AddPort "!p!"
    )
)

pause
goto MENU

:DELETE
cls
echo ===========================================
echo          Удаление TCP/IP портов
echo ===========================================
echo.
echo Существующие порты:
cscript //nologo "%SCRIPT%" -l
echo.
set /p delport=Введите имя порта для удаления: 
call :DelPort "!delport!"

pause
goto MENU

:AddPort
set "PORTNAME=%~1"
:: Проверка существования
cscript //nologo "%SCRIPT%" -l | findstr /I "%PORTNAME%" >nul
if %errorlevel%==0 (
    echo Порт уже существует: %PORTNAME%
    exit /b
)

echo Добавляем TCP/IP порт: %PORTNAME% ...
cscript //nologo "%SCRIPT%" -a -r "%PORTNAME%" -h "%PORTNAME%" -o raw -n 9100

if %errorlevel%==0 (
    echo Порт успешно создан: %PORTNAME%
) else (
    echo Ошибка при создании порта %PORTNAME%
)
exit /b

:DelPort
set "PORTNAME=%~1"
:: Проверка существования
cscript //nologo "%SCRIPT%" -l | findstr /I "%PORTNAME%" >nul
if %errorlevel%==1 (
    echo Порт не найден: %PORTNAME%
    exit /b
)

echo Удаляем TCP/IP порт: %PORTNAME% ...
cscript //nologo "%SCRIPT%" -d -r "%PORTNAME%"

if %errorlevel%==0 (
    echo Порт успешно удалён: %PORTNAME%
) else (
    echo Ошибка при удалении порта %PORTNAME%
)
exit /b
