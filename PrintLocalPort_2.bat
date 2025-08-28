@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
color 71
title Управление TCP/IP (Standard) портами принтеров

:: ======== Путь к встроенному скрипту Windows ========
set SCRIPT=%SystemRoot%\System32\Printing_Admin_Scripts\ru-RU\prnport.vbs
if not exist "%SCRIPT%" (
    echo Не найден %SCRIPT%
    pause
    exit /b
)

:: ======== Лог-файл ========
set LOGFILE=%~dp0PrinterPorts.log

:: ======== Предустановленные порты ========
set ports[1]=KIP7170
set ports[2]=OCE-360 PW360 WPD2
set ports[3]=OCE-360.AZOT.COM.BY_PW360_WPD2
set ports[4]=PRN-OKI-2BCB2D
set ports[5]=PRN-EPS-AED7D2
set ports[6]=PRN-PAN-49509E
set ports[7]=PRN-PAN-4954AE
set ports[8]=PRN-PAN-32DC63
set ports[9]=PRN-PAN-DF6D5A
set ports[10]=PRN-PAN-54BCA6
set ports[11]=PRN-PAN-32DBFA
set ports[12]=PRN-PAN-54BCAD
set ports[13]=PRN-PAN-374302
set ports[14]=PRN-XER-9222E4
set ports[15]=TCS300_1_TCS300_WPD2
set ports[16]=TCS300_2_TCS300 WPD2
set ports[17]=HPOA403D
set ports[18]=HPOA91AD
set ports[19]=HPOA411D
set ports[20]=HP0A411D
set ports[21]=XEROX-6705

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
echo 3) Экспорт списка портов
echo 4) Выход
echo.
set /p action=Выберите действие (1-4): 
echo %date% %time% [MENU] Выбран пункт меню: %action%>>"%LOGFILE%"

if "%action%"=="1" goto ADD
if "%action%"=="2" goto DELETE
if "%action%"=="3" goto EXPORT
if "%action%"=="4" exit
goto MENU

:ADD
cls
echo ===========================================
echo           Добавление TCP/IP портов
echo ===========================================
echo.
echo Список доступных портов:
for /L %%i in (1,1,21) do (
    call echo   %%i^) !ports[%%i]!
)
echo   22) Ввести своё имя вручную
echo.
set /p choice=Введите номера через запятую или диапазон (например: 1,3,7-9): 
echo %date% %time% [ADD] Введено: %choice%>>"%LOGFILE%"
set choice=%choice: =%
call :ParseChoice "%choice%"
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
echo %date% %time% [DELETE] Выбран порт для удаления: %delport%>>"%LOGFILE%"
set /p confirm=Вы уверены, что хотите удалить порт "%delport%"? (Y/N): 
echo %date% %time% [DELETE] Подтверждение: %confirm%>>"%LOGFILE%"
if /I "%confirm%"=="Y" (
    call :DelPort "%delport%"
) else (
    echo Отмена удаления.
    echo %date% %time% [DELETE] Удаление отменено>>"%LOGFILE%"
)
pause
goto MENU

:EXPORT
cls
echo ===========================================
echo       Экспорт списка TCP/IP портов
echo ===========================================
cscript //nologo "%SCRIPT%" -l > "%~dp0PrinterPortsList.txt"
echo Список портов сохранен в PrinterPortsList.txt
echo %date% %time% [EXPORT] Список портов сохранён>>"%LOGFILE%"
pause
goto MENU

:ParseChoice
set "input=%~1"
set "input=%input:,= %"  :: заменяем запятые на пробелы

for %%a in (%input%) do (
    set "element=%%a"
    set "element=!element: =!"
    :: Проверяем, есть ли дефис
    echo !element! | findstr "-" >nul
    if !errorlevel! == 0 (
        :: Разбираем диапазон
        for /F "tokens=1,2 delims=-" %%b in ("!element!") do (
            set /A start=%%b
            set /A end=%%c
            if "!end!"=="" set /A end=!start!
            for /L %%i in (!start!,1,!end!) do (
                call set "portname=%%ports[%%i]%%"
                if "!portname!"=="" (
                    echo [Ошибка] Порт %%i не существует, пропуск.
                    echo %date% %time% [ERROR] Порт %%i не существует, пропуск>>"%LOGFILE%"
                ) else (
                    echo %date% %time% [ADD] Добавление порта: !portname!>>"%LOGFILE%"
                    call :AddPort "!portname!"
                )
            )
        )
    ) else (
        :: Одиночный номер или пользовательский ввод
        if "!element!"=="22" (
            set /p custom=Введите имя хоста/IP: 
            if "!custom!"=="" (
                echo [Ошибка] Пустое имя порта, пропуск.
                echo %date% %time% [ERROR] Пустое имя порта, пропуск>>"%LOGFILE%"
            ) else (
                echo %date% %time% [ADD] Пользовательский порт: !custom!>>"%LOGFILE%"
                call :AddPort "!custom!"
            )
        ) else (
            call set "portname=%%ports[%%a]%%"
            if "!portname!"=="" (
                echo [Ошибка] Порт %%a не существует, пропуск.
                echo %date% %time% [ERROR] Порт %%a не существует, пропуск>>"%LOGFILE%"
            ) else (
                echo %date% %time% [ADD] Добавление выбранного порта: !portname!>>"%LOGFILE%"
                call :AddPort "!portname!"
            )
        )
    )
)
goto :eof

:AddPort
set "PORTNAME=%~1"
if "%PORTNAME%"=="" (
    echo [Ошибка] Пустое имя порта, пропуск.
    echo %date% %time% [ERROR] Пустое имя порта>>"%LOGFILE%"
    exit /b
)
cscript //nologo "%SCRIPT%" -l | findstr /I "%PORTNAME%" >nul
if %errorlevel%==0 (
    echo [Ошибка] Порт уже существует: %PORTNAME%
    echo %date% %time% [ERROR] Порт уже существует: %PORTNAME%>>"%LOGFILE%"
    exit /b
)
echo Добавляем TCP/IP порт: %PORTNAME% ...
cscript //nologo "%SCRIPT%" -a -r "%PORTNAME%" -h "%PORTNAME%" -o raw -n 9100
if %errorlevel%==0 (
    echo [OK] Порт успешно создан: %PORTNAME%
    echo %date% %time% [OK] Порт успешно создан: %PORTNAME%>>"%LOGFILE%"
) else (
    echo [Ошибка] Ошибка при создании порта %PORTNAME%
    echo %date% %time% [ERROR] Ошибка при создании порта %PORTNAME%>>"%LOGFILE%"
)
exit /b

:DelPort
set "PORTNAME=%~1"
if "%PORTNAME%"=="" (
    echo [Ошибка] Имя порта не указано!
    echo %date% %time% [ERROR] Имя порта не указано>>"%LOGFILE%"
    exit /b
)
echo Удаляем TCP/IP порт: %PORTNAME%
echo %date% %time% [DELETE] Удаляем порт: %PORTNAME%>>"%LOGFILE%"
cscript //nologo "%SCRIPT%" -l | findstr /I "%PORTNAME%" >nul
if %errorlevel%==1 (
    echo [Ошибка] Порт не найден: %PORTNAME%
    echo %date% %time% [ERROR] Порт не найден: %PORTNAME%>>"%LOGFILE%"
    exit /b
)
for /f "delims=" %%o in ('cscript //nologo "%SCRIPT%" -d -r "%PORTNAME%" 2^>^&1') do (
    echo %%o
    echo %date% %time% [DELETE_OUTPUT] %%o>>"%LOGFILE%"
)
if %errorlevel%==0 (
    echo [OK] Порт успешно удалён: %PORTNAME%
    echo %date% %time% [OK] Порт успешно удалён: %PORTNAME%>>"%LOGFILE%"
) else (
    echo [Ошибка] Не удалось удалить порт %PORTNAME% (код ошибки: %errorlevel%)
    echo %date% %time% [ERROR] Не удалось удалить порт %PORTNAME% (код ошибки: %errorlevel%)>>"%LOGFILE%"
)
exit /b
