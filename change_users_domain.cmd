::  script to update "Users Domain" on vue Pacs
::	PTD	08-31-20222
::	copy the 2 files to service drive
::		change_users_domain.cmd
::		change_users_domain.sql
::	run change_users_domain.cmd in Admin CMD
::
::
@ECHO off
SetLocal EnableDelayedExpansion

SET PACSNAME=
SET OLDDOM=
SET NEWDOM=
SET SCD=%CD%

:SetUD
ECHO.
ECHO Enter Pacs Name
SET /P PACNAM=" - Enter Pacs Name in lower case "
SET /P OLDDOM=" - Enter old users domain, 2 letters lower case "
SET /P NEWDOM=" - Enter new users domain, 2 letters lower case "

IF [[%PACNAM%]]==[[]] GoTo SetUD
IF [[%OLDDOM%]]==[[]] GoTo SetUD
IF [[%NEWDOM%]]==[[]] GoTo SetUD

:SetDIR
ECHO.
ECHO Enter the directory where the change users domain scripts are located (like C:\temp)
SET /P UDSD=" - Do NOT use spaces: "

IF [[%UDSD%]]==[[]] GoTo SetDIR


:: ask to continue
:ASK
ECHO.
ECHO Ready to users domain:
ECHO %NEWDOM%
ECHO - answering N or n will exit
SET /P QA="Are you sure you want to continue[Y/N]? "
  IF /I "%QA%" EQU "Y" GOTO CHANGEUD
  IF /I "%QA%" EQU "N" GOTO BYBY
ECHO.
ECHO You have to answer the question . . .
GOTO ASK

:: exit at user request
:BYBY
ECHO.
ECHO You answered No, exiting!
ENDLOCAL
GOTO :EOF

:CHANGEUD
ECHO Current Users Domain is
CALL %imaginet_root%utils\tool_cfg -c get_value -p imaginet\system\nodes\%PACNAM%\properties\current_domain
ECHO Users Domain to %NEWDOM%
CD  %UDSD%
CALL %imaginet_root%scripts\update_user_domain.pl -all -new_domain %NEWDOM%
CALL sqlplus system/a1d2m7i4@mstore @change_users_domain.sql %NEWDOM% %OLDDOM% %PACNAM%


:EOF
CD /D %SCD%
ECHO.
ECHO   Started: %STRT%
ECHO Completed: !DATE! !TIME!
ECHO.
PAUSE
ENDLOCAL
EXIT /B 0
