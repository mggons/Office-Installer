@Echo off
Título Convirtiendo Office 2013 STD Retail a volumen
cd /d "%~dp0" && ( if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs" ) && fsutil dirty query %systemdrive% 1>nul 2>nul || (  echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/k cd ""%~sdp0"" && %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs" && "%temp%\getadmin.vbs" && exit /B )
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15"

cls
cscript ospp.vbs /remhst
cscript ospp.vbs /ckms-domain
ping 127.0.0.1 -n 1 > nul

cls
echo.-------------------------------------------------------------
echo        Installing Office Certificates Please Wait...
echo.-------------------------------------------------------------
ping 127.0.0.1 -n 3 > nul
echo Standard 2013
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\StandardVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\StandardVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\StandardVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\StandardVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\StandardVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\StandardVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\StandardVL_MAK-ul-phn.xrm-ms"
cscript ospp.vbs /inslic:"C:\ODT\2013\STD\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
echo Ha finalizado la conversión de licencias comerciales a licencias por volumen.
echo.-------------------------------------------------------------
echo	Instala su licencia de Office 2013 Standard
echo.-------------------------------------------------------------
SET /P LICENSE= (Type or paste the Office license 2013):
cscript //nologo ospp.vbs /inpkey:%LICENSE%
cscript //nologo ospp.vbs /unpkey:92CD4
cscript //nologo ospp.vbs /act
ECHO Espere un minuto mientras verificamos si se aplico tu licencia
cscript //nologo ospp.vbs /dstatus > C:\licenseOffice.txt
start C:\licenseOffice.txt

echo gracias por activar tu Office 2013 Standard
ping 127.0.0.1 -n 5 > nul