@Echo off
Title Convirtiendo Office 2016 PRO Retail a volumen
cd /d "%~dp0" && ( if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs" ) && fsutil dirty query %systemdrive% 1>nul 2>nul || (  echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/k cd ""%~sdp0"" && %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs" && "%temp%\getadmin.vbs" && exit /B )

if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

cls
cscript ospp.vbs /remhst
cscript ospp.vbs /ckms-domain
ping 127.0.0.1 -n 1 > nul

cls
echo.-------------------------------------------------------------
echo        Installing Office Certificates Please Wait...
echo.-------------------------------------------------------------
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms"
ping 127.0.0.1 -n 5 > nul
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cls
echo Ha finalizado la conversión de licencias comerciales a licencias por volumen.
echo.-------------------------------------------------------------
echo	Instala su licencia de Office 2016 Pro Plus
echo.-------------------------------------------------------------
SET /P LICENSE= (Type or paste the Office license 2016):
cscript //nologo ospp.vbs /inpkey:%LICENSE%
cscript //nologo ospp.vbs /unpkey:WFG99
cscript //nologo ospp.vbs /unpkey:BTDRB
cscript //nologo ospp.vbs /act
ECHO Espere un minuto mientras verificamos si se aplico tu licencia
cscript //nologo ospp.vbs /dstatus > C:\licenseOffice.txt
start C:\licenseOffice.txt

echo gracias por activar tu Office 2016 Pro Plus
ping 127.0.0.1 -n 5 > nul
exit