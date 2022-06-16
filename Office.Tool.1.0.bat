@echo off
color 1F
::===========================================================================
fsutil dirty query %systemdrive%  >nul 2>&1 || (
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
exit
)
 :Start
title Office Tool [By RCG10]
mode con cols=98 lines=30
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Install your Microsoft Office Version             ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] Activate your Microsoft Office                    ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Exit Script                                       ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' goto:OFI
if '%wahl%' == '2' goto:OFAc
if '%wahl%' == '3' goto:close
goto:Start

 :OFAc
mode con cols=98 lines=30
echo.  
echo                           What Version of Office do you want to Activate ?
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2021                             ^|
Echo.                    ^|   [2] Microsoft Office 2019                             ^|
Echo.                    ^|   [3] Microsoft Office 2016                             ^|
Echo.                    ^|   [4] Microsoft Office 2013                             ^|
Echo.                    ^|   [5] Microsoft Office 2010                             ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' goto:OF21Ac
if '%wahl%' == '2' goto:OF19Ac
if '%wahl%' == '3' goto:OF16Ac
if '%wahl%' == '4' goto:OF13Ac
if '%wahl%' == '5' goto:OF10Ac
if '%wahl%' == '100' goto:Start
goto:OFAc

 :OF10Ac
title Activate Microsoft Office 2010 Volume for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo – Microsoft Office 2010 Standard Volume&echo – Microsoft Office 2010 Professional Plus Volume&echo.&echo.&(if exist “%ProgramFiles%\Microsoft Office\Office14\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office14”)&(if exist “%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office14”)&echo.&echo ============================================================================&echo Activating your Office…&cscript //nologo ospp.vbs /unpkey:8R6BM >nul&cscript //nologo ospp.vbs /unpkey:H3GVB >nul&cscript //nologo ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul&cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i “successful” && (echo.&echo============================================================================&choice /n /c YN /m “Would you like to visit my Website [Y,N]?” & if errorlevel 2 goto:Start) || (echo The connection to my KMS server failed! Trying to connect to another one… & echo Please wait… & echo. & echo. & set /a i+=1 & goto server)
explorer “http://rcg10.webador.de”&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.
:halt
echo Press any key to return to Start Menu
pause >nul
goto:Start

 :OF13Ac
title Activate Microsoft Office 2013 Volume for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo - Microsoft Office 2013 Standard Volume&echo - Microsoft Office 2013 Professional Plus Volume&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15")&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:92CD4 >nul&cscript //nologo ospp.vbs /unpkey:GVGXT >nul&cscript //nologo ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4 >nul&cscript //nologo ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo============================================================================&choice /n /c YN /m "Would you like to visit my Website [Y,N]?" & if errorlevel 2 goto:Start) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://rcg10.webador.de"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.
:halt
echo Press any key to return to Start Menu
pause >nul
goto:Start

 :OF16Ac
title Activate Microsoft Office 2016 ALL versions for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2016&echo - Microsoft Office Professional Plus 2016&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&cscript //nologo ospp.vbs /unpkey:CPQVG >nul&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo============================================================================&choice /n /c YN /m "Would you like to visit my Website [Y,N]?" & if errorlevel 2 goto:Start) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://rcg10.webador.de"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.&echo Please try installing the latest version here: bit.ly/downloadmsp
:halt
echo Press any key to return to Start Menu
pause >nul
goto:Start

 :OF19Ac
title Activate Microsoft Office 2019 ALL versions for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2019&echo - Microsoft Office Professional Plus 2019&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo============================================================================&choice /n /c YN /m "Would you like to visit my Website [Y,N]?" & if errorlevel 2 goto:Start) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://rcg10.webador.de"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.&echo Please try installing the latest version here: bit.ly/aiomsp
:halt
echo Press anykey to Return to the Start Menu
pause >nul
goto:Start

 :OF21Ac
mode con cols=98 lines=30
title Activate Microsoft Office 2021 (ALL versions) for FREE - [rcg10lite]&cls&echo =====================================================================================&echo #Project: Activating Microsoft software products for FREE without additional software&echo =====================================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2021&echo - Microsoft Office Professional Plus 2021&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo =====================================================================================&echo Activating your product...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6F7TH >nul&set i=1&cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul||goto notsupported
:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=s8.uk.to
if %i% EQU 3 set KMS=s9.us.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo =====================================================================================&echo.&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo=====================================================================================&choice /n /c YN /m "Would you like to visit my Website [Y,N]?" & if errorlevel 2 goto:Start) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto skms)
explorer "http://www.rcg10.webador.de"&goto halt
:notsupported
echo =====================================================================================&echo.&echo Sorry, your version is not supported.&echo.&goto halt
:busy
echo =====================================================================================&echo.&echo Sorry, the server is busy and can't respond to your request. Please try again.&echo.
:halt
Press any Key to return to Start menu
pause >nul
goto:Start

  :OFI
mode con cols=98 lines=30
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2021 German                      ^|
Echo.                    ^|   [2] Microsoft Office 2021 English                     ^|
Echo.                    ^|   [3] Microsoft Office 2019 German                      ^|
Echo.                    ^|   [4] Microsoft Office 2019 English                     ^|
Echo.                    ^|   [5] Microsoft Office 2016 German                      ^|
Echo.                    ^|   [6] Microsoft Office 2016 English                     ^|
Echo.                    ^|   [7] Microsoft Office 2013 German                      ^|
Echo.                    ^|   [8] Microsoft Office 2013 English                     ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' goto:OF21G
if '%wahl%' == '1' goto:OF21E
if '%wahl%' == '3' goto:OF19G
if '%wahl%' == '4' goto:OF19E
if '%wahl%' == '5' goto:OF16G
if '%wahl%' == '6' goto:OF16E
if '%wahl%' == '7' goto:OF13G
if '%wahl%' == '8' goto:OF13E
if '%wahl%' == '9' goto:OF10G
if '%wahl%' == '10' goto:OF10E
if '%wahl%' == '100' goto:Start
goto:OFI

 :OF21G
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2021 German Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2021 Pro Plus                    ^|
Echo.                    ^|   [2] Microsoft Office 2021 Home Business               ^|
Echo.                    ^|   [3] Word 2021                                         ^|
Echo.                    ^|   [4] Excel 2021                                        ^|
Echo.                    ^|   [5] Powerpoint 2021                                   ^|
Echo.                    ^|   [6] Outlook 2021                                      ^|
Echo.                    ^|   [7] Publisher 2021                                    ^|
Echo.                    ^|   [8] Access 2021                                       ^|
Echo.                    ^|   [9] Project Pro 2021                                  ^|
Echo.                    ^|   [10] Visio Pro 2021                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ProPlus2021Retail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/HomeBusiness2021Retail.img
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Word2021Retail.img
if '%wahl%' == '4' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Excel2021Retail.img
if '%wahl%' == '5' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/PowerPoint2021Retail.img
if '%wahl%' == '6' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Outlook2021Retail.img
if '%wahl%' == '7' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Publisher2021Retail.img
if '%wahl%' == '8' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Access2021Retail.img
if '%wahl%' == '9' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ProjectPro2021Retail.img
if '%wahl%' == '10' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/VisioPro2021Retail.img
if '%wahl%' == '100' goto:Start
goto:OF21G

 :OF21E
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2021 English Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2021 Pro Plus (+MS Teams)        ^|
Echo.                    ^|   [2] Microsoft Office 2021 Home Business (+MS Teams)   ^|
Echo.                    ^|   [3] Word 2021                                         ^|
Echo.                    ^|   [4] Excel 2021                                        ^|
Echo.                    ^|   [5] Powerpoint 2021                                   ^|
Echo.                    ^|   [6] Outlook 2021                                      ^|
Echo.                    ^|   [7] Publisher 2021                                    ^|
Echo.                    ^|   [8] Access 2021                                       ^|
Echo.                    ^|   [9] Project Pro 2021                                  ^|
Echo.                    ^|   [10] Visio Pro 2021                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/ProPlus2021Retail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/HomeBusiness2021Retail.img
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Word2021Retail.img
if '%wahl%' == '4' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Excel2021Retail.img
if '%wahl%' == '5' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/PowerPoint2021Retail.img
if '%wahl%' == '6' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Outlook2021Retail.img
if '%wahl%' == '7' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Publisher2021Retail.img
if '%wahl%' == '8' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Access2021Retail.img
if '%wahl%' == '9' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/ProjectPro2021Retail.img
if '%wahl%' == '10' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/VisioPro2021Retail.img
if '%wahl%' == '100' goto:Start
goto:OF21E

 :OF19G
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2019 German Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2019 Pro Plus (+MS Teams)        ^|
Echo.                    ^|   [2] Microsoft Office 2019 Home Business (+MS Teams)   ^|
Echo.                    ^|   [3] Word 2019                                         ^|
Echo.                    ^|   [4] Excel 2019                                        ^|
Echo.                    ^|   [5] Powerpoint 2019                                   ^|
Echo.                    ^|   [6] Outlook 2019                                      ^|
Echo.                    ^|   [7] Publisher 2019                                    ^|
Echo.                    ^|   [8] Access 2019                                       ^|
Echo.                    ^|   [9] Project Pro 2019                                  ^|
Echo.                    ^|   [10] Visio Pro 2019                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ProPlus2019Retail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/HomeBusiness2019Retail.img
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Word2019Retail.img
if '%wahl%' == '4' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Excel2019Retail.img
if '%wahl%' == '5' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/PowerPoint2019Retail.img
if '%wahl%' == '6' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Outlook2019Retail.img
if '%wahl%' == '7' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Publisher2019Retail.img
if '%wahl%' == '8' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Access2019Retail.img
if '%wahl%' == '9' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ProjectPro2019Retail.img
if '%wahl%' == '10' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/VisioPro2019Retail.img
if '%wahl%' == '100' goto:Start
goto:OF19G


 :OF19E
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2019 English Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2019 Pro Plus (+MS Teams)        ^|
Echo.                    ^|   [2] Microsoft Office 2019 Home Business (+MS Teams)   ^|
Echo.                    ^|   [3] Word 2019                                         ^|
Echo.                    ^|   [4] Excel 2019                                        ^|
Echo.                    ^|   [5] Powerpoint 2019                                   ^|
Echo.                    ^|   [6] Outlook 2019                                      ^|
Echo.                    ^|   [7] Publisher 2019                                    ^|
Echo.                    ^|   [8] Access 2019                                       ^|
Echo.                    ^|   [9] Project Pro 2019                                  ^|
Echo.                    ^|   [10] Visio Pro 2019                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/ProPlus2019Retail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/HomeBusiness2019Retail.img
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Word2019Retail.img
if '%wahl%' == '4' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Excel2019Retail.img
if '%wahl%' == '5' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/PowerPoint2019Retail.img
if '%wahl%' == '6' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Outlook2019Retail.img
if '%wahl%' == '7' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Publisher2019Retail.img
if '%wahl%' == '8' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/Access2019Retail.img
if '%wahl%' == '9' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/ProjectPro2019Retail.img
if '%wahl%' == '10' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/VisioPro2019Retail.img
if '%wahl%' == '100' goto:Start
goto:OF19GE

  :OF16G
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2016 German Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2016 Pro Plus                    ^|
Echo.                    ^|   [2] Microsoft Office 2016 Home Business               ^|
Echo.                    ^|   [3] Word 2016                                         ^|
Echo.                    ^|   [4] Excel 2016                                        ^|
Echo.                    ^|   [5] Powerpoint 2016                                   ^|
Echo.                    ^|   [6] Outlook 2016                                      ^|
Echo.                    ^|   [7] Publisher 2016                                    ^|
Echo.                    ^|   [8] Access 2016                                       ^|
Echo.                    ^|   [9] Project Pro 2016                                  ^|
Echo.                    ^|   [10] Visio Pro 2016                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ProPlusRetail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/HomeBusinessRetail.img
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/WordRetail.img
if '%wahl%' == '4' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ExcelRetail.img
if '%wahl%' == '5' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/PowerPointRetail.img
if '%wahl%' == '6' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/OutlookRetail.img
if '%wahl%' == '7' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/PublisherRetail.img
if '%wahl%' == '8' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/AccessRetail.img
if '%wahl%' == '9' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ProjectProRetail.img
if '%wahl%' == '10' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/VisioProRetail.img
if '%wahl%' == '100' goto:Start
goto:OF16G


  :OF16E
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2016 English Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2016 Pro Plus                    ^|
Echo.                    ^|   [2] Microsoft Office 2016 Home Business               ^|
Echo.                    ^|   [3] Word 2016                                         ^|
Echo.                    ^|   [4] Excel 2016                                        ^|
Echo.                    ^|   [5] Powerpoint 2016                                   ^|
Echo.                    ^|   [6] Outlook 2016                                      ^|
Echo.                    ^|   [7] Publisher 2016                                    ^|
Echo.                    ^|   [8] Access 2016                                       ^|
Echo.                    ^|   [9] Project Pro 2016                                  ^|
Echo.                    ^|   [10] Visio Pro 2016                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/ProPlusRetail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/HomeBusinessRetail.img
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/WordRetail.img
if '%wahl%' == '4' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/ExcelRetail.img
if '%wahl%' == '5' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/PowerPointRetail.img
if '%wahl%' == '6' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/OutlookRetail.img
if '%wahl%' == '7' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/PublisherRetail.img
if '%wahl%' == '8' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/AccessRetail.img
if '%wahl%' == '9' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/ProjectProRetail.img
if '%wahl%' == '10' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/VisioProRetail.img
if '%wahl%' == '100' goto:Start
goto:OF16E

  :OF13G
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2013 German Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2013 Pro                         ^|
Echo.                    ^|   [2] Microsoft Office 2013 Home Business               ^|
Echo.                    ^|   [3] Word 2013                                         ^|
Echo.                    ^|   [4] Excel 2013                                        ^|
Echo.                    ^|   [5] Powerpoint 2013                                   ^|
Echo.                    ^|   [6] One Note 2013                                     ^|
Echo.                    ^|   [7] Outlook 2013                                      ^|
Echo.                    ^|   [8] Publisher 2013                                    ^|
Echo.                    ^|   [9] Access 2013                                       ^|
Echo.                    ^|   [10] Project Pro 2013 2013                            ^|
Echo.                    ^|   [11] Visio Pro 2013                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=ProfessionalRetail
if '%wahl%' == '2' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=HomeBusinessRetail
if '%wahl%' == '3' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=WordRetail
if '%wahl%' == '4' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=ExcelRetail
if '%wahl%' == '5' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=PowerPointRetail
if '%wahl%' == '6' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=OneNoteRetail
if '%wahl%' == '7' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=OutlookRetail
if '%wahl%' == '8' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=PublisherRetail
if '%wahl%' == '9' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=AccessRetail
if '%wahl%' == '10' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=ProjectProRetail
if '%wahl%' == '11' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=de-DE&p3=VisioProRetail
if '%wahl%' == '100' goto:Start
goto:OF13G

  :OF13E
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2013 English Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2013 Pro                         ^|
Echo.                    ^|   [2] Microsoft Office 2013 Home Business               ^|
Echo.                    ^|   [3] Word 2013                                         ^|
Echo.                    ^|   [4] Excel 2013                                        ^|
Echo.                    ^|   [5] Powerpoint 2013                                   ^|
Echo.                    ^|   [6] One Note 2013                                     ^|
Echo.                    ^|   [7] Outlook 2013                                      ^|
Echo.                    ^|   [8] Publisher 2013                                    ^|
Echo.                    ^|   [9] Access 2013                                       ^|
Echo.                    ^|   [10] Project Pro 2013 2013                            ^|
Echo.                    ^|   [11] Visio Pro 2013                                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=ProfessionalRetail
if '%wahl%' == '2' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=HomeBusinessRetail
if '%wahl%' == '3' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=WordRetail
if '%wahl%' == '4' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=ExcelRetail
if '%wahl%' == '5' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=PowerPointRetail
if '%wahl%' == '6' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=OneNoteRetail
if '%wahl%' == '7' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=OutlookRetail
if '%wahl%' == '8' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=PublisherRetail
if '%wahl%' == '9' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=AccessRetail
if '%wahl%' == '10' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=ProjectProRetail
if '%wahl%' == '11' Start https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=en-US&p3=VisioProRetail
if '%wahl%' == '100' goto:Start
goto:OF13E

:close
mode con cols=98 lines=30
echo.
echo.
echo.
echo.                        ===========================================
echo.                                                                   
echo.                                    Created by @RCG10
echo.                                                                    
echo.                        ===========================================
echo.
echo.
echo Press any key to Exit.
pause > nul
exit