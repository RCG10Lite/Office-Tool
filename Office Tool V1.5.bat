@echo off
@setlocal DisableDelayedExpansion
color 1F
 :Start
title Office Tool [By RCG10]
mode con cols=98 lines=30
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Install your Microsoft Office Version             ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] Activate your Microsoft Office                    ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Other                                             ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [4] Exit Script                                       ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' goto:OFI
if '%wahl%' == '2' goto:OFAc
if '%wahl%' == '3' goto:other
if '%wahl%' == '4' goto:close
goto:Start

 :other
title Office Tool [By RCG10]
mode con cols=98 lines=30
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Check if your Office is activated                 ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] Check Version and Edition this tool is running    ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Return to Main Menu                               ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' goto:IOFA
if '%wahl%' == '2' goto:WH
if '%wahl%' == '4' goto:close
goto:other



:WH
mode con cols=98 lines=30
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|       ==== Office Tool Version/Edition ====             ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|        Your Office Tool is on Version V1.5              ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|           and runs "Standart key" Edition               ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|      Press any key to Return to Start Menu...           ^|
Echo.                    ^|_________________________________________________________^|
pause >nul
goto:Start
 :OFAc
mode con cols=98 lines=30
fsutil dirty query %systemdrive%  >nul 2>&1 || (
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|                    ==== ERROR ====                      ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   This Action require administrator privileges.         ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   To do so, right click on this script and select       ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|                'Run as administrator'                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   Press any key to Return to Start Menu...              ^|
Echo.                    ^|_________________________________________________________^|
pause >nul
goto:Start
)
echo.  
echo                           What Version of Office do you want to Activate ?
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 365                              ^|
Echo.                    ^|   [2] Microsoft Office 2021                             ^|
Echo.                    ^|   [3] Microsoft Office 2019                             ^|
Echo.                    ^|   [4] Microsoft Office 2016                             ^|
Echo.                    ^|   [5] Microsoft Office 2013                             ^|
Echo.                    ^|   [6] Microsoft Office 2010                             ^|
Echo.                    ^|   [7] Microsoft Office 2007                             ^|
Echo.                    ^|   [8] Microsoft Office 2002 XP                          ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' goto:OF365Ac
if '%wahl%' == '2' goto:OF21Ac
if '%wahl%' == '3' goto:OF19Ac
if '%wahl%' == '4' goto:OF16Ac
if '%wahl%' == '5' goto:OF13Ac
if '%wahl%' == '6' goto:OF10Ac
if '%wahl%' == '7' goto:OF07Ac
if '%wahl%' == '8' goto:OF02Ac
if '%wahl%' == '100' goto:Start
goto:OFAc

:OF02Ac
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   Office XP 2002 dont uses Volume Activation so         ^|
Echo.                    ^|   the Product Key for The MS Office XP 2002             ^|
Echo.                    ^|   in this Script is:                                    ^|
Echo.                    ^|   FM9FY-TMF7Q-KCKCT-V9T29-TBBBG                         ^|
Echo.                    ^|   Just Copy out of this cmd                             ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '100' goto:Start
goto:OF02Ac


:OF07Ac
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   Office 2007 dont uses Volume Activation so            ^|
Echo.                    ^|   the Product Key for The MS Office 2007 Enterprise     ^|
Echo.                    ^|   in this Script is:                                    ^|
Echo.                    ^|   KGFVY-7733B-8WCK9-KTG64-BC7D8                         ^|
Echo.                    ^|   Just Copy out of this cmd                             ^|
Echo.                    ^|   (For English and German)                              ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '100' goto:Start
goto:OF07Ac

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

 :OF365Ac
title Activate Office 365 ProPlus for FREE - MSGuides.com&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products: Office 365 ProPlus (x86-x64)&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&choice /n /c YN /m "Would you like to visit my Website [Y,N]?" & if errorlevel 2 goto:Start) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://rcg10.webador.de"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.&echo Please try installing the latest version here: bit.ly/odt2k16
:halt
echo Press any key to return to Start Menu
pause >nul
goto:Start

  :OFI
mode con cols=98 lines=30
echo.                     __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 365 German                       ^|
Echo.                    ^|   [2] Microsoft Office 365 English                      ^|
Echo.                    ^|   [3] Microsoft Office 2021 German                      ^|
Echo.                    ^|   [4] Microsoft Office 2021 English                     ^|
Echo.                    ^|   [5] Microsoft Office 2019 German                      ^|
Echo.                    ^|   [6] Microsoft Office 2019 English                     ^|
Echo.                    ^|   [7] Microsoft Office 2016 German                      ^|
Echo.                    ^|   [8] Microsoft Office 2016 English                     ^|
Echo.                    ^|   [9] Microsoft Office 2013 German                      ^|
Echo.                    ^|   [10] Microsoft Office 2013 English                    ^|
Echo.                    ^|   [11] Microsoft Office 2010 German                     ^|
Echo.                    ^|   [12] Microsoft Office 2010 English                    ^|
Echo.                    ^|   [13] Microsoft Office 2007 English                    ^|
Echo.                    ^|   [13] Microsoft Office 2007 German                     ^|
Echo.                    ^|   [15] Microsoft Office XP 2002 German                  ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' goto:OF365G
if '%wahl%' == '2' goto:OF365E
if '%wahl%' == '3' goto:OF21G
if '%wahl%' == '4' goto:OF21E
if '%wahl%' == '5' goto:OF19G
if '%wahl%' == '6' goto:OF19E
if '%wahl%' == '7' goto:OF16G
if '%wahl%' == '8' goto:OF16E
if '%wahl%' == '9' goto:OF13G
if '%wahl%' == '10' goto:OF13E
if '%wahl%' == '11' goto:OF10G
if '%wahl%' == '12' goto:OF10E
if '%wahl%' == '13' goto:OF07E
if '%wahl%' == '14' goto:OF07G
if '%wahl%' == '15' goto:OF02xpG
if '%wahl%' == '100' goto:Start
goto:OFI

 :OF365G
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 365 German Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 365 Pro Plus                     ^|
Echo.                    ^|   [2] Microsoft Office 365 Home Premium                 ^|
Echo.                    ^|   [3] Microsoft Office 365 Business                     ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start Menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/O365BusinessRetail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/O365HomePremRetail.img
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/O365ProPlusRetail.img
if '%wahl%' == '100' goto:Start
goto:OF365G

 :OF365E
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 365 English Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 365 Pro Plus                     ^|
Echo.                    ^|   [2] Microsoft Office 365 Home Premium                 ^|
Echo.                    ^|   [3] Microsoft Office 365 Business                     ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start Menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '3' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/O365ProPlusRetail.img
if '%wahl%' == '2' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/O365HomePremRetail.img
if '%wahl%' == '1' Start https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-US/O365BusinessRetail.img
if '%wahl%' == '100' goto:Start
goto:OF365E

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
Echo.                    ^|   [10] Project Pro 2013                                 ^|
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
Echo.                    ^|   [10] Project Pro 2013                                 ^|
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

  :OF10G
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2010 German Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2010 Pro Plus (x64)              ^|
Echo.                    ^|   [2] Microsoft Office 2010 Pro Plus (x32)              ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://download1.winandoffice.com/Volume/office/2010/DE/Office_Professional_Plus_2010_64Bit_German.ISO
if '%wahl%' == '2' Start https://download1.winandoffice.com/Volume/office/2010/DE/Office_Professional_Plus_2010_32Bit_German.ISO
if '%wahl%' == '100' goto:Start
goto:OF10G

  :OF10E
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2010 English Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2010 Pro Plus (x64)              ^|
Echo.                    ^|   [2] Microsoft Office 2010 Pro Plus (x32)              ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://download1.winandoffice.com/Volume/office/2010/EN/MicrosoftOffice2010ProfessionalPlus64bit.ISO
if '%wahl%' == '2' Start https://download1.winandoffice.com/Volume/office/2010/EN/MicrosoftOffice2010ProfessionalPlus32bit.ISO
if '%wahl%' == '100' goto:Start
goto:OF10E

  :OF07E
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2007 English Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2007 Enterprise English (x64)    ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://drive.google.com/uc?id=1JwY6NZQdeRvKYRh2qAkW2q601c6BTJ8y&export=download&goto:of07keyE
if '%wahl%' == '100' goto:Start
goto:OF07E

:of07keyE
mode con cols=98 lines=30
echo.      
echo                                    Microsoft Office English 2007    
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   Password for Extracting is 123SL                      ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   Extract to Desktop not to the TEMP Folder             ^|
Echo.                    ^|   Then after extracting run setup.exe                   ^|
Echo.                    ^|   then enter Product Key:                               ^|
Echo.                    ^|   Product Key is: KGFVY-7733B-8WCK9-KTG64-BC7D8         ^|
Echo.                    ^|   and it Installs Microsoft Office 2007 Enterprise      ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '100' goto:Start
goto:of07keyE

  :OF07G
mode con cols=98 lines=30
echo.      
echo                                 Microsoft Office 2007 German Downloads   
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office 2007 Enterprise German (x64)     ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://mega.nz/file/p4JXBQLZ#AGoTaQ5Cku-IAQ2X2esl0LYeVzLanSBU-hcXz911ZZM&goto:of07keyG
if '%wahl%' == '100' goto:Start
goto:OF07G

 :of07keyG
mode con cols=98 lines=30
echo.      
echo                                    Microsoft Office German 2007    
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   Extract the Folder with WinRAR                        ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   Then after extracting run setup.exe                   ^|
Echo.                    ^|   then enter Product Key:                               ^|
Echo.                    ^|   Product Key is: KGFVY-7733B-8WCK9-KTG64-BC7D8         ^|
Echo.                    ^|   and it Installs Microsoft Office 2007 Enterprise      ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '100' goto:Start
goto:of07keyG

 :OF02xpG
mode con cols=98 lines=30
echo.      
echo                                    Microsoft Office XP 2002 German    
echo.                      __________________________________________________________
Echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Microsoft Office XP 2002 German                   ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [100] Return to Start menu                            ^|
Echo.                    ^|_________________________________________________________^|
SET /p wahl=
if '%wahl%' == '1' Start https://www.mediafire.com/file/f0ngas8hmkbts6v/Office+XP+2002++[by+RCG10].zip/file
if '%wahl%' == '100' goto:Start
goto:OF02xpG


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

:IOFA
set WMI_VBS=0

@cls
set "_cmdf=%~f0"
if exist "%SystemRoot%\Sysnative\cmd.exe" (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" "
exit /b
)
if exist "%SystemRoot%\SysArm32\cmd.exe" if /i %PROCESSOR_ARCHITECTURE%==AMD64 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" "
exit /b
)
color 1F
title Check Activation Status [wmi]
set wspp=SoftwareLicensingProduct
set wsps=SoftwareLicensingService
set ospp=OfficeSoftwareProtectionProduct
set osps=OfficeSoftwareProtectionService
set winApp=55c92734-d682-4d71-983e-d6ec3f16059f
set o14App=59a52881-a989-479d-af46-f275c6370663
set o15App=0ff1ce15-a989-479d-af46-f275c6370663
for %%# in (spp_get,ospp_get,cW1nd0ws,sppw,c0ff1ce15,sppo,osppsvc,ospp14,ospp15) do set "%%#="
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set "spp_get=Description, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, EvaluationEndDate, GracePeriodRemaining, ID, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceProductKeyID, LicenseStatus, LicenseStatusReason, Name, PartialProductKey, ProductKeyID, VLActivationInterval, VLRenewalInterval"
set "ospp_get=%spp_get%"
if %winbuild% GEQ 9200 set "spp_get=%spp_get%, KeyManagementServiceLookupDomain, VLActivationTypeEnabled"
if %winbuild% GEQ 9600 set "spp_get=%spp_get%, DiscoveredKeyManagementServiceMachineIpAddress, ProductKeyChannel"
set "_work=%~dp0"
set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_Local=%LocalAppData%"
set _Identity=0
setlocal EnableDelayedExpansion
dir /b /s /a:-d "!_Local!\Microsoft\Office\Licenses\*1*" 1>nul 2>nul && set _Identity=1
dir /b /s /a:-d "!ProgramData!\Microsoft\Office\Licenses\*1*" 1>nul 2>nul && set _Identity=1
pushd "!_work!"
setlocal DisableDelayedExpansion
if %winbuild% LSS 9200 if not exist "%SystemRoot%\servicing\Packages\Microsoft-Windows-PowerShell-WTR-Package~*.mum" set _Identity=0
set _pwrsh=1
if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set _pwrsh=0
set "_csg=cscript.exe //NoLogo //Job:WmiMulti "%~nx0?.wsf""
set "_csq=cscript.exe //NoLogo //Job:WmiQuery "%~nx0?.wsf""
set "_csx=cscript.exe //NoLogo //Job:XPDT "%~nx0?.wsf""
if %winbuild% GEQ 22483 set WMI_VBS=1
if %WMI_VBS% EQU 0 (
set "_zz1=wmic path"
set "_zz2=where"
set "_zz3=get"
set "_zz4=/value"
set "_zz5=("
set "_zz6=)"
set "_zz7="wmic path"
set "_zz8=/value""
) else (
set "_zz1=%_csq%"
set "_zz2="
set "_zz3="
set "_zz4="
set "_zz5=""
set "_zz6=""
set "_zz7=%_csq%"
set "_zz8="
)

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
set "line2=************************************************************"
set "line3=____________________________________________________________"

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
reg query HKU\S-1-5-19 1>nul 2>nul && (
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
)
if %_WSH% EQU 0 if %WMI_VBS% NEQ 0 goto :E_VBS

set OsppHook=1
sc query osppsvc >nul 2>&1
if %errorlevel% EQU 1060 set OsppHook=0

net start sppsvc /y >nul 2>&1
call :casWpkey %wspp% %winApp% cW1nd0ws sppw
if %winbuild% GEQ 9200 call :casWpkey %wspp% %o15App% c0ff1ce15 sppo
if %OsppHook% NEQ 0 (
net start osppsvc /y >nul 2>&1
call :casWpkey %ospp% %o14App% osppsvc ospp14
if %winbuild% LSS 9200 call :casWpkey %ospp% %o15App% osppsvc ospp15
)


:casWcon
set winID=0
set verbose=1
if not defined c0ff1ce15 (
if defined osppsvc goto :casWospp
goto :casWend
)
echo %line2%
echo ***                   Office Status                      ***
echo %line2%
set "_qr=%_zz7% %wspp% %_zz2% %_zz5%ApplicationID='%o15App%' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz8%"
for /f "tokens=2 delims==" %%# in ('%_qr%') do (
  set "chkID=%%#"
  call :casWdet "%wspp%" "%wsps%" "%spp_get%"
  call :casWout
  echo %line3%
  echo.
)
set verbose=0
if defined osppsvc goto :casWospp
goto :casWend

:casWospp
if %verbose% EQU 1 (
echo %line2%
echo ***                   Office Status                      ***
echo %line2%
)
set "_qr=%_zz7% %ospp% %_zz2% %_zz5%ApplicationID='%o15App%' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz8%"
if defined ospp15 for /f "tokens=2 delims==" %%# in ('%_qr%') do (
  set "chkID=%%#"
  call :casWdet "%ospp%" "%osps%" "%ospp_get%"
  call :casWout
  echo %line3%
  echo.
)
set "_qr=%_zz7% %ospp% %_zz2% %_zz5%ApplicationID='%o14App%' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz8%"
if defined ospp14 for /f "tokens=2 delims==" %%# in ('%_qr%') do (
  set "chkID=%%#"
  call :casWdet "%ospp%" "%osps%" "%ospp_get%"
  call :casWout
  echo %line3%
  echo.
)
goto :casWend

:casWpkey
set "_qr=%_zz1% %1 %_zz2% %_zz5%ApplicationID='%2' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz4%"
%_qr% 2>nul | findstr /i ID 1>nul && (set %3=1&set %4=1)
exit /b

:casWdet
for %%# in (%~3) do set "%%#="
if /i %~1==%ospp% for %%# in (DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled) do set "%%#="
set "cKmsClient="
set "cTblClient="
set "cAvmClient="
set "ExpireMsg="
set "_xpr="
set "_qr="wmic path %~1 where ID='%chkID%' get %~3 /value" ^| findstr ^="
if %WMI_VBS% NEQ 0 set "_qr=%_csg% %~1 "ID='%chkID%'" "%~3""
for /f "tokens=* delims=" %%# in ('%_qr%') do set "%%#"

set /a _gpr=(GracePeriodRemaining+1440-1)/1440
echo %Description%| findstr /i VOLUME_KMSCLIENT 1>nul && (set cKmsClient=1&set _mTag=Volume)
echo %Description%| findstr /i TIMEBASED_ 1>nul && (set cTblClient=1&set _mTag=Timebased)
echo %Description%| findstr /i VIRTUAL_MACHINE_ACTIVATION 1>nul && (set cAvmClient=1&set _mTag=Automatic VM)
cmd /c exit /b %LicenseStatusReason%
set "LicenseReason=%=ExitCode%"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%_gpr% day(s))"
if %_gpr% GEQ 1 if %_WSH% EQU 1 (
for /f "tokens=* delims=" %%# in ('%_csx% %GracePeriodRemaining%') do set "_xpr=%%#"
)
if %_gpr% GEQ 1 if %_pwrsh% EQU 1 if not defined _xpr (
for /f "tokens=* delims=" %%# in ('powershell "$([DateTime]::Now.addMinutes(%GracePeriodRemaining%)).ToString('yyyy-MM-dd HH:mm:ss')" 2^>nul') do set "_xpr=%%#"
title Check Activation Status [wmi]
)

if %LicenseStatus% EQU 0 (
set "License=Unlicensed"
set "LicenseMsg="
)
if %LicenseStatus% EQU 1 (
set "License=Licensed"
set "LicenseMsg="
if %GracePeriodRemaining% EQU 0 (
  if %winID% EQU 1 (set "ExpireMsg=The machine is permanently activated.") else (set "ExpireMsg=The product is permanently activated.")
  ) else (
  set "LicenseMsg=%_mTag% activation expiration: %GracePeriodRemaining% minute(s) (%_gpr% day(s))"
  if defined _xpr set "ExpireMsg=%_mTag% activation will expire %_xpr%"
  )
)
if %LicenseStatus% EQU 2 (
set "License=Initial grace period"
if defined _xpr set "ExpireMsg=Initial grace period ends %_xpr%"
)
if %LicenseStatus% EQU 3 (
set "License=Additional grace period (KMS license expired or hardware out of tolerance)"
if defined _xpr set "ExpireMsg=Additional grace period ends %_xpr%"
)
if %LicenseStatus% EQU 4 (
set "License=Non-genuine grace period."
if defined _xpr set "ExpireMsg=Non-genuine grace period ends %_xpr%"
)
if %LicenseStatus% EQU 6 (
set "License=Extended grace period"
if defined _xpr set "ExpireMsg=Extended grace period ends %_xpr%"
)
if %LicenseStatus% EQU 5 (
set "License=Notification"
  if "%LicenseReason%"=="C004F200" (set "LicenseMsg=Notification Reason: 0xC004F200 (non-genuine)."
  ) else if "%LicenseReason%"=="C004F009" (set "LicenseMsg=Notification Reason: 0xC004F009 (grace time expired)."
  ) else (set "LicenseMsg=Notification Reason: 0x%LicenseReason%"
  )
)
if %LicenseStatus% GTR 6 (
set "License=Unknown"
set "LicenseMsg="
)
if not defined cKmsClient exit /b

if %KeyManagementServicePort%==0 set KeyManagementServicePort=1688
set "KmsReg=Registered KMS machine name: %KeyManagementServiceMachine%:%KeyManagementServicePort%"
if "%KeyManagementServiceMachine%"=="" set "KmsReg=Registered KMS machine name: KMS name not available"

if %DiscoveredKeyManagementServiceMachinePort%==0 set DiscoveredKeyManagementServiceMachinePort=1688
set "KmsDns=KMS machine name from DNS: %DiscoveredKeyManagementServiceMachineName%:%DiscoveredKeyManagementServiceMachinePort%"
if "%DiscoveredKeyManagementServiceMachineName%"=="" set "KmsDns=DNS auto-discovery: KMS name not available"

set "_qr="wmic path %~2 get ClientMachineID, KeyManagementServiceHostCaching /value" ^| findstr ^="
if %WMI_VBS% NEQ 0 set "_qr=%_csg% %~2 "ClientMachineID, KeyManagementServiceHostCaching""
for /f "tokens=* delims=" %%# in ('%_qr%') do set "%%#"
if /i %KeyManagementServiceHostCaching%==True (set KeyManagementServiceHostCaching=Enabled) else (set KeyManagementServiceHostCaching=Disabled)

if %winbuild% LSS 9200 exit /b
if /i %~1==%ospp% exit /b

if "%KeyManagementServiceLookupDomain%"=="" set "KeyManagementServiceLookupDomain="

if %VLActivationTypeEnabled% EQU 3 (
set VLActivationType=Token
) else if %VLActivationTypeEnabled% EQU 2 (
set VLActivationType=KMS
) else if %VLActivationTypeEnabled% EQU 1 (
set VLActivationType=AD
) else (
set VLActivationType=All
)

if %winbuild% LSS 9600 exit /b
if "%DiscoveredKeyManagementServiceMachineIpAddress%"=="" set "DiscoveredKeyManagementServiceMachineIpAddress=not available"
exit /b

:casWout
echo.
echo Name: %Name%
echo Description: %Description%
echo Activation ID: %ID%
echo Extended PID: %ProductKeyID%
if defined ProductKeyChannel echo Product Key Channel: %ProductKeyChannel%
echo Partial Product Key: %PartialProductKey%
echo License Status: %License%
if defined LicenseMsg echo %LicenseMsg%
if not %LicenseStatus%==0 if not %EvaluationEndDate:~0,8%==16010101 echo Evaluation End Date: %EvaluationEndDate:~0,4%-%EvaluationEndDate:~4,2%-%EvaluationEndDate:~6,2% %EvaluationEndDate:~8,2%:%EvaluationEndDate:~10,2% UTC
if not defined cKmsClient (
if defined ExpireMsg echo.&echo.    %ExpireMsg%
exit /b
)
if defined VLActivationTypeEnabled echo Configured Activation Type: %VLActivationType%
echo.
if not %LicenseStatus%==1 (
echo Please activate the product in order to update KMS client information values.
exit /b
)
echo Most recent activation information:
echo Key Management Service client information
echo.    Client Machine ID (CMID): %ClientMachineID%
echo.    %KmsDns%
echo.    %KmsReg%
if defined DiscoveredKeyManagementServiceMachineIpAddress echo.    KMS machine IP address: %DiscoveredKeyManagementServiceMachineIpAddress%
echo.    KMS machine extended PID: %KeyManagementServiceProductKeyID%
echo.    Activation interval: %VLActivationInterval% minutes
echo.    Renewal interval: %VLRenewalInterval% minutes
echo.    KMS host caching: %KeyManagementServiceHostCaching%
if defined KeyManagementServiceLookupDomain echo.    KMS SRV record lookup domain: %KeyManagementServiceLookupDomain%
if defined ExpireMsg echo.&echo.    %ExpireMsg%
exit /b

:casWend
if %_Identity% EQU 1 if %_pwrsh% EQU 1 (
echo %line2%
echo ***                  Office vNext Status                 ***
echo %line2%
setlocal EnableDelayedExpansion
powershell "$f=[IO.File]::ReadAllText('!_batp!') -split ':vNextDiag\:.*';iex ($f[1])"
title Check Activation Status [wmi]
echo %line3%
echo.
)
echo.
echo Press any key to exit.
pause >nul
exit /b

:E_VBS
echo ==== ERROR ====
echo Windows Script Host is disabled.
echo It is required for this script to work.
echo.
echo Press any key to exit.
pause >nul
exit /b

:vNextDiag:
function PrintModePerPridFromRegistry
{
	$vNextRegkey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext"
	$vNextPrids = Get-Item -Path $vNextRegkey -ErrorAction Ignore | Select-Object -ExpandProperty 'property' | Where-Object -FilterScript {$_ -Ne 'InstalledGraceKey' -And $_ -Ne 'MigrationToV5Done' -And $_ -Ne 'test' -And $_ -Ne 'unknown'}
	If ($vNextPrids -Eq $null)
	{
		Write-Host "No registry keys found."
		Return
	}
	$vNextPrids | ForEach `
	{
		$mode = (Get-ItemProperty -Path $vNextRegkey -Name $_).$_
		Switch ($mode)
		{
			2 { $mode = "vNext"; Break }
			3 { $mode = "Device"; Break }
			Default { $mode = "Legacy"; Break }
		}
		Write-Host $_ = $mode
	}
}
function PrintSharedComputerLicensing
{
	$scaRegKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$scaValue = Get-ItemProperty -Path $scaRegKey -ErrorAction Ignore | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction Ignore
	$scaRegKey2 = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing"
	$scaValue2 = Get-ItemProperty -Path $scaRegKey2 -ErrorAction Ignore | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction Ignore
	$scaPolicyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing"
	$scaPolicyValue = Get-ItemProperty -Path $scaPolicyKey -ErrorAction Ignore | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction Ignore
	If ($scaValue -Eq $null -And $scaValue2 -Eq $null -And $scaPolicyValue -Eq $null)
	{
		Write-Host "No registry keys found."
		Return
	}
	$scaModeValue = $scaValue -Or $scaValue2 -Or $scaPolicyValue
	If ($scaModeValue -Eq 0)
	{
		$scaMode = "Disabled"
	}
	If ($scaModeValue -Eq 1)
	{
		$scaMode = "Enabled"
	}
	Write-Host "SharedComputerLicensing" = $scaMode
	Write-Host
	$tokenFiles = $null
	$tokenPath = "${env:LOCALAPPDATA}\Microsoft\Office\16.0\Licensing"
	If (Test-Path $tokenPath)
	{
		$tokenFiles = Get-ChildItem -Path $tokenPath -Recurse -File -Filter "*authString*"
	}
	If ($tokenFiles.length -Eq 0)
	{
		Write-Host "No tokens found."
		Return
	}
	$tokenFiles | ForEach `
	{
		$tokenParts = (Get-Content -Encoding Unicode -Path $_.FullName).Split('_')
		$output = [PSCustomObject] `
			@{
				ACID = $tokenParts[0];
				User = $tokenParts[3]
				NotBefore = $tokenParts[4];
				NotAfter = $tokenParts[5];
			} | ConvertTo-Json
		Write-Host $output
	}
}
function PrintLicensesInformation
{
	Param(
		[ValidateSet("NUL", "Device")]
		[String]$mode
	)
	If ($mode -Eq "NUL")
	{
		$licensePath = "${env:LOCALAPPDATA}\Microsoft\Office\Licenses"
	}
	ElseIf ($mode -Eq "Device")
	{
		$licensePath = "${env:PROGRAMDATA}\Microsoft\Office\Licenses"
	}
	$licenseFiles = $null
	If (Test-Path $licensePath)
	{
		$licenseFiles = Get-ChildItem -Path $licensePath -Recurse -File
	}
	If ($licenseFiles.length -Eq 0)
	{
		Write-Host "No licenses found."
		Return
	}
	$licenseFiles | ForEach `
	{
		$license = (Get-Content -Encoding Unicode $_.FullName | ConvertFrom-Json).License
		$decodedLicense = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($license)) | ConvertFrom-Json
		$licenseType = $decodedLicense.LicenseType
		$userId = $decodedLicense.Metadata.UserId
		$identitiesRegkey = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity\Identities\${userId}*" -ErrorAction Ignore
		$licenseState = $null
		If ((Get-Date) -Gt (Get-Date $decodedLicense.MetaData.NotAfter))
		{
			$licenseState = "RFM"
		}
		ElseIf (($decodedLicense.ExpiresOn -Eq $null) -Or
			((Get-Date) -Lt (Get-Date $decodedLicense.ExpiresOn)))
		{
			$licenseState = "Licensed"
		}
		Else
		{
			$licenseState = "Grace"
		}
		if ($mode -Eq "NUL")
		{
			$output = [PSCustomObject] `
			@{
				Version = $_.Directory.Name
				Type = "User|${licenseType}";
				Product = $decodedLicense.ProductReleaseId;
				Acid = $decodedLicense.Acid;
				LicenseState = $licenseState;
				EntitlementStatus = $decodedLicense.Status;
				ReasonCode = $decodedLicense.ReasonCode;
				NotBefore = $decodedLicense.Metadata.NotBefore;
				NotAfter = $decodedLicense.Metadata.NotAfter;
				NextRenewal = $decodedLicense.Metadata.RenewAfter;
				Expiration = $decodedLicense.ExpiresOn;
				TenantId = $decodedLicense.Metadata.TenantId;
			} | ConvertTo-Json
		}
		ElseIf ($mode -Eq "Device")
		{
			$output = [PSCustomObject] `
			@{
				Version = $_.Directory.Name
				Type = "Device|${licenseType}";
				Product = $decodedLicense.ProductReleaseId;
				Acid = $decodedLicense.Acid;
				DeviceId = $decodedLicense.Metadata.DeviceId;
				LicenseState = $licenseState;
				EntitlementStatus = $decodedLicense.Status;
				ReasonCode = $decodedLicense.ReasonCode;
				NotBefore = $decodedLicense.Metadata.NotBefore;
				NotAfter = $decodedLicense.Metadata.NotAfter;
				NextRenewal = $decodedLicense.Metadata.RenewAfter;
				Expiration = $decodedLicense.ExpiresOn;
				TenantId = $decodedLicense.Metadata.TenantId;
			} | ConvertTo-Json
		}
		Write-Output $output
	}
}
	Write-Host
	Write-Host "========== Mode per ProductReleaseId =========="
	Write-Host
PrintModePerPridFromRegistry
	Write-Host
	Write-Host "========== Shared Computer Licensing =========="
	Write-Host
PrintSharedComputerLicensing
	Write-Host
	Write-Host "========== vNext licenses =========="
	Write-Host
PrintLicensesInformation -Mode "NUL"
	Write-Host
	Write-Host "========== Device licenses =========="
	Write-Host
PrintLicensesInformation -Mode "Device"
:vNextDiag:
color 1F
----- Begin wsf script --->
<package>
   <job id="WmiQuery">
      <script language="VBScript">
         If WScript.Arguments.Count = 3 Then
            wExc = "Select " & WScript.Arguments.Item(2) & " from " & WScript.Arguments.Item(0) & " where " & WScript.Arguments.Item(1)
            wGet = WScript.Arguments.Item(2)
         Else
            wExc = "Select " & WScript.Arguments.Item(1) & " from " & WScript.Arguments.Item(0)
            wGet = WScript.Arguments.Item(1)
         End If
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For Each objItm in objCol
            For each Prop in objItm.Properties_
               If LCase(Prop.Name) = LCase(wGet) Then
                  WScript.Echo Prop.Name & "=" & Prop.Value
                  Exit For
               End If
            Next
         Next
      </script>
   </job>
   <job id="WmiMulti">
      <script language="VBScript">
         If WScript.Arguments.Count = 3 Then
            wExc = "Select " & WScript.Arguments.Item(2) & " from " & WScript.Arguments.Item(0) & " where " & WScript.Arguments.Item(1)
         Else
            wExc = "Select " & WScript.Arguments.Item(1) & " from " & WScript.Arguments.Item(0)
         End If
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For Each objItm in objCol
            For each Prop in objItm.Properties_
               WScript.Echo Prop.Name & "=" & Prop.Value
            Next
         Next
      </script>
   </job>
   <job id="XPDT">
      <script language="VBScript">
         WScript.Echo DateAdd("n", WScript.Arguments.Item(0), Now)
      </script>
   </job>
color 1F
</package>
color 1F