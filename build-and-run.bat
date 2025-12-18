@echo off@echo off@echo off

:: ═══════════════════════════════════════════════════════════════════════════════

:: XNRGY ENGINEERING AUTOMATION TOOLS - BUILD & RUN SCRIPT (BAT):: ═══════════════════════════════════════════════════════════════════════════════:: ═══════════════════════════════════════════════════════════════════════════════

:: Version: 1.0.0

:: Auteur: Smart Tools Amine - XNRGY Climate Systems ULC:: VAULT AUTOMATION TOOL - BUILD & RUN SCRIPT:: VAULT AUTOMATION TOOL - BUILD & RUN SCRIPT (BAT)

:: ═══════════════════════════════════════════════════════════════════════════════

:: Version: 1.0.0:: ═══════════════════════════════════════════════════════════════════════════════

setlocal enabledelayedexpansion

title XNRGY Engineering Automation Tools - Build ^& Run:: Auteur: Smart Tools Amine - XNRGY Climate Systems ULC:: Version: 2.0

cd /d "%~dp0"

:: ═══════════════════════════════════════════════════════════════════════════════:: Auteur: Smart Tools Amine - XNRGY Climate Systems ULC

echo.

echo ===============================================================:: Date: 2025-12-17

echo   XNRGY ENGINEERING AUTOMATION TOOLS - BUILD ^& RUN

echo ===============================================================setlocal enabledelayedexpansion:: ═══════════════════════════════════════════════════════════════════════════════

echo.

title Vault Automation Tool - Build and Run

:: ═══════════════════════════════════════════════════════════════════════════════

:: ETAPE 1: Forcer l'arret de l'application si elle est en courscd /d "%~dp0"setlocal enabledelayedexpansion

:: ═══════════════════════════════════════════════════════════════════════════════

echo [1/3] Arret des instances existantes...title Vault Automation Tool - Build ^& Run



tasklist /FI "IMAGENAME eq XnrgyEngineeringAutomationTools.exe" 2>NUL | find /I "XnrgyEngineeringAutomationTools.exe" >NULecho.cd /d "%~dp0"

if %ERRORLEVEL%==0 (

    echo       Instance trouvee, arret en cours...echo =========================================================

    taskkill /F /IM "XnrgyEngineeringAutomationTools.exe" >NUL 2>&1

    timeout /t 2 /nobreak >NULecho   VAULT AUTOMATION TOOL - BUILD ^& RUNecho.

    echo       [OK] Instance arretee

) else (echo =========================================================echo ===============================================================

    echo       Aucune instance en cours

)echo.echo   VAULT AUTOMATION TOOL - BUILD ^& RUN

echo.

echo ===============================================================

:: ═══════════════════════════════════════════════════════════════════════════════

:: ETAPE 2: Trouver MSBuild:: ═══════════════════════════════════════════════════════════════════════════════echo.

:: ═══════════════════════════════════════════════════════════════════════════════

echo [2/3] Compilation en mode Release...:: ÉTAPE 1: Forcer l'arrêt de l'application si elle est en cours



set "MSBUILD=":: ═══════════════════════════════════════════════════════════════════════════════:: ═══════════════════════════════════════════════════════════════════════════════



:: Chercher MSBuild VS 2022 Enterpriseecho [1/3] Arret des instances existantes...:: ÉTAPE 1: Forcer l'arrêt de l'application si elle est en cours

if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe" (

    set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe":: ═══════════════════════════════════════════════════════════════════════════════

)

tasklist /FI "IMAGENAME eq VaultAutomationTool.exe" 2>NUL | find /I "VaultAutomationTool.exe" >NULecho [1/4] Arret des instances existantes...

:: Chercher MSBuild VS 2022 Professional

if not defined MSBUILD (if %ERRORLEVEL%==0 (

    if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe" (

        set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe"    echo       Instance trouvee, arret en cours...:: Vérifier si l'application est en cours

    )

)    taskkill /F /IM VaultAutomationTool.exe >NUL 2>&1tasklist /FI "IMAGENAME eq VaultAutomationTool.exe" 2>NUL | find /I "VaultAutomationTool.exe" >NUL



:: Chercher MSBuild VS 2022 Community    timeout /t 2 /nobreak >NULif %ERRORLEVEL%==0 (

if not defined MSBUILD (

    if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" (    echo       [OK] Instance arretee    echo       Instance trouvee, arret en cours...

        set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"

    )) else (    taskkill /F /IM VaultAutomationTool.exe >NUL 2>&1

)

    echo       Aucune instance en cours    timeout /t 2 /nobreak >NUL

if not defined MSBUILD (

    echo       [ERREUR] MSBuild non trouve!)    echo       [OK] Instance arretee

    echo       Installez Visual Studio 2022 avec le workload .NET Desktop

    pauseecho.) else (

    exit /b 1

)    echo       Aucune instance en cours



echo       MSBuild: %MSBUILD%:: ═══════════════════════════════════════════════════════════════════════════════)

echo       Compilation...

:: ÉTAPE 2: Trouver MSBuild et Compilerecho.

"%MSBUILD%" XnrgyEngineeringAutomationTools.csproj /p:Configuration=Release /p:Platform=x64 /t:Rebuild /v:minimal /nologo /m

:: ═══════════════════════════════════════════════════════════════════════════════

if %ERRORLEVEL% neq 0 (

    echo.echo [2/3] Compilation en mode Release...:: ═══════════════════════════════════════════════════════════════════════════════

    echo       [ERREUR] Echec de la compilation!

    pause:: ÉTAPE 2: Trouver MSBuild (VS 2022)

    exit /b 1

)set "MSBUILD=":: ═══════════════════════════════════════════════════════════════════════════════



echo       [OK] Compilation reussie!echo [2/4] Recherche de MSBuild (VS 2022)...

echo.

if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe" (

:: ═══════════════════════════════════════════════════════════════════════════════

:: ETAPE 3: Lancer l'application    set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe"set "MSBUILD="

:: ═══════════════════════════════════════════════════════════════════════════════

echo [3/3] Lancement de l'application...    goto :compile



set "EXE_PATH=bin\Release\XnrgyEngineeringAutomationTools.exe"):: Enterprise



if not exist "%EXE_PATH%" (if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe" (if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe" (

    echo       [ERREUR] Executable introuvable: %EXE_PATH%

    pause    set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe"    set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe"

    exit /b 1

)    goto :compile    echo       Trouve: VS 2022 Enterprise



echo       Chemin: %EXE_PATH%)    goto :found_msbuild

start "" "%EXE_PATH%"

timeout /t 2 /nobreak >NULif exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" ()

echo       [OK] Application lancee!

    set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"

echo.

echo ===============================================================    goto :compile:: Professional

echo   TERMINE

echo ===============================================================)if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe" (

echo.

    set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe"

endlocal

echo       [ERREUR] MSBuild (VS 2022) introuvable!    echo       Trouve: VS 2022 Professional

pause    goto :found_msbuild

exit /b 1)



:compile:: Community

echo       MSBuild: %MSBUILD%if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" (

    set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"

"%MSBUILD%" VaultAutomationTool.csproj /p:Configuration=Release /t:Rebuild /v:minimal /nologo /m    echo       Trouve: VS 2022 Community

    goto :found_msbuild

if %ERRORLEVEL% NEQ 0 ()

    echo.

    echo       [ERREUR] Echec de la compilation:: BuildTools

    pauseif exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\amd64\MSBuild.exe" (

    exit /b 1    set "MSBUILD=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\amd64\MSBuild.exe"

)    echo       Trouve: VS 2022 BuildTools

    goto :found_msbuild

echo.)

echo       [OK] Compilation reussie!

echo.:: Utiliser vswhere si disponible

if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" (

if not exist "bin\Release\VaultAutomationTool.exe" (    for /f "usebackq tokens=*" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -version "[17.0,18.0)" -property installationPath -latest`) do (

    echo       [ERREUR] Executable introuvable        set "VSINSTALL=%%i"

    pause    )

    exit /b 1    if defined VSINSTALL (

)        if exist "!VSINSTALL!\MSBuild\Current\Bin\amd64\MSBuild.exe" (

            set "MSBUILD=!VSINSTALL!\MSBuild\Current\Bin\amd64\MSBuild.exe"

:: ═══════════════════════════════════════════════════════════════════════════════            echo       Trouve via vswhere

:: ÉTAPE 3: Lancement de l'application            goto :found_msbuild

:: ═══════════════════════════════════════════════════════════════════════════════        )

echo [3/3] Lancement de l'application...    )

echo       Chemin: bin\Release\VaultAutomationTool.exe)



start "" "bin\Release\VaultAutomationTool.exe"echo.

echo [ERREUR] MSBuild (VS 2022) introuvable!

timeout /t 2 /nobreak >NULecho          Installez Visual Studio 2022 avec '.NET desktop development'

echo.

tasklist /FI "IMAGENAME eq VaultAutomationTool.exe" 2>NUL | find /I "VaultAutomationTool.exe" >NULpause

if %ERRORLEVEL%==0 (exit /b 1

    echo       [OK] Application lancee!

) else (:found_msbuild

    echo       [ATTENTION] L'application ne semble pas s'etre lanceeecho.

)

:: ═══════════════════════════════════════════════════════════════════════════════

echo.:: ÉTAPE 3: Compilation

echo =========================================================:: ═══════════════════════════════════════════════════════════════════════════════

echo   TERMINEecho [3/4] Compilation en mode Release...

echo =========================================================echo       Projet: VaultAutomationTool.csproj

echo.echo       Platform: x64

echo       Framework: .NET 4.8
echo.

"%MSBUILD%" VaultAutomationTool.csproj /p:Configuration=Release /p:Platform=x64 /t:Rebuild /v:minimal /nologo /m /nr:false

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERREUR] Echec de la compilation (Code: %ERRORLEVEL%)
    echo.
    pause
    exit /b 1
)

echo.
echo       [OK] Compilation reussie!
echo.

:: Vérifier que l'exécutable existe
if not exist "bin\Release\VaultAutomationTool.exe" (
    echo [ERREUR] Executable introuvable: bin\Release\VaultAutomationTool.exe
    pause
    exit /b 1
)

:: ═══════════════════════════════════════════════════════════════════════════════
:: ÉTAPE 4: Lancement de l'application
:: ═══════════════════════════════════════════════════════════════════════════════
echo [4/4] Lancement de l'application...
echo       Chemin: bin\Release\VaultAutomationTool.exe
echo.

start "" "bin\Release\VaultAutomationTool.exe"

timeout /t 2 /nobreak >NUL

:: Vérifier que l'application s'est lancée
tasklist /FI "IMAGENAME eq VaultAutomationTool.exe" 2>NUL | find /I "VaultAutomationTool.exe" >NUL
if %ERRORLEVEL%==0 (
    echo       [OK] Application lancee avec succes!
) else (
    echo       [ATTENTION] L'application ne semble pas s'etre lancee
)

echo.
echo ===============================================================
echo   TERMINE
echo ===============================================================
echo.




