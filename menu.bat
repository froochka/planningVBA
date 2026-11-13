@echo off
rem ==========================================================================
rem  Menu Git du projet : a placer a la racine du depot, a cote du classeur.
rem  Double-clic pour l'ouvrir (ou Excel : Complements > VbaSync > Menu Git).
rem  Il travaille toujours dans SON dossier : impossible de se tromper de projet.
rem  Fichier en ASCII (pas d'accents) avec fins de ligne Windows.
rem ==========================================================================
setlocal
cd /d "%~dp0"
title Menu Git - %~dp0

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo Ce dossier n'est pas un depot Git : %CD%
    pause
    exit /b 1
)

:menu
cls
set BRANCHE=
for /f "delims=" %%b in ('git branch --show-current') do set BRANCHE=%%b
set DEPOT=
for /f "delims=" %%r in ('git config --get remote.origin.url') do set DEPOT=%%r
echo ==========================================================================
echo   Dossier : %CD%
echo   Depot   : %DEPOT%
echo   Branche : %BRANCHE%
echo ==========================================================================
echo.
echo   1 - Recuperer les mises a jour (git pull)
echo   2 - Etat : fichiers modifies
echo   3 - Envoyer mes modifications (commit + push)
echo   4 - Historique (10 dernieres versions)
echo   5 - Ouvrir le classeur
echo   6 - Console Git dans ce dossier
echo   Q - Quitter
echo.
choice /c 123456Q /n /m "Votre choix : "
if errorlevel 7 goto :eof
if errorlevel 6 goto console
if errorlevel 5 goto ouvrir
if errorlevel 4 goto historique
if errorlevel 3 goto envoyer
if errorlevel 2 goto etat
if errorlevel 1 goto maj
goto menu

rem --------------------------------------------------------------------------
:maj
echo.
call :excel_ferme || goto pause
rem Le classeur est souvent modifie par un simple import du code (ImporterVBA) :
rem on propose d'annuler ces changements pour que le pull puisse passer.
git diff --quiet -- "*.xlsm"
if errorlevel 1 (
    echo Le classeur .xlsm a ete modifie depuis la derniere mise a jour.
    echo  - si c'est seulement un import du code : repondez O, le code sera reimporte apres ;
    echo  - si vous avez modifie des feuilles a garder : repondez N, puis utilisez le choix 3.
    choice /c ON /m "Annuler les modifications du classeur"
    if errorlevel 2 goto pause
    git checkout -- "*.xlsm"
)
git pull
if errorlevel 1 goto pause
echo.
echo Version recue :
git log --oneline -1
echo.
echo Ensuite : ouvrez le classeur (choix 5), puis Complements ^> VbaSync ^> Importer le code.
goto pause

rem --------------------------------------------------------------------------
:etat
echo.
git status
goto pause

rem --------------------------------------------------------------------------
:envoyer
echo.
call :excel_ferme || goto pause
git status -s
echo.
set MSG=
set /p MSG=Message decrivant vos modifications (vide = annuler) :
if "%MSG%"=="" goto menu
git add -A
git commit -m "%MSG%"
git push
goto pause

rem --------------------------------------------------------------------------
:historique
echo.
git log --oneline -10
goto pause

rem --------------------------------------------------------------------------
:ouvrir
for %%f in (*.xlsm) do (
    start "" "%%f"
    goto menu
)
echo Aucun classeur .xlsm dans ce dossier.
goto pause

rem --------------------------------------------------------------------------
:console
start "Git - %CD%" cmd /k "git status"
goto menu

rem --------------------------------------------------------------------------
:pause
echo.
pause
goto menu

rem --------------------------------------------------------------------------
rem Verifie qu'aucun classeur .xlsm du dossier n'est ouvert dans Excel
rem (un fichier ouvert dans Excel ne peut pas etre ouvert en ecriture).
:excel_ferme
set OUVERT=
for %%f in (*.xlsm) do (
    2>nul ( >>"%%f" (call ) ) || set OUVERT=%%f
)
if defined OUVERT (
    echo Le classeur %OUVERT% est ouvert dans Excel : enregistrez-le et fermez-le d'abord.
    exit /b 1
)
exit /b 0
