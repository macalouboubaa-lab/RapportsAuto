@echo off
REM Lance le processus complet : nettoyage du fichier ListTransactions + captures du dashboard par filiale
REM Tous les fichiers doivent se trouver dans C:\Users\OusseynouBoubacarCIS\RapportsAuto
REM (en dehors de Bureau/Documents a cause de l'Acces controle aux dossiers)
REM Utilise "uv" pour gerer Python et les librairies automatiquement (pas besoin d'installation manuelle)
REM Le resultat (fichier nettoye + captures) est range automatiquement dans un sous-dossier par heure (ex: 09H, 16H)

set DOSSIER_SCRIPT=%~dp0
set FICHIER_BRUT=%DOSSIER_SCRIPT%ListTransactions_GU3.xlsx
set DASHBOARD=%DOSSIER_SCRIPT%dashboard.html
set DOSSIER_SORTIE=%DOSSIER_SCRIPT:~0,-1%
set UV=C:\Users\OusseynouBoubacarCIS\.local\bin\uv.exe

REM Installe le navigateur Chromium pour Playwright (une seule fois, les fois suivantes c'est instantane)
"%UV%" run --with playwright playwright install chromium >> "%DOSSIER_SCRIPT%log_nettoyage.txt" 2>&1

"%UV%" run --with pandas --with openpyxl --with playwright "%DOSSIER_SCRIPT%executer_rapport.py" "%FICHIER_BRUT%" "%DASHBOARD%" "%DOSSIER_SORTIE%" >> "%DOSSIER_SCRIPT%log_nettoyage.txt" 2>&1
