@echo off
REM lancer.bat : lance la chaine SI CGPP (extraction, controle qualite, analyses, SitRep CGPP).
REM Double-clic, ou tache planifiee (chaque samedi a 18 h) :
REM   schtasks /create /sc weekly /d SAT /st 18:00 /tn SitRepCGPP /tr "C:\cgpp_si\lancer.bat"
REM Les options de pipeline.py sont transmises : lancer.bat --etapes qc,sitrep
setlocal
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PY=python
if exist ".venv\Scripts\python.exe" set PY=.venv\Scripts\python.exe
if not exist sorties mkdir sorties
echo ==== %date% %time% >> sorties\journal_lancer.log
"%PY%" pipeline.py %* >> sorties\journal_lancer.log 2>&1
if errorlevel 1 (
  echo Echec de la chaine : voir sorties\journal_lancer.log et sorties\journal_pipeline.log
  endlocal
  exit /b 1
)
echo Chaine terminee : SitRep dans sorties\sitrep, controle qualite dans sorties\qc
endlocal
