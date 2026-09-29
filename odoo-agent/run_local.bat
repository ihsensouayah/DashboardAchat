@echo off
REM Lance la synchro depuis un PC du bureau (si Odoo n est pas joignable depuis Internet).
REM A cote de ce fichier : secrets.env (lignes NOM=valeur) et firebase.json
cd /d "%~dp0"
for /f "usebackq eol=# tokens=1,* delims==" %%A in ("secrets.env") do set "%%A=%%B"
set "FIREBASE_SERVICE_ACCOUNT_FILE=%~dp0firebase.json"
echo ==== %date% %time% >> sync.log
python odoo_sync.py >> sync.log 2>&1
