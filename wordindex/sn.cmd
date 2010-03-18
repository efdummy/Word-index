@REM .
@REM .                STRONG NAMES (SN.EXE)
@REM .

@set _snpath="C:\Program Files\Microsoft SDKs\Windows\v6.1\Bin\sn.exe"

@REM ___ Générer une nouvelle paire de clés publique-privée
@REM %_snpath% -s wordindex.snk

@REM ___ Afficher le jeton de la clé publique contenue dans la DLL
%_snpath% -T bin\Release\Fullenbaum.wordindex.dll