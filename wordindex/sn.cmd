@REM .
@REM .                STRONG NAMES (SN.EXE)
@REM .

@set _snpath="C:\Program Files\Microsoft SDKs\Windows\v6.1\Bin\sn.exe"

@REM ___ G�n�rer une nouvelle paire de cl�s publique-priv�e
@REM %_snpath% -s Word2003Tools4Dominique.snk

@REM ___ Afficher le jeton de la cl� publique contenue dans la DLL
%_snpath% -T bin\Release\Fullenbaum.Word2003Tools4Dominique.dll