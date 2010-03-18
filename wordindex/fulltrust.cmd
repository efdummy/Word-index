@REM .
@REM .                FULLTRUST FOR A DIRECTORY (CASPOL.EXE)
@REM .

@set _caspolpath="C:\Windows\Microsoft.NET\Framework\v2.0.50727\caspol.exe"
@set _targetdir="C:\Program Files (x86)\Fullenbaum\wordindex\*"
@set _codegroupname="WordIndexAddIn"
@set _codegroupdesc="Word index Word add-in"

@REM ___ Donner confiance à .NET 2 sur le répertoire
%_caspolpath% -q -u -ag All_Code -url %_targetdir% FullTrust -n %_codegroupname% -d %_codegroupdesc%
