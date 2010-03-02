@REM .
@REM .                FULLTRUST FOR A DIRECTORY (CASPOL.EXE)
@REM .

@set _caspolpath="C:\Windows\Microsoft.NET\Framework\v2.0.50727\caspol.exe"
@set _targetdir="C:\Program Files (x86)\Fullenbaum\Word2003Tools4Dominique"
@set _codegroupname="DomAddIn"
@set _codegroupdesc="Word 2003 Add-In for Dominique"

@REM ___ Donner confiance à .NET 2 sur le répertoire
%_caspolpath% -u -ag All_Code -url %_targetdir%\* FullTrust -n %_codegroupname% -d %_codegroupdesc%
