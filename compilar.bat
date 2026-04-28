@echo off
cd /d "%~dp0"

set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.8.9-hotspot"
set "MVN=%USERPROFILE%\.m2\wrapper\dists\apache-maven-3.9.6-bin\3311e1d4\apache-maven-3.9.6\bin\mvn.cmd"

echo Compilando...
call "%MVN%" package -q
if %ERRORLEVEL% neq 0 (
    echo ERRO na compilacao. Verifique o log acima.
    pause
    exit /b 1
)
echo.
echo Compilado com sucesso!
echo Execute: executar.bat
pause
