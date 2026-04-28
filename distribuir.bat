@echo off
cd /d "%~dp0"
set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.8.9-hotspot"
set "JAR=target\s1210-extrator-1.1-jar-with-dependencies.jar"
set "DEST=dist"
set "NOME=SIGEP S-1210 Extrator"

if not exist "%JAR%" (
    echo JAR nao encontrado. Execute compilar.bat primeiro.
    pause
    exit /b 1
)

echo Removendo distribuicao anterior...
if exist "%DEST%" rd /s /q "%DEST%"

echo Gerando executavel portatil...
"%JAVA_HOME%\bin\jpackage.exe" ^
  --type app-image ^
  --input target ^
  --main-jar s1210-extrator-1.1-jar-with-dependencies.jar ^
  --name "%NOME%" ^
  --app-version 1.1 ^
  --dest "%DEST%" ^
  --icon src\main\resources\icon.ico ^
  --java-options "-Xmx512m"

if %ERRORLEVEL% neq 0 (
    echo ERRO ao gerar executavel.
    pause
    exit /b 1
)

echo.
echo Concluido!
echo Executavel em: %DEST%\%NOME%\%NOME%.exe
echo.
echo Para distribuir, compacte a pasta: %DEST%\%NOME%\
pause
