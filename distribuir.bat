@echo off
cd /d "%~dp0"
set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.8.9-hotspot"
set "JAR=target\s1210-extrator-1.1-jar-with-dependencies.jar"
set "DEST=dist"
set "NOME=SIGEP S-1210 Extrator"
set "WARP=tools\warp-packer.exe"

if not exist "%JAR%" (
    echo JAR nao encontrado. Execute compilar.bat primeiro.
    pause
    exit /b 1
)

echo Removendo distribuicao anterior...
if exist "%DEST%" rd /s /q "%DEST%"

echo Gerando app-image (pasta portatil)...
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
    echo ERRO ao gerar app-image.
    pause
    exit /b 1
)

echo Gerando executavel unico portatil (.exe)...
if exist "%WARP%" (
    "%WARP%" --arch windows-x64 ^
        --input_dir "%DEST%\%NOME%" ^
        --exec "%NOME%.exe" ^
        --output "%DEST%\SIGEP_S1210_Extrator.exe"

    if %ERRORLEVEL% neq 0 (
        echo AVISO: Falha ao gerar .exe unico. A pasta portatil ainda esta disponivel.
    ) else (
        echo.
        echo === Executavel unico gerado ===
        echo %DEST%\SIGEP_S1210_Extrator.exe
    )
) else (
    echo AVISO: warp-packer.exe nao encontrado em tools\. Pulando geracao do .exe unico.
)

echo.
echo === Pasta portatil gerada ===
echo %DEST%\%NOME%\%NOME%.exe
echo.
pause
