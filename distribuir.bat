@echo off
cd /d "%~dp0"
set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.8.9-hotspot"
set "DEST=dist"
set "WARP=tools\warp-packer.exe"

for %%f in (target\s1210-extrator-*-jar-with-dependencies.jar) do set JAR=%%f
if not defined JAR (
    echo JAR nao encontrado. Execute compilar.bat primeiro.
    pause
    exit /b 1
)

rem Extrai versao do nome do JAR (ex: s1210-extrator-1.2-jar-with-dependencies.jar)
for %%f in (target\s1210-extrator-*-jar-with-dependencies.jar) do (
    set FNAME=%%~nf
)
set VERSAO=%FNAME:s1210-extrator-=%
set VERSAO=%VERSAO:-jar-with-dependencies=%
set "NOME=SIGEP S-1210 Extrator"

echo Versao: %VERSAO%
echo Removendo distribuicao anterior...
if exist "%DEST%" rd /s /q "%DEST%"

echo Gerando app-image...
"%JAVA_HOME%\bin\jpackage.exe" ^
  --type app-image ^
  --input target ^
  --main-jar %JAR:target\=% ^
  --name "%NOME%" ^
  --app-version %VERSAO% ^
  --dest "%DEST%" ^
  --icon src\main\resources\icon.ico ^
  --java-options "-Xmx512m -Xms64m -XX:+UseSerialGC"

if %ERRORLEVEL% neq 0 (
    echo ERRO ao gerar app-image.
    pause
    exit /b 1
)

echo Gerando executavel unico portatil...
if exist "%WARP%" (
    "%WARP%" --arch windows-x64 ^
        --input_dir "%DEST%\%NOME%" ^
        --exec "%NOME%.exe" ^
        --output "%DEST%\SIGEP_S1210_Extrator_v%VERSAO%.exe"

    if %ERRORLEVEL% neq 0 (
        echo AVISO: Falha ao gerar .exe unico.
    ) else (
        rem Corrige subsistema PE: console (3) -> GUI (2) — elimina janela CMD ao iniciar
        set "EXE_FINAL=%DEST%\SIGEP_S1210_Extrator_v%VERSAO%.exe"
        powershell -NonInteractive -Command "$p='%EXE_FINAL%'; $b=[IO.File]::ReadAllBytes($p); $pe=[BitConverter]::ToInt32($b,0x3C); $b[$pe+0x5C]=2; [IO.File]::WriteAllBytes($p,$b); Write-Host 'Subsistema corrigido: GUI (sem janela CMD)'"
        echo.
        echo === Executavel unico gerado ===
        echo %DEST%\SIGEP_S1210_Extrator_v%VERSAO%.exe
    )
) else (
    echo AVISO: warp-packer.exe nao encontrado em tools\
)

echo.
echo === Pasta portatil ===
echo %DEST%\%NOME%\%NOME%.exe
echo.
pause
