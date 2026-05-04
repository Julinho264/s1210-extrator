@echo off
cd /d "%~dp0"
set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.8.9-hotspot"

for %%f in (target\s1210-extrator-*-jar-with-dependencies.jar) do set JAR=%%f

if not defined JAR (
    echo JAR nao encontrado. Execute compilar.bat primeiro.
    pause
    exit /b 1
)

start "" "%JAVA_HOME%\bin\javaw" -Xmx512m -Xms64m -XX:+UseSerialGC -jar "%JAR%"
