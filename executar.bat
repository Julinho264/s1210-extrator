@echo off
cd /d "%~dp0"
set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.8.9-hotspot"
start "" "%JAVA_HOME%\bin\javaw" -jar target\s1210-extrator-1.0-jar-with-dependencies.jar
