@echo off
REM Maven Wrapper Script mit automatischem JAVA_HOME Setup
REM Verwendung: maven-run.cmd [maven-kommando]

set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.7.6-hotspot"

if "%1"=="" (
    echo Verwendung: maven-run.cmd [maven-kommando]
    echo Beispiele:
    echo   maven-run.cmd --version
    echo   maven-run.cmd clean compile
    echo   maven-run.cmd test
    echo   maven-run.cmd clean package
    goto :end
)

echo JAVA_HOME: %JAVA_HOME%
echo Führe Maven-Kommando aus: %*
echo.

mvnw.cmd %*

:end
