@echo off
echo.
echo 📊 MANUELLE PDF-INHALTSANALYSE
echo ═══════════════════════════════════════════════
echo.

REM Überprüfe ob PDF existiert
if not exist "Generalversammlung_output.pdf" (
    echo ❌ PDF nicht gefunden: Generalversammlung_output.pdf
    echo    Führe zuerst die Konvertierung aus!
    pause
    exit /b 1
)

REM Zeige Dateiinformationen
for %%f in ("Generalversammlung_output.pdf") do (
    echo 📄 Datei: %%~nxf
    echo 📏 Größe: %%~zf Bytes
    echo 📅 Erstellt: %%~tf
)

echo.
echo 🔧 Führe PDF-Inhaltsanalyse aus...
echo.

REM Führe Java-Analyzer aus (falls verfügbar)
java -cp "target\docconverter-1.0.0.jar" ch.std.doc.converter.app.PdfAnalyzer 2>nul

if errorlevel 1 (
    echo ⚠️  Java-Analyzer nicht verfügbar
    echo    Verwende PowerShell-Alternative...
    echo.
    
    REM PowerShell Alternative für grundlegende Analyse
    powershell -Command ^
        "$file = Get-Item 'Generalversammlung_output.pdf'; ^
         Write-Host '📊 BASIC FILE ANALYSIS' -ForegroundColor Green; ^
         Write-Host ('Size: {0} KB' -f [math]::Round($file.Length/1024, 2)); ^
         Write-Host ('Created: {0}' -f $file.CreationTime); ^
         Write-Host ('Modified: {0}' -f $file.LastWriteTime); ^
         if ($file.Length -gt 1000) { ^
             Write-Host '✅ PDF seems substantial' -ForegroundColor Green ^
         } else { ^
             Write-Host '⚠️  PDF seems quite small' -ForegroundColor Yellow ^
         }"
)

echo.
echo 📁 PDF-Datei: %CD%\Generalversammlung_output.pdf
echo.
echo ✅ Analyse abgeschlossen!
echo    Du kannst die PDF-Datei jetzt mit einem PDF-Viewer öffnen.
echo.
pause
