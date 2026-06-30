@echo off
chcp 65001 >nul
setlocal

rem ============================================================
rem  Kopiowanie etykiet (.btw) wg listy produktow z Excela.
rem
rem  Sposob uzycia:
rem    1) Najprosciej: umiesc obok tego pliku swoj Excel (.xlsx
rem       lub .csv) oraz pliki .btw i kliknij dwukrotnie ten .bat.
rem    2) Lub przeciagnij plik Excel na ikone tego .bat.
rem    3) Lub uruchom z parametrami z wiersza polecen, np.:
rem         kopiuj_etykiety.bat "C:\dane\produkty.xlsx" "C:\etykiety" "C:\wynik"
rem
rem  Kolejnosc parametrow (wszystkie opcjonalne):
rem    %%1 = plik Excel/CSV
rem    %%2 = folder z plikami .btw
rem    %%3 = folder docelowy (kopie)
rem    %%4 = rozszerzenie plikow (domyslnie btw)
rem    %%5 = litera kolumny z nazwami produktow (domyslnie A)
rem    %%6 = prog podobienstwa 0-100 dla dopasowania rozmytego (domyslnie 75)
rem
rem  UWAGA: domyslnie wlaczone jest dopasowanie ROZMYTE (radzi sobie z
rem  podobnymi, nie identycznymi nazwami). Jesli za duzo lapie lub myli
rem  pliki - podnies prog (np. 90). Jesli za malo lapie - obniz (np. 65).
rem ============================================================

set "PS1=%~dp0kopiuj_etykiety.ps1"

set "ARGS="
if not "%~1"=="" set "ARGS=%ARGS% -ExcelPath ""%~1"""
if not "%~2"=="" set "ARGS=%ARGS% -SourceFolder ""%~2"""
if not "%~3"=="" set "ARGS=%ARGS% -TargetFolder ""%~3"""
if not "%~4"=="" set "ARGS=%ARGS% -Extension ""%~4"""
if not "%~5"=="" set "ARGS=%ARGS% -Column ""%~5"""
if not "%~6"=="" set "ARGS=%ARGS% -Threshold %~6"

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"%ARGS%

echo.
echo Gotowe. Nacisnij dowolny klawisz, aby zamknac to okno...
pause >nul
endlocal
