@echo off
REM ---------------------------------------------------------------------------
REM Recon Daily - Windows launcher (point Task Scheduler at this file)
REM Edit RECON_ROOT to your share folder, then this runs today's reconciliation.
REM ---------------------------------------------------------------------------

REM Path to the share folder that holds the dated sub-folders (YYYY-MM-DD).
REM For a synced SharePoint/OneDrive library this is a normal local path, e.g.:
REM   set "RECON_ROOT=C:\Users\%USERNAME%\Hubpay\Recon - Documents\Daily"
set "RECON_ROOT=C:\path\to\your\share\folder"

REM Days to look back when aging a break (default 10).
set "RECON_LOOKBACK=10"

cd /d "%~dp0"
python -m recon --root "%RECON_ROOT%" --lookback %RECON_LOOKBACK%
exit /b %ERRORLEVEL%
