@echo off
:: Get the current user's name
set "username=%USERNAME%"
echo Hello, %username%!
echo.

:: Get the current folder path
set "folderPath=%~dp0"
echo Current folder path is: %folderPath%
echo.

:: Run PowerShell commands within the batch file
powershell -command ^
    "$word = New-Object -ComObject Word.Application;" ^
    "$word.Visible = $false;" ^
    "$files = Get-ChildItem -Path '%folderPath%' -Filter *.docx;" ^
    "foreach ($file in $files) {" ^
    "    Write-Host 'Converting' $file.Name 'to PDF...';" ^
    "    $doc = $word.Documents.Open($file.FullName);" ^
    "    $pdfName = [System.IO.Path]::ChangeExtension($file.FullName, 'pdf');" ^
    "    $doc.SaveAs([ref]$pdfName, [ref]17);" ^
    "    $doc.Close();" ^
    "    Write-Host 'Conversion complete:' $pdfName -ForegroundColor Cyan;" ^
    "    Write-Host '';" ^
    "};" ^
    "$word.Quit();"

:: Print the final message in green using PowerShell
powershell -command "Write-Host 'All conversions are complete.' -ForegroundColor Green"
echo.
pause
