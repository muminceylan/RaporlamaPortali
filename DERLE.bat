@echo off
echo.
echo ============================================
echo    RAPORLAMA PORTALI - DERLEME
echo    Dogus Cay - Afyon Seker Fabrikasi
echo ============================================
echo.

echo [1/3] NuGet paketleri yukleniyor...
dotnet restore

echo.
echo [2/3] Proje derleniyor...
dotnet publish -c Release -o ./publish

echo.
echo [3/3] Derleme tamamlandi!
echo.
echo ============================================
echo  EXE dosyasi: publish\RaporlamaPortali.exe
echo ============================================
echo.
echo Bu dosyayi istediginiz yere kopyalayabilirsiniz.
echo Cift tiklayinca uygulama acilacak.
echo.
pause
