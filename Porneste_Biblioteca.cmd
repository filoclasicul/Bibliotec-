@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Mediul Python .venv lipseste. Contacteaza persoana care a instalat aplicatia.
    pause
    exit /b 1
)
echo La prima pornire introdu adresa de e-mail si parola administratorului.
echo Parola nu se afiseaza in timp ce o tastezi.
echo Dupa pornirea serverului, deschide http://127.0.0.1:5000
echo Pentru editare, deschide http://127.0.0.1:5000/admin
echo Nu deschide admin.html prin previzualizarea de fisiere din VS Code.
echo Lasa aceasta fereastra deschisa cat timp folosesti site-ul.
echo.
".venv\Scripts\python.exe" app_admin.py
echo.
echo Serverul s-a oprit. Citeste eventualul mesaj de eroare de mai sus.
pause
