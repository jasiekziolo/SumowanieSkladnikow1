SUMOWANIE SKŁADNIKÓW EXCEL – BUILD WINDOWS

Najłatwiejsza droga do gotowego EXE bez instalowania czegokolwiek na komputerze firmowym:

1. Utwórz konto / zaloguj się na GitHub.
2. Utwórz nowe repozytorium, np. SumowanieSkladnikow.
3. Wgraj CAŁĄ zawartość tego folderu do repozytorium (razem z folderem .github).
4. Wejdź w zakładkę Actions.
5. Wybierz workflow „Build Windows EXE”.
6. Kliknij „Run workflow”.
7. Po zakończeniu kliknij wykonany workflow.
8. Na dole strony, w Artifacts, pobierz „SumowanieSkladnikow-Windows”.
9. W ZIP-ie będzie gotowy SumowanieSkladnikow.exe.

EXE jest budowany na Windows i uruchamia się bez Pythona, openpyxl i bez uprawnień administratora.

Aplikacja:
- odczytuje .xlsx,
- B = Numer składnika,
- D = Opis składnika,
- E = Ilość,
- F = Jednostka,
- grupuje po numerze + opisie + jednostce,
- sumuje ilości,
- pokazuje wynik w tabeli,
- eksportuje wynik do nowego .xlsx,
- zgłasza błędne wartości w kolumnie E.

Standardowy Tkinter nie ma natywnego Drag & Drop. Ta wersja używa bezpiecznego wyboru pliku przez okno dialogowe, bez dodatkowych bibliotek w gotowym EXE.
