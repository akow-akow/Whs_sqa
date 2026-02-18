# WSQA PRO + UPS Auto-Logic - Instrukcja Obsługi

Program służy do analizy braków w skanowaniu paczek na podstawie plików AK0 oraz weryfikacji ich statusu w systemie UPS.

## 1. Wybór folderu AK0
- Kliknij "1. WYBIERZ FOLDER AK0" i wskaż folder zawierający pliki Excel o nazwie zaczynającej się od "AK0" (np. AK0_15.02.2024.xlsx).
- Program wymaga minimum 2 plików, aby wyliczyć braki między dniami.

## 2. Wczytywanie Grafiku (Ranking Pracowników)
- Kliknij "2. WCZYTAJ GRAFIK".
- Otwórz swój grafik w Excelu, zaznacz obszar (Lokalizacja w pierwszej kolumnie, dni miesiąca w nagłówkach) i skopiuj (Ctrl+C).
- Wróć do programu i w oknie grafiku naciśnij **Ctrl+V**.
- Kliknij "ZAPISZ I ZAMKNIJ". Program będzie teraz wiedział, kto pracował na danej lokalizacji w dniu wystąpienia braku.

## 3. Konfiguracja UPS API
- Kliknij "USTAWIENIA UPS API" i wprowadź swoje dane (License Number, UserID, Password).
- Dane zostaną zapisane w pliku `ups_settings.ini` w folderze programu.

## 4. Generowanie Raportu
- Zaznacz magazyny, które chcesz przeanalizować.
- Zaznacz opcję "Automatycznie oznaczaj paczki poza Strykowem" (Auto-Green).
- Kliknij "3. GENERUJ RAPORT".

### Logika "Auto-Green":
Jeśli program wykryje brak skanu w ostatnim dniu, zapyta serwer UPS o status. Jeśli lokalizacja w UPS jest inna niż "Dobra" lub "Stryków", program uzna paczkę za wydaną (DORĘCZONA/WYDANA), oznaczy ją na zielono i nie doliczy błędu pracownikowi w rankingu.

### Wynik:
Nowy plik Excel zostanie zapisany w folderze źródłowym pod nazwą `Raport_AK0_Data_Godzina.xlsx`.
