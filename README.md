📦 AK0 Warehouse Scan Quality Analyzer
AK0 Warehouse Scan Quality Analyzer to wyspecjalizowane narzędzie stworzone dla działów logistyki i operacji magazynowych (HUB Stryków / Dobra). Program automatyzuje proces wykrywania brakujących skanów w codziennych raportach inwentaryzacyjnych (AK0) i przypisuje odpowiedzialność za dany sektor na podstawie grafiku pracowników.
🚀 Główne Funkcje
•	Analiza historyczna: Porównuje wiele plików Excel z różnych dni, aby wykryć paczki, które przestały być skanowane (zniknęły z inwentarza).
•	Inteligentne mapowanie grafiku: Automatycznie przypisuje imię i nazwisko osoby odpowiedzialnej za dany sektor (Magazyn/Smalls) w konkretnym dniu.
•	Weryfikacja UPS API (Auto-Green): Automatycznie sprawdza status paczki w systemie UPS. Jeśli paczka ma status "Delivered" poza HUBem, program oznacza ją jako bezpieczną (zieloną).
•	Obsługa przesyłek zwolnionych (Released): Możliwość wczytania pliku WHOFILEXPT.DAT lub wklejenia tekstu z systemu, aby oznaczyć paczki, które opuściły magazyn, ale nie mają jeszcze skanu doręczenia.
•	Generowanie raportów Excel: Tworzy przejrzysty arkusz z historią skanów, statusami UPS i komentarzami dotyczącymi personelu.
________________________________________
🛠 Instrukcja Obsługi
1. Przygotowanie plików
Program szuka plików Excel w wybranym folderze. Pliki muszą:
•	Zaczynać się od frazy "AK0".
•	Zawierać datę w formacie dd.MM.yyyy (np. AK0_raport_19.02.2026.xlsx).
2. Konfiguracja grafiku (Krok 2a)
Kliknij przycisk "WCZYTAJ GRAFIK" i wklej dane z arkusza grafiku (Ctrl+V). Program obsługuje:
•	Sektory typu MAG 1, MAG 2 itp. (mapowane na IWMAGAZYN / EWMAGEXP).
•	Sektory Smalls (automatyczne łączenie dwóch pracowników z sąsiadujących wierszy).
3. Przesyłki Zwolnione (Krok 2b)
Jeśli posiadasz listę przesyłek, które przeszły przez "Release", kliknij "PRZESYŁKI ZWOLNIONE". Możesz:
•	Wskazać plik systemowy WHOFILEXPT.DAT.
•	Wkleić surowy tekst z raportu.
Paczki te zostaną oznaczone w raporcie kolorem jasnoniebieskim.
4. Generowanie Raportu (Krok 3)
•	Wybierz magazyny, które Cię interesują (Filtry Import/Export).
•	(Opcjonalnie) Zaznacz "Automatyczna weryfikacja UPS", jeśli masz skonfigurowane API.
•	Kliknij "GENERUJ RAPORT". Wynikowy plik Excel pojawi się w folderze z raportami AK0.
________________________________________
🎨 Legenda kolorów w raporcie
Kolor	Znaczenie
Biały	Paczka obecna na stanie (zeskanowana).
Czerwony (Salmon)	BRAK SKANU – paczka powinna być, a jej nie ma.
Zielony	DORĘCZONA – UPS potwierdza doręczenie (paczka bezpieczna).
Jasnoniebieski	RELEASED – paczka zwolniona do wyjazdu (znaleziona w pliku .DAT).
________________________________________
⚙️ Wymagania techniczne
•	System operacyjny: Windows 10/11.
•	Biblioteki: .NET Framework 4.7.2+.
•	Zależności: ClosedXML (do obsługi plików Excel).
•	Uprawnienia: Dostęp do zapisu/odczytu w wybranym folderze z raportami.
________________________________________
🔐 Konfiguracja UPS API
Aby funkcja Auto-Green działała, należy w ustawieniach (ikona zębatki) wprowadzić dane dostępowe do UPS XML API:
1.	Access License Number
2.	User ID
3.	Password
Dane te są przechowywane lokalnie w pliku ups_settings.ini.
________________________________________
Uwaga: Program jest narzędziem wspomagającym. Zawsze należy zweryfikować krytyczne braki w systemach nadrzędnych.

