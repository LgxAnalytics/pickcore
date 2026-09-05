# PickCore

Warehouse picking cockpit. Jeden plik Pythona, GUI w Tkinter, zero zaleznosci
serwerowych. Narzedzie powstalo jako odpowiedz na konkretne waskie gardlo:
kompletacja zamowien weryfikowana wzrokowo z wydruku, bez kontroli bledow
i bez zadnych danych o tym, gdzie realnie schodzi czas.

**Status:** projekt portfolio. Kod jest generyczny, dane wejsciowe syntetyczne,
zadne nazwy klientow, dostawcow ani danych operacyjnych nie sa czescia repozytorium.

## Co robi

| Modul | Funkcja |
|---|---|
| Document pipeline | `pdfplumber` parsuje picking liste do struktury, render HTML, druk przez Edge headless |
| Scan validation | walidacja skanow SKU z tolerancja wariantow i odrzucaniem falszywych trafien |
| Speed-reject | analiza odstepow miedzyklawiszowych odroznia skaner od reki operatora |
| Serial capture | wychwytywanie numerow seryjnych w trakcie kompletacji, z undo |
| Put-away | rejestracja przyjec i relokacji binow |
| Handheld console | konsola HTTP dla terminala Wi-Fi, allowlista IP i token |
| Scanner listener | odbior skanow po TCP z terminali DataWedge |
| Pick heat map | mapa cieplna wizyt w regalach, widok 2D i izometryczny 3D |
| Star schema export | eksport do modelu gwiazdy pod analitykę slottingu |
| Label generator | etykiety ZPL na drukarki Zebra, generowane natywnie |

## Architektura

Jeden proces, watki wydzielone dla operacji blokujacych: watcher katalogu,
render PDF, listener skanera, serwer HTTP. Komunikacja z warstwa GUI wylacznie
przez kolejki, bo Tkinter nie jest thread safe.

Zapisy stanu ida zapisem atomowym (temp plus zamiana), po tym jak cicha awaria
zapisu raz kosztowala kartoteke. Instancja pilnowana mutexem, zeby dwa okna nie
pisaly po tym samym pliku.

## Uruchomienie

```
pip install pdfplumber pywin32
python pickcore.py
```

Build do exe:

```
python -m PyInstaller --onedir --windowed --name PickCore ^
    --collect-all pdfplumber --collect-all pdfminer pickcore.py
```

Onedir, nie onefile. Onefile rozpakowywal okolo 100 MB do `Temp\_MEIxxxxx`
przy kazdym starcie, a antywirus trzymal uchwyty przy zamykaniu, co konczylo
sie bledem usuwania katalogu tymczasowego.

## Bezpieczenstwo

Konsola HTTP stoi na LAN, bo terminale lacza sie po Wi-Fi. Dostep zamykaja
dwie bramki: allowlista IP z dopasowaniem prefiksu (wpis bez ostatniego
oktetu przepuszcza cala podsiec) oraz token przekazywany raz w URL i dalej
w cookie. Sekrety integracji nigdy nie trafiaja do configu ani do kodu,
tylko do zmiennej srodowiskowej albo Windows Credential Manager.

## Licencja

MIT
