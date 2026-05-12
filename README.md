Folyamat automatizálási script – MedWorks pácienskezelés
📘 Rövid leírás

Ez a script a MedWorks rendszerben végzi a páciensek automatikus keresését, felvételét és tételjelölését a felhasználói felületen keresztül.
A folyamat teljes egészében automatizált, és felhasználói beavatkozás nélkül futtatható.
A program az Excel-ben (Tételek, Páciensek munkalap) megadott adatok alapján dolgozik, és a megadott kémiai vizsgálati tételeket kijelöli a rendszerben.

Tételek felvitele exe:


✅ Config.json külső fájlból való betöltés (rugalmas deployment)
✅ Automatikus bejelentkezés
✅ Mai dátumú páciensek szűrése az Excel-ből
✅ Automatikus tétel-bepipálás (D oszlop kód alapján)
✅ Mintavétel és mintaszállítás dátumok automatikus beállítása
✅ Várás a felhasználóra - ellenőrzés és mentés
✅ Automatikus mentés észlelés (nem kell ENTER!)
✅ Következő páciens feldolgozása automatikusan
✅ Kilépés amikor minden páciens kész
flowchart TD

    A[Start - script indul] --> B[.env beolvasása]
    B --> C[WebDriver indítása]
    C --> D[Navigálás: login oldal]
    D --> E[Login gomb kattintása]
    E --> F{Login sikeres?}

    F -- Nem --> Z1[Hiba: login sikertelen] --> END
    F -- Igen --> G[Dashboard betöltve<br/>Log: Login sikeres]

    G --> H[Excel betöltése<br/>adatok.xlsm / 'Tételek']
    H --> I[Fejléc nélküli formátum felismerése]
    I --> J[Páciensek számának logolása]
    J --> K[For ciklus a pácienseken]

    K --> L[#i páciens neve logolva]
    L --> M[open_and_search_patient(driver, name)]

    M --> M1{Sikeres páciens felvétel?}
    M1 -- Nem --> E1[Hiba logolása<br/>Páciens felvétel sikertelen] --> N
    M1 -- Igen --> O[Páciens sikeresen felvéve]

    O --> P[Tételek feldolgozása (loop)]
    P --> P1[Minden tétel bejelölve]
    P1 --> P2[Dátum beállítása JS-sel]

    P2 --> P3[Sticker dialog kezelése<br/>Bezárás vagy skip]

    %% 🔥 ITT az új rész — egyetlen gomb
    P3 --> S[💾📤 'Mentés és feladás' gomb megnyomása]

    S --> S1{Sikeres művelet?}
    S1 -- Nem --> S2[Hiba logolása<br/>Retry / manuális beavatkozás] --> U
    S1 -- Igen --> U[Log: páciens sikeresen mentve és feladva]

    U --> V{Van még páciens?}
    V -- Igen --> K
    V -- Nem --> W[🏁 Tételek futtatása véget ért]

    W --> END([End])

## Páciensregisztráció – regisztracio.py

Ez a script a MedWorks rendszerben végzi az új páciensek 
automatikus regisztrációját. Az Excel fájlban (Páciensek munkalap) 
megadott adatok alapján automatikusan kitölti a regisztrációs 
űrlapot a webes felületen, majd megvárja a felhasználó jóváhagyását.

### Főbb funkciók:
✅ Config.json külső fájlból való betöltés  
✅ Automatikus bejelentkezés  
✅ Páciensadatok beolvasása Excel fájlból  
✅ Automatikus adatbevitel a webes felületen  
✅ Várás a felhasználóra – ellenőrzés és manuális mentés  