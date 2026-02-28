# OCRProject

OCR = Optical Character Recognition, also Text aus Bildern lesen
Tesseract = https://github.com/charlesw/tesseract
.NET = Framework
ClosedXML = erlaubt Excel zu lesen/bearbeiten, ohne dass Excel installiert werden muss
P/Invoke — Windows-Systemfunktionen direkt aus C# aufrufen

## Wie der Code funktioniert — Schritt für Schritt
### Programm startet
    → TesseractEngine wird initialisiert
    → Hotkey Ctrl+Shift+S wird beim Windows-System registriert
    → Fenster wartet im Hintergrund

### User drückt Strg+Shift+S
    → Windows schickt eine WM_HOTKEY Nachricht ans Fenster
    → WndProc() fängt diese Nachricht ab
    → CaptureOcrAndWriteToExcel() wird aufgerufen
        → Screenshot der definierten Region wird gemacht
        → Screenshot wird als PNG in den RAM gespeichert
        → Tesseract liest den Text aus dem Bild
        → ClosedXML schreibt Datum + Text in Excel
    → Piepton als Bestätigung

## Befehle
```
dotnet build
dotnet run --project OcrToExcel.Win
```
