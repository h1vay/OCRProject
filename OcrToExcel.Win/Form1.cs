using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tesseract;
using Excel = Microsoft.Office.Interop.Excel;

namespace OcrToExcel.Win;

public partial class Form1 : Form
{
    // =========================
    // KONFIGURATION (HIER ANPASSEN)
    // =========================

    // Voller Pfad zur Excel-Datei, die in Excel GEÖFFNET sein muss
    private const string ExcelFullPath = @"C:\Temp\werte.xlsx";

    // Blattname, in das geschrieben werden soll (muss existieren)
    private const string SheetName = "Daten";

    // Screenshot-Region (Pixel): x, y, width, height
    // Beispiel "links oben": x=0,y=0
    private static readonly Rectangle CaptureRegion = new Rectangle(x: 0, y: 0, width: 400, height: 120);

    // Hotkey: Ctrl + Shift + S
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const int HOTKEY_ID = 1;

    // OCR Sprache
    private const string OcrLanguage = "deu";

    // =========================

    private TesseractEngine? _engine;

    public Form1()
    {
        InitializeComponent();

        Text = "OCR -> Excel (Hotkey: Ctrl+Shift+S)";
        StartPosition = FormStartPosition.CenterScreen;

        // OCR Engine initialisieren
        var tessDataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        if (!Directory.Exists(tessDataPath))
        {
            MessageBox.Show(
                $"Ordner fehlt: {tessDataPath}\nLege 'tessdata' neben die EXE und füge '{OcrLanguage}.traineddata' hinzu.",
                "Tesseract fehlt",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        _engine = new TesseractEngine(tessDataPath, OcrLanguage, EngineMode.Default);

        // Optional: wenn du vor allem Zahlen erwartest, Whitelist aktivieren:
        // _engine.SetVariable("tessedit_char_whitelist", "0123456789.,:-");

        // Hotkey registrieren
        RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (uint)Keys.S);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        try
        {
            UnregisterHotKey(Handle, HOTKEY_ID);
        }
        catch { /* ignore */ }

        _engine?.Dispose();
        _engine = null;

        base.OnFormClosed(e);
    }

    // Hotkey abfangen
    protected override void WndProc(ref Message m)
    {
        const int WM_HOTKEY = 0x0312;

        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
        {
            try
            {
                CaptureOcrAndWriteToExcel();
                System.Media.SystemSounds.Asterisk.Play();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        base.WndProc(ref m);
    }

    private void CaptureOcrAndWriteToExcel()
    {
        if (_engine == null)
            throw new InvalidOperationException("OCR Engine ist nicht initialisiert.");

        using var bmp = CaptureRegionBitmap(CaptureRegion);

        // Optional: Debug-Screenshot speichern
        // bmp.Save(Path.Combine(AppContext.BaseDirectory, "debug.png"), ImageFormat.Png);

        var text = OcrBitmap(bmp, _engine);

        // Wenn nichts erkannt wurde, optional nichts schreiben
        if (string.IsNullOrWhiteSpace(text))
            return;

        AppendToOpenExcel(
            excelFullPath: ExcelFullPath,
            sheetName: SheetName,
            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            text
        );
    }

    private static Bitmap CaptureRegionBitmap(Rectangle region)
    {
        var bmp = new Bitmap(region.Width, region.Height, PixelFormat.Format24bppRgb);
        using var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(region.Left, region.Top, 0, 0, region.Size, CopyPixelOperation.SourceCopy);
        return bmp;
    }

    private static string OcrBitmap(Bitmap bmp, TesseractEngine engine)
    {
        using var pix = PixConverter.ToPix(bmp);
        using var page = engine.Process(pix);
        return (page.GetText() ?? string.Empty).Trim();
    }

    private static void AppendToOpenExcel(string excelFullPath, string sheetName, params object[] values)
    {
        // Excel muss laufen
        var app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

        // Workbook anhand FullName finden
        Excel.Workbook? wb = null;
        foreach (Excel.Workbook w in app.Workbooks)
        {
            if (string.Equals(w.FullName, excelFullPath, StringComparison.OrdinalIgnoreCase))
            {
                wb = w;
                break;
            }
        }

        if (wb == null)
            throw new InvalidOperationException("Die Excel-Datei ist nicht geöffnet:\n" + excelFullPath);

        Excel.Worksheet? ws = null;
        try
        {
            ws = (Excel.Worksheet)wb.Worksheets[sheetName];
        }
        catch
        {
            throw new InvalidOperationException($"Sheet '{sheetName}' existiert nicht in der Arbeitsmappe.");
        }

        // Nächste freie Zeile in Spalte A
        int lastRow = ws.Cells[ws.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
        int nextRow = (lastRow == 1 && ws.Cells[1, 1].Value2 == null) ? 1 : lastRow + 1;

        for (int i = 0; i < values.Length; i++)
            ws.Cells[nextRow, i + 1].Value2 = values[i];

        wb.Save();
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}