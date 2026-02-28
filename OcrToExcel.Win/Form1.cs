using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tesseract;
using ClosedXML.Excel;

namespace OcrToExcel.Win;

public partial class Form1 : Form {
    private const string ExcelFullPath = @"C:\Users\Can\Desktop\OCRProject\ExcelTabelleOCR.xlsx";
    private const string SheetName = "Daten";

    // Wert für Screenshot eingeben
    private static readonly Rectangle CaptureRegion = new Rectangle(x: 0, y: 0, width: 400, height: 120);

    // Tasten für Screenshot
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const int HOTKEY_ID = 1;
    private const string OcrLanguage = "deu";

    //UI
    private readonly PictureBox _preview;

    // Zeigt den erkannten Text und Status an
    private readonly Label _statusLabel;

    private TesseractEngine? _engine;

    public Form1()
    {

        InitializeComponent();

        // Windows-Fenster einrichten
        Text            = "Programm gestartet: Drücke Strg+Shift+S gleichzeitig, um einen Screenshot zu erstellen.";
        StartPosition   = FormStartPosition.CenterScreen;
        Size            = new Size(500, 320);
        FormBorderStyle = FormBorderStyle.FixedSingle; // Größe nicht veränderbar

        // Vorschau-Bild erstellen
        _preview = new PictureBox
        {
            Location  = new Point(10, 10),
            Size      = new Size(460, 180),
            BorderStyle = BorderStyle.FixedSingle,
            SizeMode  = PictureBoxSizeMode.Zoom,  // Bild wird proportional skaliert
            BackColor = Color.WhiteSmoke
        };
        Controls.Add(_preview);
        _preview.Paint += (s, e) =>
{
    if (_preview.Image == null)
    {
        var text = "Noch kein Screenshot gemacht";
        var font = new Font("Segoe UI", 10);
        var size = e.Graphics.MeasureString(text, font);
        e.Graphics.DrawString(text, font, Brushes.Gray,
            (_preview.Width - size.Width) / 2,
            (_preview.Height - size.Height) / 2);
    }
};

        // Statustext erstellen
        _statusLabel = new Label
        {
            Location  = new Point(10, 200),
            Size      = new Size(460, 80),
            Text      = "Bereit. Drücke Strg+Shift+S um einen Screenshot zu machen.",
            Font      = new Font("Segoe UI", 10),
            ForeColor = Color.DimGray
        };
        Controls.Add(_statusLabel);

        // Tesseract (OCR) initialisieren
        var tessDataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        if (!Directory.Exists(tessDataPath))
        {
            MessageBox.Show(
                $"Ordner fehlt: {tessDataPath}\n'tessdata' mit '{OcrLanguage}.traineddata' neben die EXE legen.",
                "Tesseract fehlt", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        _engine = new TesseractEngine(tessDataPath, OcrLanguage, EngineMode.Default);
        RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (uint)Keys.S);

        /* Falls TIMER gewünscht
        var timer = new System.Windows.Forms.Timer();
        timer.Interval = 5000; // 5000ms = 5 Sekunden
        timer.Tick += (s, e) => 
        {
            try { CaptureOcrAndWriteToExcel(); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        };
        timer.Start();
            */
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        try { UnregisterHotKey(Handle, HOTKEY_ID); } catch { }
        _engine?.Dispose();
        _engine = null;
        base.OnFormClosed(e);
    }

    // Fehler fangen
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

    //HAUPTPROGRAMM
    private void CaptureOcrAndWriteToExcel() {
    if (_engine == null)
        throw new InvalidOperationException("OCR Engine ist nicht initialisiert.");

    var bmp = CaptureRegionBitmap(CaptureRegion);
    _preview.Image?.Dispose();
    _preview.Image = bmp;

    var text = OcrBitmap(bmp, _engine);

    AppendToOpenExcel(ExcelFullPath, SheetName,
        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
        text);
    _statusLabel.ForeColor = Color.Green;
    _statusLabel.Text = $"✓ Gespeichert: {text}";
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
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        var bytes = ms.ToArray();

        using var pix = Pix.LoadFromMemory(bytes);
        using var page = engine.Process(pix);
        return (page.GetText() ?? string.Empty).Trim();
    }

    private static void AppendToOpenExcel(string excelFullPath, string sheetName, params object[] values)
    {
        using var wb = new XLWorkbook(excelFullPath);

        if (!wb.TryGetWorksheet(sheetName, out var ws))
            throw new InvalidOperationException($"Sheet '{sheetName}' existiert nicht.");

        int nextRow = ws.LastRowUsed()?.RowNumber() + 1 ?? 1;

        for (int i = 0; i < values.Length; i++)
            ws.Cell(nextRow, i + 1).Value = values[i]?.ToString();

        wb.Save();
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }