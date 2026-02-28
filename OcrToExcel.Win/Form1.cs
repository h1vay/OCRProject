using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tesseract;
using ClosedXML.Excel;

namespace OcrToExcel.Win;

public partial class Form1 : Form
{
    private const string ExcelFullPath = @"C:\Users\Can\Desktop\OCRProject\ExcelTabelleOCR.xlsx";
    private const string SheetName = "Daten";
    private static readonly Rectangle CaptureRegion = new Rectangle(x: 0, y: 0, width: 400, height: 120);

    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const int HOTKEY_ID = 1;
    private const string OcrLanguage = "deu";

    private TesseractEngine? _engine;

    public Form1()
    {
        InitializeComponent();
        Text = "OCR -> Excel (Hotkey: Ctrl+Shift+S)";
        StartPosition = FormStartPosition.CenterScreen;

        var tessDataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        if (!Directory.Exists(tessDataPath))
        {
            MessageBox.Show(
                $"Ordner fehlt: {tessDataPath}",
                "Tesseract fehlt",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        _engine = new TesseractEngine(tessDataPath, OcrLanguage, EngineMode.Default);
        RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (uint)Keys.S);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        try { UnregisterHotKey(Handle, HOTKEY_ID); } catch { }
        _engine?.Dispose();
        _engine = null;
        base.OnFormClosed(e);
    }

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
        var text = OcrBitmap(bmp, _engine);

        if (string.IsNullOrWhiteSpace(text))
            return;

        AppendToOpenExcel(ExcelFullPath, SheetName,
            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            text);
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