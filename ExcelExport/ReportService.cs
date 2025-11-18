using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;

public class ReportService
{
    public async Task<ReportFileViewModel> ExportStudentAdmissionReport(ReportDSK dsk, string userId)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Hesabat");

        // ==========================================
        // 1. ÜMUMİ STİLLƏR VƏ ŞRİFTLƏR
        // ==========================================
        var allCells = ws.Style;
        allCells.Font.FontName = "Calibri";
        allCells.Font.FontSize = 9;
        allCells.Alignment.WrapText = true;
        allCells.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        allCells.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

        int totalCols = 33; // Şəkilə əsasən 33 sütun

        // ==========================================
        // 2. BAŞLIQ SƏTİRLƏRİ (HEADER) - Şəkilə uyğun
        // ==========================================

        // --- Sətir 1: Əsas Başlıq ---
        ws.Range(1, 1, 1, totalCols).Merge().Value = "1 oktyabr 2024-cü il vəziyyətinə tələbələrin və məzunların ixtisaslar üzrə sayı";
        ws.Range(1, 1, 1, totalCols).Style.Font.SetBold().Font.SetFontSize(11)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

        // --- Sətir 2, 3, 4: Cədvəl Başlıqları ---

        // Sol tərəfdəki sabit sütunlar (Sətir 2-dən 4-ə qədər birləşir)
        ws.Range(2, 1, 4, 1).Merge().Value = "Tabeçilik";
        ws.Range(2, 2, 4, 2).Merge().Value = "Təhsil müəssisəsinin adı";
        ws.Range(2, 3, 4, 3).Merge().Value = "Təhsil forması";
        ws.Range(2, 4, 4, 4).Merge().Value = "İxtisasların adı";
        ws.Range(2, 5, 4, 5).Merge().Value = "İxtisasın kodu";
        ws.Range(2, 6, 4, 6).Merge().Value = "Qəbul planı";

        // "Qəbul olunub" bloku
        ws.Range(2, 7, 2, 12).Merge().Value = "Qəbul olunub 1)";

        ws.Range(3, 7, 4, 7).Merge().Value = "DİM xətti ilə";
        ws.Range(3, 8, 4, 8).Merge().Value = "onlardan ödənişli əsaslarla (süt. 2)";
        ws.Range(3, 9, 4, 9).Merge().Value = "DİM xətti ilə qəbul olanlardan qadınlar (süt. 2)";
        ws.Range(3, 10, 4, 10).Merge().Value = "onlardan ödənişli əsaslarla (süt. 4)";
        ws.Range(3, 11, 4, 11).Merge().Value = "Təkrar təhsil";
        ws.Range(3, 12, 4, 12).Merge().Value = "onlardan qadınlar";

        // "Kurslar üzrə təhsil alanlar" bloku
        ws.Range(2, 13, 2, 24).Merge().Value = "kurslar üzrə təhsil alanlar";
        string[] kurslar = { "I", "II", "III", "IV", "V", "VI" };
        for (int i = 0; i < 6; i++)
        {
            int col = 13 + (i * 2);
            ws.Range(3, col, 3, col + 1).Merge().Value = kurslar[i];
            // Sol tərəf (Cəmi) başlıqsız qalır və ya boş olur, sağ tərəf "qadınlar"
            ws.Cell(4, col + 1).Value = "onlardan qadınlar";
            ws.Cell(4, col + 1).Style.Alignment.TextRotation = 90; // Yer qənaəti üçün şaquli yazı
        }

        // Yekun Statistikalar (Sarı sütunlar daxil)
        ws.Range(2, 25, 4, 25).Merge().Value = "Bütün kurslarda təhsil alanlar, (süt. 8,10,12,14, 16,18 cəmi)";
        ws.Range(2, 26, 4, 26).Merge().Value = "onlardan ödənişli əsaslarla təhsil alanlar (süt. 20)";
        ws.Range(2, 27, 4, 27).Merge().Value = "Cəmi təhsil alanlardan qadınlar (süt. 20)";
        ws.Range(2, 28, 4, 28).Merge().Value = "onlardan ödənişli əsaslarla (süt. 22)";

        // Buraxılış Bloku
        ws.Range(2, 29, 2, 32).Merge().Value = "01.10.2023-cü ildən 01.10.2024-cü ilədək faktiki buraxılış";
        ws.Range(3, 29, 4, 29).Merge().Value = "Yekun dövlət attestasiyasına buraxılanlar";
        ws.Range(3, 30, 4, 30).Merge().Value = "onlardan qadınlar";
        ws.Range(3, 31, 4, 31).Merge().Value = "Bakalavr diplomu alanlar";
        ws.Range(3, 32, 4, 32).Merge().Value = "onlardan qadınlar";

        // Gözlənilən Buraxılış
        ws.Range(2, 33, 4, 33).Merge().Value = "01.10.2024-cü ildən 01.10.2025-ci ilədək gözlənilən buraxılış";


        // ==========================================
        // 3. MƏLUMATLARIN DOLDURULMASI (Şəkildəki eyni data)
        // ==========================================

        var rowsData = new List<object[]>
        {
            // R1: Beynəlxalq münasibətlər
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Beynəlxalq münasibətlər", "050201",
                           80, 79, 79, 48, 48, null, null,
                           77, 46, 94, 56, 67, 39, 66, 39, 32, 9, 7, 2, // Kurslar I-VI
                           343, 343, 191, 191, // Cəmi
                           56, 39, 56, 39, 81 }, // Buraxılış

            // R2: Dövlət və ictimai münasibətlər
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Dövlət və ictimai münasibətlər", "050203",
                           80, 80, 80, 55, 55, null, null,
                           80, 55, 101, 75, 94, 72, 70, 51, 17, 13, 2, 1,
                           364, 364, 267, 267,
                           86, 62, 86, 62, 62 },

            // R3: Hüquqşünaslıq
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Hüquqşünaslıq", "050206",
                           100, 100, 100, 60, 60, null, null,
                           99, 59, 148, 95, 84, 50, 97, 68, 18, 9, 5, 1,
                           451, 451, 282, 282,
                           74, 49, 74, 49, 88 },

            // R4: Kommunikasiya və rəqəmsal media
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Kommunikasiya və rəqəmsal media", "050216",
                           40, 40, 40, 28, 28, null, null,
                           40, 30, 54, 41, 43, 34, 15, 12, null, null, null, null,
                           152, 152, 117, 117,
                           null, null, null, null, 10 }
        };

        int currentRow = 5;
        foreach (var rowData in rowsData)
        {
            for (int i = 0; i < rowData.Length; i++)
            {
                if (rowData[i] != null)
                {
                    ws.Cell(currentRow, i + 1).Value = rowData[i].ToString(); // ClosedXML bəzən object tipini birbaşa sevmir

                    // Rəqəmdirsə double kimi set edək ki, Excel xəta verməsin
                    if (double.TryParse(rowData[i].ToString(), out double num))
                    {
                        ws.Cell(currentRow, i + 1).Value = num;
                    }
                }
            }
            // Şəkildəki sarı rəngli xanalar (Sütun 25, 26, 27, 28)
            ws.Cell(currentRow, 25).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204); // Açıq sarı
            ws.Cell(currentRow, 26).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204);
            ws.Cell(currentRow, 27).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204);
            ws.Cell(currentRow, 28).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204);

            currentRow++;
        }


        // ==========================================
        // 4. FORMATLAŞDIRMA VƏ SƏRHƏDLƏR
        // ==========================================

        // Sütun genişlikləri (Şəkilə uyğun tənzimləmə)
        ws.Column(1).Width = 5;   // Tabeçilik
        ws.Column(2).Width = 25;  // Uni adı
        ws.Column(3).Width = 8;   // Forma
        ws.Column(4).Width = 35;  // İxtisas adı
        ws.Column(5).Width = 10;  // Kod
        ws.Column(6).Width = 6;   // Plan

        // Rəqəm sütunlarını daraldırıq
        for (int c = 7; c <= 33; c++) ws.Column(c).Width = 5;

        // Başlıqları olan geniş sütunlar
        ws.Column(25).Width = 8;
        ws.Column(26).Width = 8;
        ws.Column(27).Width = 8;
        ws.Column(28).Width = 8;

        // Sərhədlər (Bütün cədvəl üçün)
        var tableRange = ws.Range(2, 1, currentRow - 1, totalCols);
        tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Başlıqların arxa plan rəngi (Açıq boz - şəkilə oxşar)
        ws.Range(2, 1, 4, totalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);

        // Ad və İxtisas sütunlarını sola, digərlərini ortaya düzləndir
        ws.Range(5, 2, currentRow - 1, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
        ws.Range(5, 4, currentRow - 1, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);


        // ==========================================
        // 5. ÇIXIŞ (EXPORT)
        // ==========================================
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var bytes = ms.ToArray();

        return ReportFileViewModel.FileSuccess(bytes, "Tələbə və Məzun Hesabatı.xlsx");
    }
}

// --- Köməkçi Modellər ---
public class ReportDSK { }

public class ReportFileViewModel
{
    public byte[]? FileBytes { get; set; }
    public string? FileName { get; set; }
    public bool IsSuccess { get; set; }
    public static ReportFileViewModel FileSuccess(byte[] bytes, string fileName)
    {
        return new ReportFileViewModel { FileBytes = bytes, FileName = fileName, IsSuccess = true };
    }
}