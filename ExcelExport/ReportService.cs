using ClosedXML.Excel;
using System.Collections.Generic;
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

        // Şəkildəki sütunların sayı: Tabeçilik(1) ... Qəbul planı(6) ... Kurslar(24) ... Yekun(28)
        int totalCols = 28;

        // ==========================================
        // 2. BAŞLIQ SƏTİRLƏRİ (HEADER) - Şəkilə uyğun
        // ==========================================

        // --- Sətir 1: Əsas Başlıq ---
        ws.Range(1, 1, 1, totalCols).Merge().Value = "Ödənişli əsaslarla təhsil alan tələbələrdən təhsil xərcləri dövlət büdcəsindən ödənilənlərin (imtiyazlılar) sayı";
        ws.Range(1, 1, 1, totalCols).Style.Font.SetBold().Font.SetFontSize(11)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

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
            // Sol boş (cəmi), sağ "qadınlar"
            ws.Cell(4, col).Value = "";
            ws.Cell(4, col + 1).Value = "onlardan qadınlar";
            ws.Cell(4, col + 1).Style.Alignment.TextRotation = 90; // Şaquli yazı
        }

        // Yekun Statistikalar (Sarı sütunlar)
        ws.Range(2, 25, 4, 25).Merge().Value = "Bütün kurslarda təhsil alanlar, (süt. 8,10,12,14, 16,18 cəmi)";
        ws.Range(2, 26, 4, 26).Merge().Value = "onlardan ödənişli əsaslarla təhsil alanlar (süt. 20)";
        ws.Range(2, 27, 4, 27).Merge().Value = "Cəmi təhsil alanlardan qadınlar (süt. 20)";
        ws.Range(2, 28, 4, 28).Merge().Value = "onlardan ödənişli əsaslarla (süt. 22)";


        // --- Sətir 5: Sütun Nömrələnməsi (A, B, 1, 2...) ---
        ws.Cell(5, 4).Value = "A";
        ws.Cell(5, 5).Value = "B";
        for (int i = 1; i <= 23; i++)
        {
            ws.Cell(5, 5 + i).Value = i.ToString(); // Sütun 6-dan başlayaraq 1 yazır
        }
        ws.Range(5, 1, 5, totalCols).Style.Font.SetBold();
        ws.Range(5, 1, 5, totalCols).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


        // ==========================================
        // 3. MƏLUMATLARIN DOLDURULMASI (Şəkildəki dəqiq sətirlər)
        // ==========================================

        var rowsData = new List<object[]>
        {
            // Sətir 01: İxtisaslar üzrə yekun
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "İxtisaslar üzrə yekun", "01",
                           980, 963, 963, 491, 491, null, null,
                           955, 485, 982, 498, 661, 362, 532, 304, 202, 81, 57, 16,
                           3389, 3389, 1746, 1746 },

            // Sətir 02: Yekun saydan ödənişli
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Yekun saydan (sətir 01-dən):\nödənişli əsaslarla təhsil alanlar", "02",
                           980, 963, 963, 491, 491, null, null,
                           955, 485, 982, 498, 661, 362, 532, 304, 202, 81, 57, 16,
                           3389, 3389, 1746, 1746 },

            // Sətir 03: Dövlət büdcəsindən ödənilənlər
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "onlardan (sətir 02-dən)\ntəhsil xərcləri dövlət büdcəsindən\nödənilənlərin sayı", "03",
                           null, 141, 141, 141, 64, 64, 0, 0,
                           141, 116, 283, 135, 142, 72, 140, 79, 28, 12, 3, 1,
                           737, null, 415, null },

            // Sətir 04: Valideyn himayəsindən məhrum
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "valideynlərini itirmiş və ya\nvalideyn himayəsindən məhrum\nolmuş uşaqlar", "04",
                           5, 5, 5, 3, 3, 0, 0,
                           5, 5, 7, 4, 2, 2, 5, 3, 2, 1, 0, 0,
                           21, 21, 15, 13 },

            // Sətir 05: Şəhid ailəsinin üzvü
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "şəhid ailəsinin üzvü statusu olan\nşəxslər (həyat yoldaşı və s.)", "05",
                           0, 0, 0, 0, 0, 0, 0,
                           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                           null, null, null, null },

            // Sətir 06: Şəhid övladları
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "şəhid övladları", "06",
                           2, 2, 2, 0, 0, 0, 0,
                           2, 1, 4, 2, 2, 1, 0, 0, 0, 0, 0, 0,
                           8, 8, 4, 3 },

            // Sətir 07: Müharibə əlillərinin övladları
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "müharibə əlillərinin övladları", "07",
                           17, 17, 17, 11, 11, 0, 0,
                           17, 14, 15, 6, 10, 5, 12, 7, 5, 3, 0, 0,
                           59, 58, 35, 32 },

            // Sətir 08: Müharibə veteranları
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "müharibə veteranları", "08",
                           0, 0, 0, 0, 0, 0, 0,
                           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                           null, null, null, null },

            // Sətir 09: Əlilliyi müəyyən edilmiş şəxslər
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "əlilliyi müəyyən edilmiş şəxslər", "09",
                           5, 5, 5, 3, 3, 0, 0,
                           5, 3, 3, 0, 1, 1, 0, 0, 0, 0, 0, 0,
                           9, 9, 4, 4 },

            // Sətir 10: Məcburi köçkünlər
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "məcburi köçkün statusu olan\nşəxslər", "10",
                           112, 112, 112, 47, 47, 0, 0,
                           112, 93, 254, 123, 127, 63, 123, 69, 21, 8, 3, 1,
                           640, 630, 357, 308 }
        };

        int currentRow = 6;
        foreach (var rowData in rowsData)
        {
            for (int i = 0; i < rowData.Length; i++)
            {
                if (rowData[i] != null)
                {
                    if (double.TryParse(rowData[i].ToString(), out double num))
                    {
                        ws.Cell(currentRow, i + 1).Value = num;
                    }
                    else
                    {
                        ws.Cell(currentRow, i + 1).Value = rowData[i].ToString();
                    }
                }
            }

            // Sarı rəngli xanalar (Son 4 sütun: 25, 26, 27, 28)
            ws.Range(currentRow, 25, currentRow, 28).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204); // Açıq sarı

            currentRow++;
        }


        // ==========================================
        // 4. FORMATLAŞDIRMA
        // ==========================================

        // Sütun genişlikləri
        ws.Column(1).Width = 5;   // Tabeçilik
        ws.Column(2).Width = 25;  // Uni adı
        ws.Column(3).Width = 8;   // Forma
        ws.Column(4).Width = 35;  // İxtisas adı
        ws.Column(5).Width = 8;   // Kod

        // Rəqəm sütunlarını daraldırıq (Sütun 6-dan sona qədər)
        for (int c = 6; c <= totalCols; c++) ws.Column(c).Width = 5;

        // Sarı sütunlar biraz geniş olsun başlıqlara görə
        ws.Column(25).Width = 8;
        ws.Column(26).Width = 8;
        ws.Column(27).Width = 8;
        ws.Column(28).Width = 8;

        // Sərhədlər (Border)
        var tableRange = ws.Range(2, 1, currentRow - 1, totalCols);
        tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Başlıqların arxa plan rəngi (Boz)
        ws.Range(2, 1, 5, totalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);

        // Ad və İxtisas sütunlarını sola düzləndir, digərləri mərkəz
        ws.Range(6, 2, currentRow - 1, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
        ws.Range(6, 4, currentRow - 1, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);


        // ==========================================
        // 5. ÇIXIŞ (EXPORT)
        // ==========================================
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var bytes = ms.ToArray();

        return ReportFileViewModel.FileSuccess(bytes, "Imtiyazlilar_Hesabati.xlsx");
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