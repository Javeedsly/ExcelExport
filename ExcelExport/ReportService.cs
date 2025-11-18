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
        // 1. ÜMUMİ STİLLƏR
        // ==========================================
        var allCells = ws.Style;
        allCells.Font.FontName = "Calibri";
        allCells.Font.FontSize = 9;
        allCells.Alignment.WrapText = true;
        allCells.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        allCells.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

        int totalCols = 17; // Cədvəldə 17 sütun var

        // ==========================================
        // 2. BAŞLIQ SƏTİRLƏRİ (HEADER)
        // ==========================================

        // --- Sətir 1: Əsas Başlıq ---
        ws.Range(1, 1, 1, totalCols).Merge().Value = "Ali təhsil müəssisələrində çalışan işçilərin sayı barədə məlumat";
        ws.Range(1, 1, 1, totalCols).Style.Font.SetBold().Font.SetFontSize(11);

        // --- Sətir 2, 3, 4: Mürəkkəb Başlıqlar ---

        // Sol Sabit Sütunlar (Sətir 2-dən 4-ə qədər birləşir)
        ws.Range(2, 1, 4, 1).Merge().Value = "Tabeçilik";
        ws.Range(2, 2, 4, 2).Merge().Value = "Təhsil müəssisəsinin adı";
        ws.Range(2, 3, 4, 3).Merge().Value = "Göstəricilər";
        ws.Range(2, 4, 4, 4).Merge().Value = "Sətrin №-si";

        // Blok: Əsas heyət (ştatda olanlar)
        ws.Range(2, 5, 3, 6).Merge().Value = "Əsas heyət\n(ştatda olanlar)";
        ws.Cell(4, 5).Value = "cəmi";
        ws.Cell(4, 6).Value = "qadın";

        // Blok: Əsas heyətdən (tarif)
        ws.Range(2, 7, 3, 8).Merge().Value = "Əsas heyətdən";
        ws.Cell(4, 7).Value = "tam tarif üzrə işləyənlər";
        ws.Cell(4, 8).Value = "0,5 və 0,25 tarif ilə işləyənlər";

        // Blok: Rəhbər vəzifədə olanlardan dərs deyənlər
        ws.Range(2, 9, 3, 10).Merge().Value = "Rəhbər vəzifədə\nolanlardan dərs\ndeyənlər";
        ws.Cell(4, 9).Value = "cəmi";
        ws.Cell(4, 10).Value = "qadın";

        // Blok: Kənardan cəlb olunan
        ws.Range(2, 11, 3, 12).Merge().Value = "Kənardan cəlb\nolunan əvəzedici\nheyət";
        ws.Cell(4, 11).Value = "cəmi";
        ws.Cell(4, 12).Value = "qadın";

        // Blok: Əsas heyətdən (Elmi dərəcə/ad) - Böyük blok
        ws.Range(2, 13, 2, 16).Merge().Value = "Əsas heyətdən";

        // Alt bloklar
        ws.Range(3, 13, 3, 14).Merge().Value = "elmi dərəcəsi olanlar";
        ws.Cell(4, 13).Value = "elmlər doktoru";
        ws.Cell(4, 14).Value = "fəlsəfə doktoru";

        ws.Range(3, 15, 3, 16).Merge().Value = "elmi adı olanlar";
        ws.Cell(4, 15).Value = "professor";
        ws.Cell(4, 16).Value = "dosent";

        // Sonuncu sütun
        ws.Range(2, 17, 4, 17).Merge().Value = "Bundan başqa\nəcnəbi və\nvətəndaşlığı\nolmayan\nmütəxəssislər";

        // Başlıqların stili
        ws.Range(2, 1, 4, totalCols).Style.Font.SetBold();
        ws.Range(2, 1, 4, totalCols).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


        // ==========================================
        // 3. MƏLUMATLARIN DOLDURULMASI
        // ==========================================

        var rowsData = new List<object[]>
        {
            // R01 (Yaşıl)
            new object[] { "DDQ", "ADA Universiteti", "Cəmi işçilərin sayı\n(02, 09, 14-cü sətirlərin cəmi)", "01",
                           443, 221, 439, 4, 10, 2, 111, 46, 1, 81, 1, null, 23 },
            
            // R02 (Yaşıl)
            new object[] { "DDQ", "ADA Universiteti", "onlardan:\nRəhbər vəzifədə olanlar (03-08-ci sətirlərin cəmi)", "02",
                           22, 8, 22, null, 10, 2, null, null, 1, 13, 1, null, 3 },

            // R03
            new object[] { "DDQ", "ADA Universiteti", "o cümlədən:\nrektor", "03",
                           1, null, 1, null, null, null, null, null, 1, null, 1, null, null },
            
            // R04
            new object[] { "DDQ", "ADA Universiteti", "prorektor və filial direktorları", "04",
                           4, 2, 4, null, 2, 1, null, null, null, 3, null, null, null },

            // R05
            new object[] { "DDQ", "ADA Universiteti", "laboratoriya və şöbə müdirləri", "05",
                           11, 4, 11, null, 3, null, null, null, null, 5, null, null, 2 },

            // R06
            new object[] { "DDQ", "ADA Universiteti", "dekanlar", "06",
                           6, 2, 6, null, 5, 1, null, null, null, 5, null, null, 1 },

            // R07
            new object[] { "DDQ", "ADA Universiteti", "dekan müavinləri", "07",
                           null, null, null, null, null, null, null, null, null, null, null, null, null },

            // R08
            new object[] { "DDQ", "ADA Universiteti", "kafedra müdirləri", "08",
                           null, null, null, null, null, null, null, null, null, null, null, null, null },

            // R09 (Yaşıl)
            new object[] { "DDQ", "ADA Universiteti", "Professor-müəllim heyət (10-13-cü sətirlərin cəmi)\n( rəhbər vəzifədə olanlar istisna olmaqla)", "09",
                           162, 71, 162, null, null, null, 111, 46, null, 68, null, null, 20 },

            // R10
            new object[] { "DDQ", "ADA Universiteti", "o cümlədən:\nprofessorlar", "10",
                           10, 1, 10, null, "x", "x", null, null, null, 6, null, null, 6 },

            // R11
            new object[] { "DDQ", "ADA Universiteti", "dosentlər", "11",
                           11, 1, 11, null, "x", "x", null, null, null, 11, null, null, 4 },

            // R12
            new object[] { "DDQ", "ADA Universiteti", "baş müəllimlər", "12",
                           48, 12, 48, null, "x", "x", null, null, null, 48, null, null, 8 },

            // R13
            new object[] { "DDQ", "ADA Universiteti", "müəllimlər, assistentlər", "13",
                           93, 57, 93, null, "x", "x", 111, 46, null, 3, null, null, 2 },

            // R14
            new object[] { "DDQ", "ADA Universiteti", "Sair heyət", "14",
                           259, 142, 255, 4, null, null, null, null, null, null, null, null, null },

            // R15 (Yaşıl) - Ayrılmış sətir
            new object[] { "DDQ", "ADA Universiteti", "Professor-müəllim heyət\n( rəhbər vəzifədə olanlardan dərs deyənlər\ndə daxil olmaqla )", "15",
                           172, 73, 162, 10, 10, 2, 111, 46, 1, 81, 1, null, 20 }
        };

        int currentRow = 5;
        // Hansı sətirlər yaşıl olmalıdır (Sətir nömrələri: 01, 02, 09, 15)
        // Array indeksinə görə: 0, 1, 8, 14
        var greenRowIndices = new List<int> { 0, 1, 8, 14 };

        for (int idx = 0; idx < rowsData.Count; idx++)
        {
            var rowData = rowsData[idx];
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

            // Yaşıl rəng tətbiqi (Açıq yaşıl/sarımtıl - şəkildəki kimi)
            if (greenRowIndices.Contains(idx))
            {
                ws.Range(currentRow, 5, currentRow, totalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(226, 239, 218);
                // 01 və 02 üçün şəkildə bütün sətir yaşıl görünə bilər, amma əsasən rəqəm hissəsi.
                // Şəkilə dəqiq baxsaq 01, 02, 09, 15 bütün rəqəm sütunları yaşıldır.
            }

            // Sətir hündürlüyünü tənzimlə
            ws.Row(currentRow).Height = 25;

            currentRow++;
        }


        // ==========================================
        // 4. FORMATLAŞDIRMA VƏ SƏRHƏDLƏR
        // ==========================================

        // Sütun genişlikləri
        ws.Column(1).Width = 6;   // Tabeçilik
        ws.Column(2).Width = 20;  // Ad
        ws.Column(3).Width = 45;  // Göstəricilər
        ws.Column(4).Width = 6;   // No

        // Rəqəm sütunları
        for (int c = 5; c <= totalCols; c++) ws.Column(c).Width = 7;

        // Sərhədlər
        var tableRange = ws.Range(2, 1, currentRow - 1, totalCols);
        tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Göstəricilər sütununu sola düzləndir
        ws.Range(5, 3, currentRow - 1, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
        // Ad sütununu sola düzləndir
        ws.Range(5, 2, currentRow - 1, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

        // "o cümlədən" olan sətirləri bir az sağa çəkmək (Indent)
        foreach (var cell in ws.Column(3).CellsUsed())
        {
            if (cell.GetString().Contains("o cümlədən") ||
                cell.GetString().Contains("professorlar") ||
                cell.GetString().Contains("dosentlər") ||
                cell.GetString().Contains("rektor"))
            {
                // ClosedXML-də indent style var, amma sadəlik üçün buraxırıq
                // cell.Style.Alignment.Indent = 2; 
            }
        }

        // Rəhbər vəzifədə olanların "x" işarələrini mərkəzləşdir
        ws.Range(14, 9, 17, 10).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

        // ==========================================
        // 5. ÇIXIŞ (EXPORT)
        // ==========================================
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var bytes = ms.ToArray();

        return ReportFileViewModel.FileSuccess(bytes, "Isciler_Sayi_Hesabati.xlsx");
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