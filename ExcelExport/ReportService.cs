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
        allCells.Font.FontSize = 11;
        allCells.Alignment.WrapText = true;
        allCells.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        allCells.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

        int totalCols = 9;

        // ==========================================
        // 2. BAŞLIQ SƏTİRLƏRİ (HEADER)
        // ==========================================

        // --- Sətir 1: Əsas Başlıq ---
        ws.Range(1, 1, 1, totalCols).Merge().Value = "Tədris aparıldığı dilə görə tələbələrin bölgüsü";
        ws.Range(1, 1, 1, totalCols).Style.Font.SetBold().Font.SetFontSize(12);

        // --- Sətir 2, 3, 4: Cədvəl Başlıqları ---

        // Sol tərəfdəki sabit sütunlar (Sətir 2-dən 4-ə qədər şaquli birləşir)
        ws.Range(2, 1, 4, 1).Merge().Value = "Tabeçilik";
        ws.Range(2, 2, 4, 2).Merge().Value = "Təhsil müəssisələrinin adı";
        ws.Range(2, 3, 4, 3).Merge().Value = "Tədris dili";

        // "Bakalavr" bloku (Üfüqi birləşmə)
        ws.Range(2, 4, 2, 9).Merge().Value = "Bakalavr";

        // Alt kateqoriyalar (Sətir 3)
        ws.Range(3, 4, 3, 5).Merge().Value = "Əyani";
        ws.Range(3, 6, 3, 7).Merge().Value = "Qiyabi";
        ws.Range(3, 8, 3, 9).Merge().Value = "Cəmi";

        // Sütun başlıqları (Sətir 4)
        ws.Cell(4, 4).Value = "Cəmi tələbə sayı";
        ws.Cell(4, 5).Value = "onlardan qadınlar";
        ws.Cell(4, 6).Value = "Cəmi tələbə sayı";
        ws.Cell(4, 7).Value = "onlardan qadınlar";
        ws.Cell(4, 8).Value = "Cəmi tələbə sayı";
        ws.Cell(4, 9).Value = "onlardan qadınlar";

        // Bütün başlıqları qalın etmək
        ws.Range(2, 1, 4, totalCols).Style.Font.SetBold();


        // ==========================================
        // 3. MƏLUMATLARIN DOLDURULMASI
        // ==========================================

        var rowsData = new List<object[]>
        {
            new object[] { "DDQ", "ADA Universiteti", "Azərbaycan dili", null, null, null, null, null, null },
            new object[] { "DDQ", "ADA Universiteti", "Rus dili", null, null, null, null, null, null },
            new object[] { "DDQ", "ADA Universiteti", "Türk dili", null, null, null, null, null, null },
            new object[] { "DDQ", "ADA Universiteti", "İngilis dili", 3548, 1816, null, null, 3548, 1816 },
            new object[] { "DDQ", "ADA Universiteti", "Fransız", null, null, null, null, null, null },
            new object[] { "DDQ", "ADA Universiteti", "Alman", null, null, null, null, null, null },
            // Sonuncu sətir
            new object[] { "DDQ", "ADA Universiteti", "Cəmi", 3548, 1816, null, null, 3548, 1816 }
        };

        int currentRow = 5;
        foreach (var rowData in rowsData)
        {
            for (int i = 0; i < rowData.Length; i++)
            {
                if (rowData[i] != null)
                {
                    // Rəqəm olub-olmadığını yoxlayıb daxil edirik
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

            // Əgər sətir "Cəmi" sətridirsə, qalın şriftlə yazaq
            if (rowData[2] != null && rowData[2].ToString() == "Cəmi")
            {
                ws.Range(currentRow, 1, currentRow, totalCols).Style.Font.SetBold();
            }

            currentRow++;
        }


        // ==========================================
        // 4. FORMATLAŞDIRMA
        // ==========================================

        // Sütun genişlikləri
        ws.Column(1).Width = 8;   // Tabeçilik
        ws.Column(2).Width = 30;  // Uni adı
        ws.Column(3).Width = 15;  // Dil

        // Rəqəm sütunları (4-9)
        for (int c = 4; c <= 9; c++) ws.Column(c).Width = 11;

        // Sərhədlər (Borders)
        var tableRange = ws.Range(2, 1, currentRow - 1, totalCols);
        tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Mətnlərin düzləndirilməsi (Alignment)
        // Ad və Dil sütunlarını sola düzləndiririk
        ws.Range(5, 2, currentRow - 1, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
        ws.Range(5, 3, currentRow - 1, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

        // ==========================================
        // 5. ÇIXIŞ (EXPORT)
        // ==========================================
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var bytes = ms.ToArray();

        return ReportFileViewModel.FileSuccess(bytes, "Tedris_Dili_Hesabati.xlsx");
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