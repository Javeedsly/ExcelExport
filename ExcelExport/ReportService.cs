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
        allCells.Font.FontSize = 10;
        allCells.Alignment.WrapText = true;
        allCells.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        allCells.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

        int totalCols = 23; // Şəkildə 23 sütun var

        // ==========================================
        // 2. BAŞLIQ SƏTİRLƏRİ (HEADER)
        // ==========================================

        // --- Sətir 1: Əsas Başlıq ---
        ws.Range(1, 1, 1, totalCols).Merge().Value = "Tələbələrin yaş tərkibinə görə bölgüsü (01.01.2024-cü ilə tam yaşı tamam olanlar)";
        ws.Range(1, 1, 1, totalCols).Style.Font.SetBold().Font.SetFontSize(12)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

        // "(nəfər)" yazısı sağ tərəfdə (Cəmi sütununun üstündə)
        ws.Cell(1, totalCols).Value = "(nəfər)";
        ws.Cell(1, totalCols).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right).Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);

        // Başlığın olduğu xananı təmizləyib yenidən birləşdirək ki, "(nəfər)" görünsün, 
        // amma şəkildəki kimi əsas başlıq solda, "nəfər" sağda olsun deyə, əslində 1-22-ni birləşdirmək daha səliqəli olar.
        // Lakin şəkildəki kimi sadə saxlayaq:
        ws.Range(1, 1, 1, totalCols - 1).Merge().Value = "Tələbələrin yaş tərkibinə görə bölgüsü (01.01.2024-cü ilə tam yaşı tamam olanlar)";
        ws.Range(1, 1, 1, totalCols - 1).Style.Font.SetBold().Font.SetFontSize(11).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);


        // --- Sətir 2: Sütun Başlıqları ---
        ws.Cell(2, 1).Value = "Tabeçilik";
        ws.Cell(2, 2).Value = "Təhsil müəssisəsinin adı";
        ws.Cell(2, 3).Value = "Təhsil forması";
        ws.Cell(2, 4).Value = "Göstəricilər";

        // Yaş sütunları
        ws.Cell(2, 5).Value = "15 yaş və az";
        ws.Cell(2, 6).Value = "16 yaş";
        ws.Cell(2, 7).Value = "17 yaş";
        ws.Cell(2, 8).Value = "18 yaş";
        ws.Cell(2, 9).Value = "19 yaş";
        ws.Cell(2, 10).Value = "20 yaş";
        ws.Cell(2, 11).Value = "21 yaş";
        ws.Cell(2, 12).Value = "22 yaş";
        ws.Cell(2, 13).Value = "23 yaş";
        ws.Cell(2, 14).Value = "24 yaş";
        ws.Cell(2, 15).Value = "25 yaş";
        ws.Cell(2, 16).Value = "26 yaş";
        ws.Cell(2, 17).Value = "27 yaş";
        ws.Cell(2, 18).Value = "28 yaş";
        ws.Cell(2, 19).Value = "29 yaş";
        ws.Cell(2, 20).Value = "30-34 yaş";
        ws.Cell(2, 21).Value = "35-39 yaş";
        ws.Cell(2, 22).Value = "40 yaş və daha";

        ws.Cell(2, 23).Value = "Cəmi";


        // ==========================================
        // 3. MƏLUMATLARIN DOLDURULMASI
        // ==========================================

        var rowsData = new List<object[]>
        {
            // R1: Cəmi tələbələrin sayı
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Cəmi tələbələrin sayı",
                           null, 12, 745, 785, 719, 634, 365, 167, 69, 29, 4, 7, 2, 3, 2, 3, 2, null, 3548 },

            // R2: -onlardan qadınlar
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "-onlardan qadınlar",
                           null, 8, 373, 381, 378, 360, 199, 73, 22, 13, 1, 2, 1, null, 2, 3, null, null, 1816 },

            // R3: Qəbul olunanların sayı
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Qəbul olunanların sayı",
                           null, 12, 735, 222, 29, 13, 2, 3, 2, 2, 1, null, null, null, 1, 1, null, null, 1023 },

            // R4: -onlardan qadınlar
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "-onlardan qadınlar",
                           null, 8, 379, 104, 18, 6, 1, null, 1, null, null, null, null, null, 1, 1, null, null, 519 },

            // R5: Bitirənlərin sayı
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "Bitirənlərin sayı",
                           null, null, null, null, null, 5, 86, 208, 106, 27, 6, 3, null, 1, null, null, null, null, 442 },

            // R6: -onlardan qadınlar
            new object[] { "DDQ", "ADA Universiteti", "Əyani", "-onlardan qadınlar",
                           null, null, null, null, null, 2, 54, 116, 58, 10, 1, 3, null, 1, null, null, null, null, 245 }
        };

        int currentRow = 3; // Məlumatlar 3-cü sətirdən başlayır
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

            // "Cəmi" sütunu (Sütun 23) - Yaşıl rəng
            ws.Cell(currentRow, 23).Style.Fill.BackgroundColor = XLColor.FromArgb(226, 239, 218); // Açıq yaşıl

            currentRow++;
        }

        // ==========================================
        // 4. FORMATLAŞDIRMA
        // ==========================================

        // Sütun genişlikləri
        ws.Column(1).Width = 6;   // Tabeçilik
        ws.Column(2).Width = 20;  // Uni adı
        ws.Column(3).Width = 8;   // Forma
        ws.Column(4).Width = 25;  // Göstəricilər

        // Yaş sütunları (5-dən 22-yə kimi) - Dar
        for (int c = 5; c <= 22; c++) ws.Column(c).Width = 4;

        // Cəmi sütunu
        ws.Column(23).Width = 8;

        // Sərhədlər (Bütün cədvəl üçün)
        var tableRange = ws.Range(2, 1, currentRow - 1, totalCols);
        tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Başlıq (Sətir 2) qalın şrift
        ws.Range(2, 1, 2, totalCols).Style.Font.SetBold();

        // "Cəmi" başlığı (Sətir 2, Sütun 23) - Yaşıl
        ws.Cell(2, 23).Style.Fill.BackgroundColor = XLColor.FromArgb(226, 239, 218);

        // Mətnləri sola düzləndirmək (Ad və Göstəricilər)
        ws.Range(3, 2, currentRow - 1, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left); // Uni adı
        ws.Range(3, 4, currentRow - 1, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left); // Göstəricilər

        // ==========================================
        // 5. ÇIXIŞ (EXPORT)
        // ==========================================
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var bytes = ms.ToArray();

        return ReportFileViewModel.FileSuccess(bytes, "Telebe_Yas_Hesabati.xlsx");
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