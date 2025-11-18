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

        int totalCols = 5;

        // ==========================================
        // 2. BAŞLIQ SƏTİRLƏRİ (HEADER)
        // ==========================================

        // --- Sətir 1: Əsas Başlıq ---
        ws.Range(1, 1, 1, totalCols).Merge().Value = "Təqaüd alan tələbələrin sayı";
        ws.Range(1, 1, 1, totalCols).Style.Font.SetBold().Font.SetFontSize(12);

        // --- Sətir 2, 3: Cədvəl Başlıqları ---

        // Sol tərəfdəki sabit sütunlar (Sətir 2-dən 3-ə qədər şaquli birləşir)
        ws.Range(2, 1, 3, 1).Merge().Value = "Təhsil müəssisələrinin adı";
        ws.Range(2, 2, 3, 2).Merge().Value = "Göstərici";
        ws.Range(2, 3, 3, 3).Merge().Value = "Cəmi";

        // "Onlardan" bloku (Üfüqi birləşmə)
        ws.Range(2, 4, 2, 5).Merge().Value = "onlardan";

        // Alt başlıqlar (Sətir 3)
        ws.Cell(3, 4).Value = "əlaçı təqaüdü alanlar";
        ws.Cell(3, 5).Value = "prezident təqaüdü alanlar";

        // Başlıqları qalınlaşdırırıq
        ws.Range(2, 1, 3, totalCols).Style.Font.SetBold();


        // ==========================================
        // 3. MƏLUMATLARIN DOLDURULMASI
        // ==========================================

        // --- Sətir 4: Boz rəngli sətir ---

        // 1-ci sütun (ADA Universiteti) - 4 və 5-ci sətirləri birləşdirir
        ws.Range(4, 1, 5, 1).Merge().Value = "ADA Universiteti";
        ws.Range(4, 1, 5, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left).Font.SetBold();

        // Digər sütunlar
        ws.Cell(4, 2).Value = "Təqaüd alan tələbələrin sayı";
        ws.Cell(4, 3).Value = 23;
        ws.Cell(4, 4).Value = ""; // Boş
        ws.Cell(4, 5).Value = 23;

        // Sətir 4-ün stili (Boz fon və Qalın şrift)
        var grayRow = ws.Range(4, 2, 4, 5);
        grayRow.Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217); // Açıq boz
        grayRow.Style.Font.SetBold();

        // --- Sətir 5: Ağ rəngli sətir ---
        ws.Cell(5, 2).Value = "-onlardan qadınlar";
        ws.Cell(5, 3).Value = 14;
        ws.Cell(5, 4).Value = ""; // Boş
        ws.Cell(5, 5).Value = 14;


        // ==========================================
        // 4. FORMATLAŞDIRMA
        // ==========================================

        // Sütun genişlikləri
        ws.Column(1).Width = 25;  // Uni adı
        ws.Column(2).Width = 35;  // Göstərici
        ws.Column(3).Width = 10;  // Cəmi
        ws.Column(4).Width = 20;  // Əlaçı
        ws.Column(5).Width = 20;  // Prezident

        // Sərhədlər (Borders)
        var tableRange = ws.Range(2, 1, 5, totalCols);
        tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // ==========================================
        // 5. ÇIXIŞ (EXPORT)
        // ==========================================
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var bytes = ms.ToArray();

        return ReportFileViewModel.FileSuccess(bytes, "Teqaud_Hesabati.xlsx");
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