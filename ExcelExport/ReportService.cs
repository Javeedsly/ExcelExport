using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;

public class ReportService
{
    public async Task<ReportFileViewModel> ExportStudentAdmissionReport(ReportDSK dsk, string userId)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Hesabat");

        // === ÜMUMİ STİLLƏR ===
        var styles = ws.Style;
        styles.Font.FontName = "Calibri";
        styles.Font.FontSize = 10;
        styles.Alignment.WrapText = true;
        styles.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        styles.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

        // === SÜTUNLARIN TƏYİNİ (Şəkilə əsasən 33 sütun) ===
        int totalCols = 33;

        // --- Sətir 1: Əsas Başlıq ---
        ws.Range(1, 1, 1, totalCols).Merge().Value = "1 oktyabr 2024-cü il vəziyyətinə tələbələrin və məzunların ixtisaslar üzrə sayı";
        ws.Range(1, 1, 1, totalCols).Style.Font.SetBold().Font.SetFontSize(14)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

        // === SƏTİR 2, 3, 4: BAŞLIQLAR ===

        // 1. Sabit Sol Sütunlar (Sətir 2-4)
        ws.Range(2, 1, 4, 1).Merge().Value = "Tabeçilik";
        ws.Range(2, 2, 4, 2).Merge().Value = "Təhsil müəssisəsinin adı";
        ws.Range(2, 3, 4, 3).Merge().Value = "Təhsil forması";
        ws.Range(2, 4, 4, 4).Merge().Value = "İxtisasların adı";
        ws.Range(2, 5, 4, 5).Merge().Value = "İxtisasın kodu";
        ws.Range(2, 6, 4, 6).Merge().Value = "Qəbul planı";

        // 2. "Qəbul olunub" Bloku (Sətir 2, Sütun 7-12)
        ws.Range(2, 7, 2, 12).Merge().Value = "Qəbul olunub 1)";

        ws.Range(3, 7, 4, 7).Merge().Value = "DİM xətti ilə";
        ws.Range(3, 8, 4, 8).Merge().Value = "onlardan ödənişli əsaslarla (süt. 2)";
        ws.Range(3, 9, 4, 9).Merge().Value = "DİM xətti ilə qəbul olanlardan qadınlar (süt. 2)";
        ws.Range(3, 10, 4, 10).Merge().Value = "onlardan ödənişli əsaslarla (süt. 4)";
        ws.Range(3, 11, 4, 11).Merge().Value = "Təkrar təhsil";
        ws.Range(3, 12, 4, 12).Merge().Value = "onlardan qadınlar";

        // 3. "Kurslar üzrə təhsil alanlar" Bloku (Sətir 2, Sütun 13-24)
        ws.Range(2, 13, 2, 24).Merge().Value = "kurslar üzrə təhsil alanlar";

        string[] romanNumerals = { "I", "II", "III", "IV", "V", "VI" };
        for (int i = 0; i < 6; i++)
        {
            int startCol = 13 + (i * 2);
            ws.Range(3, startCol, 3, startCol + 1).Merge().Value = romanNumerals[i];
            ws.Cell(4, startCol + 1).Value = "onlardan qadınlar";
        }

        // 4. Yekun Statistikalar (Sətir 2-4 birləşir)
        ws.Range(2, 25, 4, 25).Merge().Value = "Bütün kurslarda təhsil alanlar, (süt. 8,10,12,14, 16,18 cəmi)";
        ws.Range(2, 26, 4, 26).Merge().Value = "onlardan ödənişli əsaslarla təhsil alanlar (süt. 20)";
        ws.Range(2, 27, 4, 27).Merge().Value = "Cəmi təhsil alanlardan qadınlar (süt. 20)";
        ws.Range(2, 28, 4, 28).Merge().Value = "onlardan ödənişli əsaslarla (süt. 22)";

        // 5. Buraxılış Bloku (Sətir 2, Sütun 29-32)
        ws.Range(2, 29, 2, 32).Merge().Value = "01.10.2023-cü ildən 01.10.2024-cü ilədək faktiki buraxılış";

        ws.Range(3, 29, 4, 29).Merge().Value = "Yekun dövlət attestasiyasına buraxılanlar";
        ws.Range(3, 30, 4, 30).Merge().Value = "onlardan qadınlar";
        ws.Range(3, 31, 4, 31).Merge().Value = "Bakalavr diplomu alanlar";
        ws.Range(3, 32, 4, 32).Merge().Value = "onlardan qadınlar";

        // 6. Gözlənilən Buraxılış (Sətir 2-4 birləşir)
        ws.Range(2, 33, 4, 33).Merge().Value = "01.10.2024-cü ildən 01.10.2025-ci ilədək gözlənilən buraxılış";

        // === FORMATLAŞDIRMA ===
        var headerRange = ws.Range(2, 1, 4, totalCols);
        headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        headerRange.Style.Font.SetBold();

        // === NÜMUNƏ MƏLUMAT (Şəkildəki ADA Universiteti sətri) ===
        int row = 5;

        ws.Cell(row, 1).Value = "DDQ";
        ws.Cell(row, 2).Value = "ADA Universiteti";
        ws.Cell(row, 3).Value = "Əyani";
        ws.Cell(row, 4).Value = "Beynəlxalq münasibətlər";
        ws.Cell(row, 5).Value = "050201";
        ws.Cell(row, 6).Value = 80;
        ws.Cell(row, 7).Value = 79;
        ws.Cell(row, 8).Value = 79;
        ws.Cell(row, 9).Value = 48;
        ws.Cell(row, 10).Value = 48;
        ws.Cell(row, 13).Value = 94;
        ws.Cell(row, 14).Value = 56;
        ws.Cell(row, 25).Value = 343;
        ws.Cell(row, 25).Style.Fill.BackgroundColor = XLColor.LightYellow;
        ws.Cell(row, 26).Value = 343;
        ws.Cell(row, 27).Value = 191;
        ws.Cell(row, 27).Style.Fill.BackgroundColor = XLColor.LightYellow;
        ws.Cell(row, 28).Value = 191;
        ws.Cell(row, 29).Value = 56;
        ws.Cell(row, 30).Value = 39;
        ws.Cell(row, 31).Value = 56;
        ws.Cell(row, 32).Value = 39;
        ws.Cell(row, 33).Value = 81;

        // Sərhədlər
        ws.Range(row, 1, row, totalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range(row, 1, row, totalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Sütun Genişlikləri
        ws.Column(1).Width = 6;
        ws.Column(2).Width = 25;
        ws.Column(4).Width = 35;
        ws.Column(5).Width = 10;
        for (int i = 6; i <= 33; i++) ws.Column(i).Width = 6;
        ws.Column(9).Width = 10;
        ws.Column(25).Width = 12;

        // === YADDA SAXLAMA ===
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