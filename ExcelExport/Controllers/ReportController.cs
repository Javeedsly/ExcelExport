using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

// Həmin namespace-də olduğundan əmin olun
// (Əgər Program.cs-dən kənara çıxarsanız, 'using' əlavə etməlisiniz)

[Route("[controller]")] // Bu, /Report ünvanında işləməsini təmin edir
public class ReportController : Controller
{
    private readonly ReportService _reportService;

    // "Dependency Injection" ilə ReportService-i avtomatik əldə edirik
    public ReportController(ReportService reportService)
    {
        _reportService = reportService;
    }

    [HttpGet("Download")] // Bu Action /Report/Download ünvanında işləyəcək
    public async Task<IActionResult> DownloadComplexReport()
    {
        // Lazımi parametrləri hazırlayın
        var dsk = new ReportDSK(); // Məlumatları bazadan və ya başqa yerdən alın
        var userId = "test_user"; // İstifadəçini təyin edin

        // 1. Servis funksiyanızı çağırın
        ReportFileViewModel reportModel = await _reportService.ExportGetStudentsAndGraduatesReport(dsk, userId);

        // 2. Modelin içindən byte[] və fayl adını alıb "FileResult" qaytarın
        if (reportModel.IsSuccess && reportModel.FileBytes != null)
        {
            // Bu, brauzerə birbaşa fayl endirmə siqnalı göndərir
            return File(
                reportModel.FileBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // Excel üçün MIME type
                reportModel.FileName ?? "Hesabat.xlsx" // Fayl adı
            );
        }

        // Əgər nəsə səhv baş versə
        return Content("Hesabat yaradılarkən xəta baş verdi.");
    }
}