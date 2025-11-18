using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

public class ReportController : Controller
{
    private readonly ReportService _reportService;

    // Servisi bura daxil edirik (Dependency Injection)
    public ReportController(ReportService reportService)
    {
        _reportService = reportService;
    }

    // Brauzerdə link: /Report/Download
    public async Task<IActionResult> Download()
    {
        var dsk = new ReportDSK();
        var userId = "test_user";

        // ReportService-dəki yeni metodu çağırırıq
        var reportModel = await _reportService.ExportStudentAdmissionReport(dsk, userId);

        if (reportModel.IsSuccess && reportModel.FileBytes != null)
        {
            return File(
                reportModel.FileBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                reportModel.FileName ?? "Hesabat.xlsx"
            );
        }

        return Content("Xəta baş verdi: Fayl yaradıla bilmədi.");
    }
}