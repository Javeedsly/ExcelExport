using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

// [Route] atributunu silib sadə MVC standartına keçək
// Bu, xəta riskini azaldır.
public class ReportController : Controller
{
    private readonly ReportService _reportService;

    public ReportController(ReportService reportService)
    {
        _reportService = reportService;
    }

    // Brauzerdə link: /Report/Download
    public async Task<IActionResult> Download()
    {
        // ... (əvvəlki kodunuz eynilə qalır) ...
        var dsk = new ReportDSK();
        var userId = "test_user";

        var reportModel = await _reportService.ExportGetStudentsAndGraduatesReport(dsk, userId);

        if (reportModel.IsSuccess && reportModel.FileBytes != null)
        {
            return File(
                reportModel.FileBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                reportModel.FileName ?? "Hesabat.xlsx"
            );
        }

        return Content("Xəta baş verdi");
    }
}