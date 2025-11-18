// Brauzerdə link: /Report/Download
using Microsoft.AspNetCore.Mvc;

public async Task<IActionResult> Download()
{
    var dsk = new ReportDSK();
    var userId = "test_user";

    // DİQQƏT: Metodun adı dəyişdi
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