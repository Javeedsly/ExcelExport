// ClosedXML və Servisimiz üçün 'using' əlavə edirik
using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;

var builder = WebApplication.CreateBuilder(args);

// === MVC Servislərini əlavə edirik ===
builder.Services.AddControllersWithViews();

// === Rəapor Servisimizi qeydiyyatdan keçiririk ===
// Bu, Controller-in ReportService-dən istifadə etməsinə imkan verəcək
builder.Services.AddScoped<ReportService>();


var app = builder.Build();

// Development mühiti üçün əlavə ayarlar
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles(); // wwwroot qovluğundakı CSS/JS üçün

app.UseRouting();

app.UseAuthorization();

// === MVC "Routing" (Marşrutlama) ===
// Bu, brauzer sorğularını (məs. /Home/Index) düzgün Controller-ə yönləndirir
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();


// ===================================================================
//  KÖMƏKÇİ SİNİFLƏR (BUNLARI AYRI FAYLLARA DA KÖÇÜRƏ BİLƏRSİNİZ)
// ===================================================================

// --- Yuxarıda yazdığımız Excel kodunu saxlayan Servis ---
public class ReportService
{
    // (Burada sizin _dbContext və s. ola bilər)
    // public ReportService(ApplicationDbContext context) { ... }

    public async Task<ReportFileViewModel> ExportGetStudentsAndGraduatesReport(ReportDSK dsk, string userId)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Tələbə və Məzun Sayı");

        // ... (YUXARIDAKI BÜTÜN EXCEL KODU BURAYA GƏLİR) ...
        // ws.Range(1, 1, 1, totalCols).Merge()...
        // ...
        // ws.Cell(currentRow, 44).Value = 39;
        // ...

        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var bytes = ms.ToArray();

        var fileName = "Tələbə və Məzun Sayı (1 Oktyabr 2024).xlsx";

        // await ReportExportLog(ms, dsk, userId, fileName); 
        // await _dbContext.SaveChangesAsync();

        return ReportFileViewModel.FileSuccess(bytes, fileName);
    }
}


// --- Təmsili Modellər (Bunlar sizdə yəqin ki "Models" qovluğundadır) ---
public class ReportDSK { /* Sizin məlumat strukturunuz */ }

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