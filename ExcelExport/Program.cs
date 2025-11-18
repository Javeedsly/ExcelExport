using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;

var builder = WebApplication.CreateBuilder(args);

// === MVC Servislərini əlavə edirik ===
builder.Services.AddControllersWithViews();

// === Report Servisimizi qeydiyyatdan keçiririk ===
builder.Services.AddScoped<ReportService>();

var app = builder.Build();

// Development mühiti üçün əlavə ayarlar
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();