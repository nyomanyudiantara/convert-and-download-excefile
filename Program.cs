using DinkToPdf;
using DinkToPdf.Contracts;

var builder = WebApplication.CreateBuilder(args);

// Register MVC services
builder.Services.AddControllersWithViews();

// ✅ Register DinkToPdf service before Build()
builder.Services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));

var app = builder.Build();

// Enable static files, routing, etc.
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

// Set default route
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
