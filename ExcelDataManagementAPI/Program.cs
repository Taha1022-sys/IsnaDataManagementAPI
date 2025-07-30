using Microsoft.EntityFrameworkCore;
using Microsoft.OpenApi.Models;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Services;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

builder.WebHost.UseUrls("http://localhost:5002", "https://localhost:7002");

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo { Title = "Excel Data Management API", Version = "v1" });
});

builder.Services.AddDbContext<ExcelDataContext>(options =>
{
    var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");
    options.UseSqlServer(connectionString);
});

builder.Services.AddScoped<IExcelService, ExcelService>();
builder.Services.AddScoped<IDataComparisonService, DataComparisonService>();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();

try
{
    using var scope = app.Services.CreateScope();
    var context = scope.ServiceProvider.GetRequiredService<ExcelDataContext>();
    await context.Database.EnsureCreatedAsync();
    Console.WriteLine("✅ Veritabanı hazır!");
}
catch (Exception ex)
{
    Console.WriteLine($"❌ Veritabanı hatası: {ex.Message}");
}

app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Data Management API V1");
    c.RoutePrefix = "";
});

app.UseCors("AllowAll");
app.UseAuthorization();
app.MapControllers();

Console.WriteLine("🚀 API başlatıldı!");
Console.WriteLine("📖 Swagger UI: http://localhost:5002");
Console.WriteLine("🌐 API: http://localhost:5002/api");

app.Run();
