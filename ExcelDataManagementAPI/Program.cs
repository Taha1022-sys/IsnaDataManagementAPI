using Microsoft.EntityFrameworkCore;
using Microsoft.OpenApi.Models;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Services;
using OfficeOpenXml;

namespace ExcelDataManagementAPI
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // EPPlus lisansı
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Servis kayıtları
            builder.Services.AddControllers();
            builder.Services.AddEndpointsApiExplorer();
            
            // Swagger konfigürasyonu
            builder.Services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo
                {
                    Title = "Excel Data Management API",
                    Version = "v1",
                    Description = "Excel dosyalarını yönetmek ve karşılaştırmak için API"
                });
            });

            // Veritabanı bağlantısı
            builder.Services.AddDbContext<ExcelDataContext>(options =>
            {
                options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"));
            });

            // Dependency Injection
            builder.Services.AddScoped<IExcelService, ExcelService>();
            builder.Services.AddScoped<IDataComparisonService, DataComparisonService>();

            // Dosya upload konfigürasyonu
            builder.Services.Configure<IISServerOptions>(options =>
            {
                options.MaxRequestBodySize = 100 * 1024 * 1024; // 100MB
            });

            // CORS - Frontend için özel konfigürasyon
            builder.Services.AddCors(options =>
            {
                options.AddPolicy("ApiPolicy", policy =>
                {
                    policy.WithOrigins(
                            "http://localhost:5174",   // Frontend URL
                            "http://localhost:3000",   // React development server (alternatif)
                            "http://localhost:5173",   // Vite development server (alternatif)
                            "https://localhost:7002",  // Backend HTTPS URL (kendi kendine istek için)
                            "http://localhost:5002"    // Backend HTTP URL (kendi kendine istek için)
                        )
                        .AllowAnyHeader()
                        .AllowAnyMethod()
                        .AllowCredentials();  // Credentials desteği
                });

                // Development için daha esnek policy
                options.AddPolicy("DevelopmentPolicy", policy =>
                {
                    policy.AllowAnyOrigin()
                          .AllowAnyHeader()
                          .AllowAnyMethod();
                });
            });

            var app = builder.Build();

            // Development ortamı
            if (app.Environment.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                app.UseSwagger();
                app.UseSwaggerUI(c =>
                {
                    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Data Management API v1");
                    c.RoutePrefix = "swagger";
                });
                
                // Development ortamında daha esnek CORS policy kullan
                app.UseCors("DevelopmentPolicy");
            }
            else
            {
                app.UseHttpsRedirection();
                // Production ortamında güvenli CORS policy kullan
                app.UseCors("ApiPolicy");
            }

            // Static files
            app.UseStaticFiles();

            // Middleware sırası - CORS'u Authorization'dan önce kullan
            app.UseAuthorization();
            app.MapControllers();

            // Veritabanı kontrolü
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

            Console.WriteLine("🚀 Excel Data Management API başlatıldı!");
            Console.WriteLine("📖 Swagger UI: http://localhost:5002/swagger");
            Console.WriteLine("🌐 API Base URL: http://localhost:5002/api");
            Console.WriteLine("🔒 HTTPS Swagger UI: https://localhost:7002/swagger");
            Console.WriteLine("🔒 HTTPS API Base URL: https://localhost:7002/api");
            Console.WriteLine("🌐 Frontend URL: http://localhost:5174");
            Console.WriteLine("✅ CORS yapılandırması aktif - Frontend bağlantısı hazır!");
            Console.WriteLine("💡 LaunchSettings.json'daki portlar kullanılıyor");

            app.Run();
        }
    }
}