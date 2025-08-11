using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Models;
using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;

namespace ExcelDataManagementAPI.Services
{
    public interface IAuditService
    {
        Task LogChangeAsync(string fileName, string sheetName, int rowIndex, int originalRowId, 
                           string operationType, object? oldValue, object? newValue, 
                           string? modifiedBy, HttpContext? httpContext = null, 
                           string? changeReason = null, string[]? changedColumns = null);
        
        Task<List<GerceklesenRaporlar>> GetChangeHistoryAsync(string? fileName = null, 
                                                                  string? sheetName = null, 
                                                                  DateTime? fromDate = null, 
                                                                  DateTime? toDate = null,
                                                                  int limit = 100);
    }

    public class AuditService : IAuditService
    {
        private readonly ExcelDataContext _context;
        private readonly ILogger<AuditService> _logger;

        public AuditService(ExcelDataContext context, ILogger<AuditService> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task LogChangeAsync(string fileName, string sheetName, int rowIndex, int originalRowId,
                                       string operationType, object? oldValue, object? newValue,
                                       string? modifiedBy, HttpContext? httpContext = null,
                                       string? changeReason = null, string[]? changedColumns = null)
        {
            try
            {
                var auditLog = new GerceklesenRaporlar
                {
                    FileName = fileName,
                    SheetName = sheetName,
                    RowIndex = rowIndex,
                    OriginalRowId = originalRowId,
                    OperationType = operationType.ToUpper(),
                    OldValue = oldValue != null ? JsonSerializer.Serialize(oldValue) : null,
                    NewValue = newValue != null ? JsonSerializer.Serialize(newValue) : null,
                    ModifiedBy = modifiedBy,
                    ChangeDate = DateTime.UtcNow,
                    ChangeReason = changeReason,
                    ChangedColumns = changedColumns != null ? JsonSerializer.Serialize(changedColumns) : null,
                    IsSuccess = true
                };

                // HTTP context varsa kullanýcý bilgilerini al
                if (httpContext != null)
                {
                    auditLog.UserIP = GetClientIPAddress(httpContext);
                    auditLog.UserAgent = httpContext.Request.Headers["User-Agent"].ToString();
                }

                _context.GerceklesenRaporlar.Add(auditLog);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Audit log kaydedildi: {OperationType} - {FileName}/{SheetName} - Row {RowIndex}", 
                                     operationType, fileName, sheetName, rowIndex);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Audit log kaydedilirken hata: {OperationType} - {FileName}", operationType, fileName);
                
                // Audit log hatasý ana iþlemi durdurmamak için exception fýrlatmýyoruz
                // Sadece error log kaydediyoruz
                try
                {
                    var errorLog = new GerceklesenRaporlar
                    {
                        FileName = fileName,
                        SheetName = sheetName,
                        RowIndex = rowIndex,
                        OriginalRowId = originalRowId,
                        OperationType = operationType.ToUpper(),
                        ModifiedBy = modifiedBy,
                        ChangeDate = DateTime.UtcNow,
                        IsSuccess = false,
                        ErrorMessage = ex.Message
                    };

                    _context.GerceklesenRaporlar.Add(errorLog);
                    await _context.SaveChangesAsync();
                }
                catch (Exception innerEx)
                {
                    // Son çare olarak da hata verirse sessizce geçiyoruz
                    _logger.LogCritical(innerEx, "Audit error log bile kaydedilemedi!");
                }
            }
        }

        public async Task<List<GerceklesenRaporlar>> GetChangeHistoryAsync(string? fileName = null,
                                                                              string? sheetName = null,
                                                                              DateTime? fromDate = null,
                                                                              DateTime? toDate = null,
                                                                              int limit = 100)
        {
            try
            {
                var query = _context.GerceklesenRaporlar.AsQueryable();

                if (!string.IsNullOrEmpty(fileName))
                    query = query.Where(x => x.FileName == fileName);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(x => x.SheetName == sheetName);

                if (fromDate.HasValue)
                    query = query.Where(x => x.ChangeDate >= fromDate.Value);

                if (toDate.HasValue)
                    query = query.Where(x => x.ChangeDate <= toDate.Value);

                return await query
                    .OrderByDescending(x => x.ChangeDate)
                    .Take(limit)
                    .ToListAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Audit change history alýnýrken hata");
                return new List<GerceklesenRaporlar>();
            }
        }

        private string GetClientIPAddress(HttpContext context)
        {
            string? ipAddress = null;

            // X-Forwarded-For header'ýný kontrol et (proxy/load balancer arkasýnda)
            if (context.Request.Headers.ContainsKey("X-Forwarded-For"))
            {
                ipAddress = context.Request.Headers["X-Forwarded-For"].FirstOrDefault();
            }

            // X-Real-IP header'ýný kontrol et
            if (string.IsNullOrEmpty(ipAddress) && context.Request.Headers.ContainsKey("X-Real-IP"))
            {
                ipAddress = context.Request.Headers["X-Real-IP"].FirstOrDefault();
            }

            // Remote IP'yi al
            if (string.IsNullOrEmpty(ipAddress))
            {
                ipAddress = context.Connection.RemoteIpAddress?.ToString();
            }

            // IPv6 loopback'i IPv4'e çevir
            if (ipAddress == "::1")
            {
                ipAddress = "127.0.0.1";
            }

            return ipAddress ?? "Unknown";
        }
    }
}