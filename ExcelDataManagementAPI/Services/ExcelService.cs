using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Text.Json;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Models;
using ExcelDataManagementAPI.Models.DTOs;

namespace ExcelDataManagementAPI.Services
{
    public class ExcelService : IExcelService
    {
        private readonly ExcelDataContext _context;
        private readonly IWebHostEnvironment _environment;
        private readonly ILogger<ExcelService> _logger;
        private readonly IAuditService _auditService;

        public ExcelService(ExcelDataContext context, IWebHostEnvironment environment, ILogger<ExcelService> logger, IAuditService auditService)
        {
            _context = context;
            _environment = environment;
            _logger = logger;
            _auditService = auditService;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private async Task<T> ExecuteWithRetryAsync<T>(Func<Task<T>> operation, int maxRetries = 3)
        {
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    return await operation();
                }
                catch (DbUpdateConcurrencyException ex) when (attempt < maxRetries)
                {
                    _logger.LogWarning("Concurrency exception (attempt {Attempt}/{MaxRetries}): {Message}", attempt, maxRetries, ex.Message);
                    
                    await Task.Delay(TimeSpan.FromMilliseconds(100 * attempt));
                    
                    foreach (var entry in _context.ChangeTracker.Entries())
                    {
                        if (entry.Entity != null)
                        {
                            entry.Reload();
                        }
                    }
                }
                catch (DbUpdateConcurrencyException ex) when (attempt == maxRetries)
                {
                    _logger.LogError(ex, "Concurrency exception after {MaxRetries} attempts", maxRetries);
                    throw new InvalidOperationException(
                        "Veritabanýnda eþzamanlýlýk sorunu oluþtu. Lütfen iþlemi tekrar deneyin.", ex);
                }
            }
            
            throw new InvalidOperationException("Beklenmeyen durum: retry logic tamamlanamadý");
        }

        public async Task<ExcelFile> UploadExcelFileAsync(IFormFile file, string? uploadedBy = null)
        {
            try
            {
                var uploadsPath = Path.Combine(_environment.ContentRootPath, "uploads");
                Directory.CreateDirectory(uploadsPath);

                var fileName = $"{Path.GetFileNameWithoutExtension(file.FileName)}_{DateTime.Now:yyyyMMddHHmmss}{Path.GetExtension(file.FileName)}";
                var filePath = Path.Combine(uploadsPath, fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                var excelFile = new ExcelFile
                {
                    FileName = fileName,
                    OriginalFileName = file.FileName,
                    FilePath = filePath,
                    FileSize = file.Length,
                    UploadedBy = uploadedBy,
                    UploadDate = DateTime.UtcNow,
                    IsActive = true
                };

                _context.ExcelFiles.Add(excelFile);
                await _context.SaveChangesAsync();

                return excelFile;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya yükleme hatasý");
                throw;
            }
        }

        public async Task<List<ExcelFile>> GetExcelFilesAsync()
        {
            return await _context.ExcelFiles
                .Where(f => f.IsActive)
                .OrderByDescending(f => f.UploadDate)
                .ToListAsync();
        }

        public async Task<Dictionary<string, List<ExcelDataResponseDto>>> ReadAllSheetsFromExcelAsync(string fileName)
        {
            return await ExecuteWithRetryAsync(async () =>
            {
                try
                {
                    fileName = Uri.UnescapeDataString(fileName);
                    
                    _logger.LogInformation("ReadAllSheetsFromExcelAsync baþlatýldý: FileName={FileName}", fileName);

                    var excelFile = await _context.ExcelFiles.FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
                    if (excelFile == null)
                    {
                        _logger.LogError("Dosya bulunamadý: {FileName}", fileName);
                        throw new FileNotFoundException($"Dosya bulunamadý: {fileName}");
                    }

                    if (!System.IO.File.Exists(excelFile.FilePath))
                    {
                        _logger.LogError("Fiziksel dosya bulunamadý: {FilePath}", excelFile.FilePath);
                        throw new FileNotFoundException($"Fiziksel dosya bulunamadý: {excelFile.FilePath}");
                    }

                    var allResults = new Dictionary<string, List<ExcelDataResponseDto>>();

                    using var package = new ExcelPackage(new FileInfo(excelFile.FilePath));
                    
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        _logger.LogError("Excel dosyasýnda hiç sheet bulunamadý");
                        throw new Exception("Excel dosyasýnda hiç sheet bulunamadý");
                    }

                    using var transaction = await _context.Database.BeginTransactionAsync();
                    
                    try
                    {
                        _logger.LogInformation("Dosyada {Count} sheet bulundu, tümü iþlenecek", package.Workbook.Worksheets.Count);

                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            _logger.LogInformation("Sheet iþleniyor: {SheetName}", worksheet.Name);

                            var sheetResults = new List<ExcelDataResponseDto>();

                            if (worksheet.Dimension == null)
                            {
                                _logger.LogWarning("Worksheet boþ, atlanýyor: {SheetName}", worksheet.Name);
                                allResults[worksheet.Name] = sheetResults;
                                continue;
                            }

                            var existingDataCount = await _context.ExcelDataRows
                                .Where(r => r.FileName == fileName && r.SheetName == worksheet.Name && !r.IsDeleted)
                                .CountAsync();
                            
                            if (existingDataCount > 0)
                            {
                                var existingData = await _context.ExcelDataRows
                                    .Where(r => r.FileName == fileName && r.SheetName == worksheet.Name && !r.IsDeleted)
                                    .ToListAsync();
                                
                                foreach (var row in existingData)
                                {
                                    row.IsDeleted = true;
                                    row.ModifiedDate = DateTime.UtcNow;
                                    row.ModifiedBy = "SYSTEM_READ_ALL_SHEETS";
                                }
                                
                                await _context.SaveChangesAsync();
                                _logger.LogInformation("Sheet için önceki veriler soft delete yapýldý: {SheetName} - {Count} kayýt", 
                                    worksheet.Name, existingDataCount);
                            }

                            var headers = new List<string>();
                            var columnCount = worksheet.Dimension.Columns;
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                                headers.Add(!string.IsNullOrEmpty(headerValue) ? headerValue : $"Column{col}");
                            }

                            _logger.LogInformation("Sheet {SheetName} için header'lar okundu: {Headers}", worksheet.Name, string.Join(", ", headers));

                            var rowCount = worksheet.Dimension.Rows;
                            var processedRows = 0;
                            var dataRowsToAdd = new List<ExcelDataRow>();
                            
                            for (int row = 2; row <= rowCount; row++)
                            {
                                var rowData = new Dictionary<string, string>();
                                bool hasData = false;

                                for (int col = 1; col <= headers.Count; col++)
                                {
                                    var cellValue = worksheet.Cells[row, col].Value;
                                    var stringValue = cellValue?.ToString()?.Trim() ?? "";
                                    rowData[headers[col - 1]] = stringValue;
                                    
                                    if (!string.IsNullOrWhiteSpace(stringValue))
                                        hasData = true;
                                }

                                if (hasData)
                                {
                                    var dataRow = new ExcelDataRow
                                    {
                                        FileName = fileName,
                                        SheetName = worksheet.Name,
                                        RowIndex = row,
                                        RowData = JsonSerializer.Serialize(rowData),
                                        CreatedDate = DateTime.UtcNow,
                                        Version = 1,
                                        IsDeleted = false
                                    };

                                    dataRowsToAdd.Add(dataRow);
                                    processedRows++;

                                    sheetResults.Add(new ExcelDataResponseDto
                                    {
                                        Id = 0, 
                                        FileName = fileName,
                                        SheetName = worksheet.Name,
                                        RowIndex = row,
                                        Data = rowData,
                                        CreatedDate = dataRow.CreatedDate,
                                        Version = dataRow.Version
                                    });
                                }
                            }

                            if (dataRowsToAdd.Any())
                            {
                                _context.ExcelDataRows.AddRange(dataRowsToAdd);
                                await _context.SaveChangesAsync();
                                _logger.LogInformation("Sheet {SheetName} için veriler veritabanýna kaydedildi: {ProcessedRows} satýr", 
                                    worksheet.Name, processedRows);
                            }

                            allResults[worksheet.Name] = sheetResults;
                        }

                        await transaction.CommitAsync();

                        _logger.LogInformation("Tüm sheet'ler baþarýyla iþlendi: {FileName}, toplam {SheetCount} sheet", 
                            fileName, allResults.Count);
                        
                        return allResults;
                    }
                    catch (Exception)
                    {
                        await transaction.RollbackAsync();
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Tüm sheet'ler okunurken hata: {FileName}", fileName);
                    throw;
                }
            });
        }

        public async Task<List<ExcelDataResponseDto>> ReadExcelDataAsync(string fileName, string? sheetName = null)
        {
            return await ExecuteWithRetryAsync(async () =>
            {
                try
                {
                    fileName = Uri.UnescapeDataString(fileName);
                    
                    _logger.LogInformation("ReadExcelDataAsync baþlatýldý: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

                    if (string.IsNullOrEmpty(sheetName))
                    {
                        _logger.LogInformation("Sheet belirtilmemiþ, tüm sheet'ler okunacak: {FileName}", fileName);
                        var allSheetsData = await ReadAllSheetsFromExcelAsync(fileName);
                        
                        var combinedResults = new List<ExcelDataResponseDto>();
                        foreach (var sheetData in allSheetsData.Values)
                        {
                            combinedResults.AddRange(sheetData);
                        }
                        
                        _logger.LogInformation("Tüm sheet'ler okundu: {FileName}, toplam {Count} kayýt", fileName, combinedResults.Count);
                        return combinedResults;
                    }

                    var excelFile = await _context.ExcelFiles.FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
                    if (excelFile == null)
                    {
                        _logger.LogError("Dosya bulunamadý: {FileName}", fileName);
                        throw new FileNotFoundException($"Dosya bulunamadý: {fileName}");
                    }

                    if (!System.IO.File.Exists(excelFile.FilePath))
                    {
                        _logger.LogError("Fiziksel dosya bulunamadý: {FilePath}", excelFile.FilePath);
                        throw new FileNotFoundException($"Fiziksel dosya bulunamadý: {excelFile.FilePath}");
                    }

                    var results = new List<ExcelDataResponseDto>();

                    using var package = new ExcelPackage(new FileInfo(excelFile.FilePath));
                    
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                    {
                        var availableSheets = package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
                        _logger.LogError("Sheet bulunamadý: {SheetName}. Mevcut sheet'ler: {AvailableSheets}", 
                            sheetName, string.Join(", ", availableSheets));
                        throw new Exception($"Sheet '{sheetName}' bulunamadý. Mevcut sheet'ler: {string.Join(", ", availableSheets)}");
                    }

                    if (worksheet.Dimension == null)
                    {
                        _logger.LogWarning("Worksheet boþ: {SheetName}", worksheet.Name);
                        return results;
                    }

                    using var transaction = await _context.Database.BeginTransactionAsync();
                    
                    try
                    {
                        var existingDataCount = await _context.ExcelDataRows
                            .Where(r => r.FileName == fileName && r.SheetName == sheetName && !r.IsDeleted)
                            .CountAsync();
                        
                        if (existingDataCount > 0)
                        {
                            var existingData = await _context.ExcelDataRows
                                .Where(r => r.FileName == fileName && r.SheetName == sheetName && !r.IsDeleted)
                                .ToListAsync();
                            
                            foreach (var row in existingData)
                            {
                                row.IsDeleted = true;
                                row.ModifiedDate = DateTime.UtcNow;
                                row.ModifiedBy = "SYSTEM_READ_OPERATION";
                            }
                            
                            await _context.SaveChangesAsync();
                            _logger.LogInformation("Önceki veriler soft delete yapýldý: {FileName}/{SheetName} - {Count} kayýt", 
                                fileName, sheetName, existingDataCount);
                        }

                        var headers = new List<string>();
                        var columnCount = worksheet.Dimension.Columns;
                        for (int col = 1; col <= columnCount; col++)
                        {
                            var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                            headers.Add(!string.IsNullOrEmpty(headerValue) ? headerValue : $"Column{col}");
                        }

                        _logger.LogInformation("Header'lar okundu: {Headers}", string.Join(", ", headers));

                        var rowCount = worksheet.Dimension.Rows;
                        var processedRows = 0;
                        var dataRowsToAdd = new List<ExcelDataRow>();
                        
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var rowData = new Dictionary<string, string>();
                            bool hasData = false;

                            for (int col = 1; col <= headers.Count; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Value;
                                var stringValue = cellValue?.ToString()?.Trim() ?? "";
                                rowData[headers[col - 1]] = stringValue;
                                
                                if (!string.IsNullOrWhiteSpace(stringValue))
                                    hasData = true;
                            }

                            if (hasData)
                            {
                                var dataRow = new ExcelDataRow
                                {
                                    FileName = fileName,
                                    SheetName = worksheet.Name,
                                    RowIndex = row,
                                    RowData = JsonSerializer.Serialize(rowData),
                                    CreatedDate = DateTime.UtcNow,
                                    Version = 1,
                                    IsDeleted = false
                                };

                                dataRowsToAdd.Add(dataRow);
                                processedRows++;

                                results.Add(new ExcelDataResponseDto
                                {
                                    Id = 0, 
                                    FileName = fileName,
                                    SheetName = worksheet.Name,
                                    RowIndex = row,
                                    Data = rowData,
                                    CreatedDate = dataRow.CreatedDate,
                                    Version = dataRow.Version
                                });
                            }
                        }

                        if (dataRowsToAdd.Any())
                        {
                            _context.ExcelDataRows.AddRange(dataRowsToAdd);
                            await _context.SaveChangesAsync();
                            _logger.LogInformation("Veriler veritabanýna kaydedildi: {ProcessedRows} satýr", processedRows);
                        }

                        await transaction.CommitAsync();

                        var savedRows = await _context.ExcelDataRows
                            .Where(r => r.FileName == fileName && r.SheetName == worksheet.Name && !r.IsDeleted)
                            .OrderBy(r => r.RowIndex)
                            .ToListAsync();

                        for (int i = 0; i < results.Count && i < savedRows.Count; i++)
                        {
                            results[i].Id = savedRows[i].Id;
                        }

                        _logger.LogInformation("Excel verisi baþarýyla okundu: {FileName}, {RowCount} satýr", fileName, results.Count);
                        return results;
                    }
                    catch (Exception)
                    {
                        await transaction.RollbackAsync();
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Excel dosyasý okunurken hata: {FileName}, SheetName: {SheetName}", fileName, sheetName);
                    throw;
                }
            });
        }

        public async Task<List<ExcelDataResponseDto>> GetExcelDataAsync(string fileName, string? sheetName = null, int page = 1, int pageSize = 50)
        {
            try
            {
                _logger.LogInformation("GetExcelDataAsync çaðrýldý: FileName={FileName}, SheetName={SheetName}, Page={Page}, PageSize={PageSize}", 
                    fileName, sheetName, page, pageSize);

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                    _logger.LogInformation("Sheet filtrelemesi uygulandý: {SheetName}", sheetName);
                }

                var totalCount = await query.CountAsync();
                _logger.LogInformation("Toplam kayýt sayýsý: {TotalCount}", totalCount);

                var rows = await query
                    .OrderBy(r => r.RowIndex)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .ToListAsync();

                _logger.LogInformation("Sayfalama sonrasý dönen kayýt sayýsý: {ReturnedCount}", rows.Count);

                var result = rows.Select(r => new ExcelDataResponseDto
                {
                    Id = r.Id,
                    FileName = r.FileName,
                    SheetName = r.SheetName,
                    RowIndex = r.RowIndex,
                    Data = JsonSerializer.Deserialize<Dictionary<string, string>>(r.RowData) ?? new Dictionary<string, string>(),
                    CreatedDate = r.CreatedDate,
                    ModifiedDate = r.ModifiedDate,
                    Version = r.Version,
                    ModifiedBy = r.ModifiedBy
                }).ToList();

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetExcelDataAsync'te hata: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);
                throw;
            }
        }

        public async Task<List<ExcelDataResponseDto>> GetAllExcelDataAsync(string fileName, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("GetAllExcelDataAsync çaðrýldý: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

                var fileExists = await _context.ExcelFiles
                    .AnyAsync(f => f.FileName == fileName && f.IsActive);

                if (!fileExists)
                {
                    _logger.LogWarning("Dosya bulunamadý veya aktif deðil: {FileName}", fileName);
                    throw new FileNotFoundException($"Dosya bulunamadý veya aktif deðil: {fileName}");
                }

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                    _logger.LogInformation("Sheet filtrelemesi uygulandý: {SheetName}", sheetName);
                }

                var totalCount = await query.CountAsync();
                _logger.LogInformation("Toplam kayýt sayýsý: {TotalCount}", totalCount);

                var rows = await query
                    .OrderBy(r => r.RowIndex)
                    .ToListAsync();

                var result = rows.Select(r => new ExcelDataResponseDto
                {
                    Id = r.Id,
                    FileName = r.FileName,
                    SheetName = r.SheetName,
                    RowIndex = r.RowIndex,
                    Data = JsonSerializer.Deserialize<Dictionary<string, string>>(r.RowData) ?? new Dictionary<string, string>(),
                    CreatedDate = r.CreatedDate,
                    ModifiedDate = r.ModifiedDate,
                    Version = r.Version,
                    ModifiedBy = r.ModifiedBy
                }).ToList();

                _logger.LogInformation("Tüm Excel verileri getirildi: {FileName}, {Count} kayýt", fileName, result.Count);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Tüm Excel verileri getirilirken hata: {FileName}, {SheetName}", fileName, sheetName);
                throw;
            }
        }

        public async Task<ExcelDataResponseDto> UpdateExcelDataAsync(ExcelDataUpdateDto updateDto, HttpContext? httpContext = null)
        {
            return await ExecuteWithRetryAsync(async () =>
            {
                _logger.LogInformation("UpdateExcelDataAsync baþlatýldý: ID={Id}, ModifiedBy={ModifiedBy}", updateDto.Id, updateDto.ModifiedBy);

                var dataRow = await _context.ExcelDataRows.FindAsync(updateDto.Id);
                if (dataRow == null)
                {
                    _logger.LogError("Güncellenecek veri bulunamadý: {Id}", updateDto.Id);
                    throw new Exception($"Veri bulunamadý: {updateDto.Id}");
                }

                if (dataRow.IsDeleted)
                {
                    _logger.LogError("Silinen veri güncellenmeye çalýþýlýyor: {Id}", updateDto.Id);
                    throw new Exception($"Bu veri silinmiþ durumda ve güncellenemez: {updateDto.Id}");
                }

                var oldData = JsonSerializer.Deserialize<Dictionary<string, string>>(dataRow.RowData);
                var newData = updateDto.Data;

                var changedColumns = new List<string>();
                foreach (var newItem in newData)
                {
                    if (oldData?.ContainsKey(newItem.Key) != true || oldData[newItem.Key] != newItem.Value)
                    {
                        changedColumns.Add(newItem.Key);
                    }
                }

                if (!changedColumns.Any())
                {
                    _logger.LogInformation("Hiç deðiþiklik yapýlmamýþ, güncelleme atlanýyor: {Id}", updateDto.Id);
                    return new ExcelDataResponseDto
                    {
                        Id = dataRow.Id,
                        FileName = dataRow.FileName,
                        SheetName = dataRow.SheetName,
                        RowIndex = dataRow.RowIndex,
                        Data = oldData ?? new Dictionary<string, string>(),
                        CreatedDate = dataRow.CreatedDate,
                        ModifiedDate = dataRow.ModifiedDate,
                        Version = dataRow.Version,
                        ModifiedBy = dataRow.ModifiedBy
                    };
                }

                var oldVersion = dataRow.Version;
                dataRow.RowData = JsonSerializer.Serialize(updateDto.Data);
                dataRow.ModifiedDate = DateTime.UtcNow;
                dataRow.ModifiedBy = updateDto.ModifiedBy;
                dataRow.Version++;

                await _context.SaveChangesAsync();

                _logger.LogInformation("Veri güncellendi: ID={Id}, Version {OldVersion} -> {NewVersion}, ChangedColumns={ChangedColumns}", 
                    updateDto.Id, oldVersion, dataRow.Version, string.Join(", ", changedColumns));

                try
                {
                    await _auditService.LogChangeAsync(
                        fileName: dataRow.FileName,
                        sheetName: dataRow.SheetName,
                        rowIndex: dataRow.RowIndex,
                        originalRowId: dataRow.Id,
                        operationType: "UPDATE",
                        oldValue: oldData,
                        newValue: newData,
                        modifiedBy: updateDto.ModifiedBy,
                        httpContext: httpContext,
                        changeReason: "Web sitesi üzerinden veri güncellemesi",
                        changedColumns: changedColumns.ToArray()
                    );
                }
                catch (Exception auditEx)
                {
                    _logger.LogError(auditEx, "Audit log kaydedilemedi, ancak veri güncelleme baþarýlý: {Id}", updateDto.Id);
                }

                return new ExcelDataResponseDto
                {
                    Id = dataRow.Id,
                    FileName = dataRow.FileName,
                    SheetName = dataRow.SheetName,
                    RowIndex = dataRow.RowIndex,
                    Data = updateDto.Data,
                    CreatedDate = dataRow.CreatedDate,
                    ModifiedDate = dataRow.ModifiedDate,
                    Version = dataRow.Version,
                    ModifiedBy = dataRow.ModifiedBy
                };
            });
        }

        public async Task<List<ExcelDataResponseDto>> BulkUpdateExcelDataAsync(BulkUpdateDto bulkUpdateDto, HttpContext? httpContext = null)
        {
            var results = new List<ExcelDataResponseDto>();

            foreach (var update in bulkUpdateDto.Updates)
            {
                var result = await UpdateExcelDataAsync(update, httpContext);
                results.Add(result);
            }

            return results;
        }

        public async Task<bool> DeleteExcelDataAsync(int id, string? deletedBy = null, HttpContext? httpContext = null)
        {
            var dataRow = await _context.ExcelDataRows.FindAsync(id);
            if (dataRow == null)
                return false;

            var oldData = JsonSerializer.Deserialize<Dictionary<string, string>>(dataRow.RowData);

            dataRow.IsDeleted = true;
            dataRow.ModifiedDate = DateTime.UtcNow;
            dataRow.ModifiedBy = deletedBy;

            await _context.SaveChangesAsync();

            await _auditService.LogChangeAsync(
                fileName: dataRow.FileName,
                sheetName: dataRow.SheetName,
                rowIndex: dataRow.RowIndex,
                originalRowId: dataRow.Id,
                operationType: "DELETE",
                oldValue: oldData,
                newValue: null,
                modifiedBy: deletedBy,
                httpContext: httpContext,
                changeReason: "Web sitesi üzerinden veri silme"
            );

            return true;
        }

        public async Task<bool> DeleteExcelFileAsync(string fileName, string? deletedBy = null, HttpContext? httpContext = null)
        {
            try
            {
                var excelFile = await _context.ExcelFiles
                    .FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
                
                if (excelFile == null)
                    return false;

                excelFile.IsActive = false;

                var relatedDataRows = await _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted)
                    .ToListAsync();

                foreach (var dataRow in relatedDataRows)
                {
                    dataRow.IsDeleted = true;
                    dataRow.ModifiedDate = DateTime.UtcNow;
                    dataRow.ModifiedBy = deletedBy;
                }

                await _context.SaveChangesAsync();

                if (File.Exists(excelFile.FilePath))
                {
                    File.Delete(excelFile.FilePath);
                    _logger.LogInformation("Fiziksel dosya silindi: {FilePath}", excelFile.FilePath);
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyasý silinirken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<ExcelDataResponseDto> AddExcelRowAsync(string fileName, string sheetName, Dictionary<string, string> rowData, string? addedBy = null, HttpContext? httpContext = null)
        {
            try
            {
                _logger.LogInformation("AddExcelRowAsync baþlatýldý: FileName={FileName}, SheetName={SheetName}, AddedBy={AddedBy}", 
                    fileName, sheetName, addedBy);

                var fileExists = await _context.ExcelFiles
                    .AnyAsync(f => f.FileName == fileName && f.IsActive);

                if (!fileExists)
                {
                    _logger.LogError("Dosya bulunamadý: {FileName}", fileName);
                    throw new FileNotFoundException($"Dosya bulunamadý: {fileName}");
                }

                var sheetExists = await _context.ExcelDataRows
                    .AnyAsync(r => r.FileName == fileName && r.SheetName == sheetName && !r.IsDeleted);

                if (!sheetExists)
                {
                    var availableSheets = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .Select(r => r.SheetName)
                        .Distinct()
                        .ToListAsync();

                    _logger.LogWarning("Sheet bulunamadý: {SheetName}. Mevcut sheet'ler: {AvailableSheets}", 
                        sheetName, string.Join(", ", availableSheets));
                    
                    if (!availableSheets.Any())
                    {
                        throw new InvalidOperationException($"Dosya '{fileName}' henüz okunmamýþ. Önce dosyayý okuyun.");
                    }
                    
                    throw new ArgumentException($"Sheet '{sheetName}' bulunamadý. Mevcut sheet'ler: {string.Join(", ", availableSheets)}");
                }

                var maxRowIndex = await _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && r.SheetName == sheetName && !r.IsDeleted)
                    .MaxAsync(r => (int?)r.RowIndex) ?? 1; 

                var newRowIndex = maxRowIndex + 1;

                var dataRow = new ExcelDataRow
                {
                    FileName = fileName,
                    SheetName = sheetName,
                    RowIndex = newRowIndex,
                    RowData = JsonSerializer.Serialize(rowData),
                    CreatedDate = DateTime.UtcNow,
                    ModifiedBy = addedBy,
                    Version = 1
                };

                _context.ExcelDataRows.Add(dataRow);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Yeni satýr veritabanýna eklendi: ID={Id}, RowIndex={RowIndex}", dataRow.Id, newRowIndex);

                try
                {
                    await _auditService.LogChangeAsync(
                        fileName: fileName,
                        sheetName: sheetName,
                        rowIndex: dataRow.RowIndex,
                        originalRowId: dataRow.Id,
                        operationType: "CREATE",
                        oldValue: null,
                        newValue: rowData,
                        modifiedBy: addedBy,
                        httpContext: httpContext,
                        changeReason: "Web sitesi üzerinden yeni veri ekleme"
                    );
                }
                catch (Exception auditEx)
                {
                    _logger.LogError(auditEx, "Audit log kaydedilemedi, ancak veri ekleme baþarýlý");
                }

                return new ExcelDataResponseDto
                {
                    Id = dataRow.Id,
                    FileName = fileName,
                    SheetName = sheetName,
                    RowIndex = dataRow.RowIndex,
                    Data = rowData,
                    CreatedDate = dataRow.CreatedDate,
                    Version = dataRow.Version,
                    ModifiedBy = addedBy
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "AddExcelRowAsync'te hata: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);
                throw;
            }
        }

        public async Task<byte[]> ExportToExcelAsync(ExcelExportRequestDto exportRequest)
        {
            var query = _context.ExcelDataRows
                .Where(r => r.FileName == exportRequest.FileName && !r.IsDeleted);

            if (!string.IsNullOrEmpty(exportRequest.SheetName))
                query = query.Where(r => r.SheetName == exportRequest.SheetName);

            if (exportRequest.RowIds?.Any() == true)
                query = query.Where(r => exportRequest.RowIds.Contains(r.Id));

            var rows = await query.OrderBy(r => r.RowIndex).ToListAsync();

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(exportRequest.SheetName ?? "Export");

            if (rows.Any())
            {
                var firstRowData = JsonSerializer.Deserialize<Dictionary<string, string>>(rows.First().RowData) ?? new Dictionary<string, string>();
                var headers = firstRowData.Keys.ToList();

                for (int i = 0; i < headers.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }

                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, string>>(rows[rowIndex].RowData) ?? new Dictionary<string, string>();
                    for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                    {
                        worksheet.Cells[rowIndex + 2, colIndex + 1].Value = rowData.ContainsKey(headers[colIndex]) ? rowData[headers[colIndex]] : "";
                    }
                }
            }

            return package.GetAsByteArray();
        }

        public async Task<List<string>> GetSheetsAsync(string fileName)
        {
            var excelFile = await _context.ExcelFiles.FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
            if (excelFile == null)
                throw new FileNotFoundException($"Dosya bulunamadý: {fileName}");

            using var package = new ExcelPackage(new FileInfo(excelFile.FilePath));
            return package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
        }

        public async Task<object> GetDataStatisticsAsync(string fileName, string? sheetName = null)
        {
            var query = _context.ExcelDataRows
                .Where(r => r.FileName == fileName && !r.IsDeleted);

            if (!string.IsNullOrEmpty(sheetName))
                query = query.Where(r => r.SheetName == sheetName);

            var totalRows = await query.CountAsync();
            var modifiedRows = await query.Where(r => r.ModifiedDate != null).CountAsync();
            var sheetsCount = await query.Select(r => r.SheetName).Distinct().CountAsync();

            return new
            {
                fileName = fileName,
                sheetName = sheetName,
                totalRows = totalRows,
                modifiedRows = modifiedRows,
                sheetsCount = sheetsCount,
                lastModified = await query.Where(r => r.ModifiedDate != null).MaxAsync(r => (DateTime?)r.ModifiedDate)
            };
        }
    }
}