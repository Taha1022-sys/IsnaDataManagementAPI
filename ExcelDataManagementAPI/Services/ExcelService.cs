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

        public ExcelService(ExcelDataContext context, IWebHostEnvironment environment, ILogger<ExcelService> logger)
        {
            _context = context;
            _environment = environment;
            _logger = logger;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<ExcelFile> UploadExcelFileAsync(IFormFile file, string? uploadedBy = null)
        {
            try
            {
                var uploadsPath = Path.Combine(_environment.ContentRootPath, "uploads");
                Directory.CreateDirectory(uploadsPath);

                // Orijinal dosya ad�ndan g�venli bir ad olu�tur
                var originalFileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);
                var originalExtension = Path.GetExtension(file.FileName);
                
                // Dosya ad�n� temizle - g�venlik i�in
                var cleanFileName = string.Join("_", originalFileNameWithoutExtension.Split(Path.GetInvalidFileNameChars()));
                
                // Benzersiz dosya ad� olu�tur
                var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                var fileName = $"{cleanFileName}_{timestamp}{originalExtension}";
                
                // Dosya ad�n�n �ok uzun olmamas�n� sa�la
                if (fileName.Length > 200)
                {
                    cleanFileName = cleanFileName.Substring(0, Math.Min(cleanFileName.Length, 150));
                    fileName = $"{cleanFileName}_{timestamp}{originalExtension}";
                }

                var filePath = Path.Combine(uploadsPath, fileName);

                _logger.LogInformation("Dosya y�kleniyor: Original={OriginalName}, Generated={GeneratedName}, Path={FilePath}", 
                    file.FileName, fileName, filePath);

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

                _logger.LogInformation("Dosya ba�ar�yla y�klendi: {FileName} -> {GeneratedFileName}", file.FileName, fileName);

                return excelFile;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya y�kleme hatas�: {FileName}", file.FileName);
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

        public async Task<List<ExcelDataResponseDto>> ReadExcelDataAsync(string fileName, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("ReadExcelDataAsync ba�lat�ld�: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

                // Input validation
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("Dosya ad� bo� olamaz");
                }

                // Dosyay� veritaban�ndan bul
                var excelFile = await _context.ExcelFiles.FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
                if (excelFile == null)
                {
                    _logger.LogError("Dosya veritaban�nda bulunamad�: {FileName}", fileName);
                    throw new FileNotFoundException($"Dosya bulunamad�: {fileName}");
                }

                // Fiziksel dosya kontrol�
                if (string.IsNullOrEmpty(excelFile.FilePath) || !File.Exists(excelFile.FilePath))
                {
                    _logger.LogError("Fiziksel dosya bulunamad�. DbPath: {DbPath}, Exists: {Exists}", 
                        excelFile.FilePath, !string.IsNullOrEmpty(excelFile.FilePath) && File.Exists(excelFile.FilePath));
                    throw new FileNotFoundException($"Fiziksel dosya bulunamad�: {excelFile.FilePath}");
                }

                _logger.LogInformation("Dosya bulundu: DbFileName={DbFileName}, OriginalName={OriginalName}, FilePath={FilePath}", 
                    excelFile.FileName, excelFile.OriginalFileName, excelFile.FilePath);

                // �nceki verileri temizle (yeniden okuma durumunda)
                var existingData = await _context.ExcelDataRows
                    .Where(r => r.FileName == fileName)
                    .ToListAsync();
                
                if (existingData.Any())
                {
                    _context.ExcelDataRows.RemoveRange(existingData);
                    await _context.SaveChangesAsync();
                    _logger.LogInformation("�nceki veriler temizlendi: {FileName}, {Count} kay�t", fileName, existingData.Count);
                }

                var results = new List<ExcelDataResponseDto>();

                // Excel dosyas�n� a�
                using var package = new ExcelPackage(new FileInfo(excelFile.FilePath));
                
                if (package.Workbook.Worksheets.Count == 0)
                {
                    _logger.LogWarning("Excel dosyas�nda hi� worksheet bulunamad�: {FileName}", fileName);
                    throw new Exception("Excel dosyas�nda hi� worksheet bulunamad�");
                }

                // Worksheet se�
                var worksheet = sheetName != null 
                    ? package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName)
                    : package.Workbook.Worksheets.FirstOrDefault();
                
                if (worksheet == null)
                {
                    var availableSheets = package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
                    var errorMessage = $"Sheet bulunamad�: {sheetName ?? "�lk sheet"}. Mevcut sheet'ler: {string.Join(", ", availableSheets)}";
                    _logger.LogError(errorMessage);
                    throw new Exception(errorMessage);
                }

                _logger.LogInformation("Worksheet se�ildi: {WorksheetName}, Dimensions: {Dimensions}", worksheet.Name, worksheet.Dimension?.Address);

                // Worksheet bo� mu kontrol et
                if (worksheet.Dimension == null)
                {
                    _logger.LogWarning("Worksheet bo�: {SheetName}", worksheet.Name);
                    return results;
                }

                // Header'lar� al
                var headers = new List<string>();
                var columnCount = worksheet.Dimension.Columns;
                for (int col = 1; col <= columnCount; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                    headers.Add(!string.IsNullOrEmpty(headerValue) ? headerValue : $"Column{col}");
                }

                _logger.LogInformation("Headers al�nd�: {HeaderCount} s�tun, Headers: [{Headers}]", 
                    headers.Count, string.Join(", ", headers));

                // Verileri oku ve veritaban�na kaydet
                var rowCount = worksheet.Dimension.Rows;
                var processedRowCount = 0;
                var skippedRowCount = 0;

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

                    // Sadece veri i�eren sat�rlar� kaydet
                    if (hasData)
                    {
                        try
                        {
                            var serializedRowData = JsonSerializer.Serialize(rowData);
                            
                            var dataRow = new ExcelDataRow
                            {
                                FileName = fileName,
                                SheetName = worksheet.Name,
                                RowIndex = row,
                                RowData = serializedRowData,
                                CreatedDate = DateTime.UtcNow,
                                Version = 1
                            };

                            _context.ExcelDataRows.Add(dataRow);

                            results.Add(new ExcelDataResponseDto
                            {
                                Id = 0, // Database'e kaydedildikten sonra g�ncellenir
                                FileName = fileName,
                                SheetName = worksheet.Name,
                                RowIndex = row,
                                Data = rowData,
                                CreatedDate = dataRow.CreatedDate,
                                Version = dataRow.Version
                            });

                            processedRowCount++;
                        }
                        catch (JsonException jsonEx)
                        {
                            _logger.LogError(jsonEx, "JSON serialization hatas� - Row: {Row}, Data: {@RowData}", row, rowData);
                            skippedRowCount++;
                        }
                    }
                    else
                    {
                        skippedRowCount++;
                    }
                }

                // Toplu kaydetme
                if (results.Any())
                {
                    await _context.SaveChangesAsync();

                    // ID'leri g�ncelle
                    var savedRows = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && r.SheetName == worksheet.Name)
                        .OrderBy(r => r.RowIndex)
                        .ToListAsync();

                    for (int i = 0; i < results.Count && i < savedRows.Count; i++)
                    {
                        results[i].Id = savedRows[i].Id;
                    }
                }

                _logger.LogInformation("Excel verisi ba�ar�yla okundu: {FileName}, Sheet: {SheetName}, ��lenen: {ProcessedCount}, Atlanan: {SkippedCount}", 
                    fileName, worksheet.Name, processedRowCount, skippedRowCount);

                return results;
            }
            catch (FileNotFoundException fnfEx)
            {
                _logger.LogError(fnfEx, "Dosya bulunamad� hatas�: {FileName}, SheetName: {SheetName}", fileName, sheetName);
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyas� okunurken hata: {FileName}, SheetName: {SheetName}", fileName, sheetName);
                throw;
            }
        }

        public async Task<List<ExcelDataResponseDto>> GetExcelDataAsync(string fileName, string? sheetName = null, int page = 1, int pageSize = 50)
        {
            try
            {
                _logger.LogInformation("GetExcelDataAsync �a�r�ld�: FileName={FileName}, SheetName={SheetName}, Page={Page}, PageSize={PageSize}", 
                    fileName, sheetName, page, pageSize);

                // Parametre validasyonu
                if (string.IsNullOrEmpty(fileName))
                {
                    _logger.LogWarning("GetExcelDataAsync: FileName bo�");
                    return new List<ExcelDataResponseDto>();
                }

                if (page < 1) page = 1;
                if (pageSize < 1 || pageSize > 1000) pageSize = 50;

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                    _logger.LogInformation("Sheet filtrelemesi uyguland�: {SheetName}", sheetName);
                }

                var totalCount = await query.CountAsync();
                _logger.LogInformation("Toplam kay�t say�s�: {TotalCount}", totalCount);

                var rows = await query
                    .OrderBy(r => r.RowIndex)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .ToListAsync();

                _logger.LogInformation("Sayfalama sonras� d�nen kay�t say�s�: {ReturnedCount}", rows.Count);

                var result = new List<ExcelDataResponseDto>();

                foreach (var row in rows)
                {
                    try
                    {
                        var deserializedData = JsonSerializer.Deserialize<Dictionary<string, string>>(row.RowData);
                        
                        result.Add(new ExcelDataResponseDto
                        {
                            Id = row.Id,
                            FileName = row.FileName,
                            SheetName = row.SheetName,
                            RowIndex = row.RowIndex,
                            Data = deserializedData ?? new Dictionary<string, string>(),
                            CreatedDate = row.CreatedDate,
                            ModifiedDate = row.ModifiedDate,
                            Version = row.Version,
                            ModifiedBy = row.ModifiedBy
                        });
                    }
                    catch (JsonException jsonEx)
                    {
                        _logger.LogError(jsonEx, "JSON deserialization hatas� - RowId: {RowId}, RowData: {RowData}", 
                            row.Id, row.RowData);
                        
                        // Hatal� JSON durumunda bo� dictionary ile devam et
                        result.Add(new ExcelDataResponseDto
                        {
                            Id = row.Id,
                            FileName = row.FileName,
                            SheetName = row.SheetName,
                            RowIndex = row.RowIndex,
                            Data = new Dictionary<string, string> { { "Error", "Veri format� hatal�" } },
                            CreatedDate = row.CreatedDate,
                            ModifiedDate = row.ModifiedDate,
                            Version = row.Version,
                            ModifiedBy = row.ModifiedBy
                        });
                    }
                }

                _logger.LogInformation("GetExcelDataAsync tamamland�: {ResultCount} kay�t d�nd�r�ld�", result.Count);
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
                _logger.LogInformation("GetAllExcelDataAsync �a�r�ld�: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

                // Parametre validasyonu
                if (string.IsNullOrEmpty(fileName))
                {
                    _logger.LogWarning("GetAllExcelDataAsync: FileName bo�");
                    return new List<ExcelDataResponseDto>();
                }

                // �nce dosyan�n var olup olmad���n� kontrol et
                var fileExists = await _context.ExcelFiles
                    .AnyAsync(f => f.FileName == fileName && f.IsActive);

                if (!fileExists)
                {
                    _logger.LogWarning("Dosya bulunamad� veya aktif de�il: {FileName}", fileName);
                    throw new FileNotFoundException($"Dosya bulunamad� veya aktif de�il: {fileName}");
                }

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                    _logger.LogInformation("Sheet filtrelemesi uyguland�: {SheetName}", sheetName);
                }

                var totalCount = await query.CountAsync();
                _logger.LogInformation("Toplam kay�t say�s�: {TotalCount}", totalCount);

                var rows = await query
                    .OrderBy(r => r.RowIndex)
                    .ToListAsync();

                var result = new List<ExcelDataResponseDto>();

                foreach (var row in rows)
                {
                    try
                    {
                        var deserializedData = JsonSerializer.Deserialize<Dictionary<string, string>>(row.RowData);
                        
                        result.Add(new ExcelDataResponseDto
                        {
                            Id = row.Id,
                            FileName = row.FileName,
                            SheetName = row.SheetName,
                            RowIndex = row.RowIndex,
                            Data = deserializedData ?? new Dictionary<string, string>(),
                            CreatedDate = row.CreatedDate,
                            ModifiedDate = row.ModifiedDate,
                            Version = row.Version,
                            ModifiedBy = row.ModifiedBy
                        });
                    }
                    catch (JsonException jsonEx)
                    {
                        _logger.LogError(jsonEx, "JSON deserialization hatas� - RowId: {RowId}, RowData: {RowData}", 
                            row.Id, row.RowData);
                        
                        // Hatal� JSON durumunda bo� dictionary ile devam et
                        result.Add(new ExcelDataResponseDto
                        {
                            Id = row.Id,
                            FileName = row.FileName,
                            SheetName = row.SheetName,
                            RowIndex = row.RowIndex,
                            Data = new Dictionary<string, string> { { "Error", "Veri format� hatal�" } },
                            CreatedDate = row.CreatedDate,
                            ModifiedDate = row.ModifiedDate,
                            Version = row.Version,
                            ModifiedBy = row.ModifiedBy
                        });
                    }
                }

                _logger.LogInformation("T�m Excel verileri getirildi: {FileName}, {Count} kay�t", fileName, result.Count);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "T�m Excel verileri getirilirken hata: {FileName}, {SheetName}", fileName, sheetName);
                throw;
            }
        }

        public async Task<ExcelDataResponseDto> UpdateExcelDataAsync(ExcelDataUpdateDto updateDto)
        {
            var dataRow = await _context.ExcelDataRows.FindAsync(updateDto.Id);
            if (dataRow == null)
                throw new Exception($"Veri bulunamad�: {updateDto.Id}");

            dataRow.RowData = JsonSerializer.Serialize(updateDto.Data);
            dataRow.ModifiedDate = DateTime.UtcNow;
            dataRow.ModifiedBy = updateDto.ModifiedBy;
            dataRow.Version++;

            await _context.SaveChangesAsync();

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
        }

        public async Task<List<ExcelDataResponseDto>> BulkUpdateExcelDataAsync(BulkUpdateDto bulkUpdateDto)
        {
            var results = new List<ExcelDataResponseDto>();

            foreach (var update in bulkUpdateDto.Updates)
            {
                var result = await UpdateExcelDataAsync(update);
                results.Add(result);
            }

            return results;
        }

        public async Task<bool> DeleteExcelDataAsync(int id, string? deletedBy = null)
        {
            var dataRow = await _context.ExcelDataRows.FindAsync(id);
            if (dataRow == null)
                return false;

            dataRow.IsDeleted = true;
            dataRow.ModifiedDate = DateTime.UtcNow;
            dataRow.ModifiedBy = deletedBy;

            await _context.SaveChangesAsync();
            return true;
        }

        public async Task<bool> DeleteExcelFileAsync(string fileName, string? deletedBy = null)
        {
            try
            {
                var excelFile = await _context.ExcelFiles
                    .FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
                
                if (excelFile == null)
                    return false;

                // Dosyay� inaktif olarak i�aretle (soft delete)
                excelFile.IsActive = false;

                // �lgili t�m verileri de soft delete yap
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

                // Fiziksel dosyay� da sil
                if (File.Exists(excelFile.FilePath))
                {
                    File.Delete(excelFile.FilePath);
                    _logger.LogInformation("Fiziksel dosya silindi: {FilePath}", excelFile.FilePath);
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyas� silinirken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<ExcelDataResponseDto> AddExcelRowAsync(string fileName, string sheetName, Dictionary<string, string> rowData, string? addedBy = null)
        {
            var maxRowIndex = await _context.ExcelDataRows
                .Where(r => r.FileName == fileName && r.SheetName == sheetName)
                .MaxAsync(r => (int?)r.RowIndex) ?? 0;

            var dataRow = new ExcelDataRow
            {
                FileName = fileName,
                SheetName = sheetName,
                RowIndex = maxRowIndex + 1,
                RowData = JsonSerializer.Serialize(rowData),
                CreatedDate = DateTime.UtcNow,
                ModifiedBy = addedBy,
                Version = 1
            };

            _context.ExcelDataRows.Add(dataRow);
            await _context.SaveChangesAsync();

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

        public async Task<byte[]> ExportToExcelAsync(ExcelExportRequestDto exportRequest)
        {
            try
            {
                _logger.LogInformation("Excel export ba�lat�ld�: {FileName}, Sheet: {Sheet}, RowIds: {RowIdCount}", 
                    exportRequest.FileName, exportRequest.SheetName, exportRequest.RowIds?.Count ?? 0);

                // Input validation
                if (string.IsNullOrEmpty(exportRequest.FileName))
                {
                    throw new ArgumentException("Dosya ad� gerekli");
                }

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == exportRequest.FileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(exportRequest.SheetName))
                    query = query.Where(r => r.SheetName == exportRequest.SheetName);

                if (exportRequest.RowIds?.Any() == true)
                    query = query.Where(r => exportRequest.RowIds.Contains(r.Id));

                var rows = await query.OrderBy(r => r.RowIndex).ToListAsync();

                _logger.LogInformation("Export i�in {Count} sat�r bulundu", rows.Count);

                if (!rows.Any())
                {
                    _logger.LogWarning("Export i�in veri bulunamad�");
                    // Bo� Excel dosyas� olu�tur
                    using var emptyPackage = new ExcelPackage();
                    var emptyWorksheet = emptyPackage.Workbook.Worksheets.Add(exportRequest.SheetName ?? "Export");
                    emptyWorksheet.Cells[1, 1].Value = "Veri bulunamad�";
                    return emptyPackage.GetAsByteArray();
                }

                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add(exportRequest.SheetName ?? "Export");

                // Header'lar� belirle - ilk sat�rdan
                var firstRowData = JsonSerializer.Deserialize<Dictionary<string, string>>(rows.First().RowData) ?? new Dictionary<string, string>();
                var headers = firstRowData.Keys.ToList();

                if (!headers.Any())
                {
                    _logger.LogWarning("Header bulunamad�, default header kullan�l�yor");
                    headers = new List<string> { "Data" };
                }

                _logger.LogInformation("Export header'lar�: {Headers}", string.Join(", ", headers));

                // Header'lar� yaz
                for (int i = 0; i < headers.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Verileri yaz
                int exportedRows = 0;
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    try
                    {
                        var rowData = JsonSerializer.Deserialize<Dictionary<string, string>>(rows[rowIndex].RowData) ?? new Dictionary<string, string>();
                        
                        for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                        {
                            var cellValue = rowData.ContainsKey(headers[colIndex]) ? rowData[headers[colIndex]] : "";
                            worksheet.Cells[rowIndex + 2, colIndex + 1].Value = cellValue;
                        }

                        // Modification history ekle (iste�e ba�l�)
                        if (exportRequest.IncludeModificationHistory && rows[rowIndex].ModifiedDate != null)
                        {
                            var historyColumn = headers.Count + 1;
                            if (rowIndex == 0)
                            {
                                worksheet.Cells[1, historyColumn].Value = "Son De�i�iklik";
                                worksheet.Cells[1, historyColumn].Style.Font.Bold = true;
                                worksheet.Cells[1, historyColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Cells[1, historyColumn].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                            }
                            worksheet.Cells[rowIndex + 2, historyColumn].Value = $"{rows[rowIndex].ModifiedDate:yyyy-MM-dd HH:mm} ({rows[rowIndex].ModifiedBy ?? "Bilinmiyor"})";
                        }

                        exportedRows++;
                    }
                    catch (JsonException jsonEx)
                    {
                        _logger.LogWarning(jsonEx, "Export s�ras�nda JSON parsing hatas�: Row {RowIndex}", rowIndex);
                        // Hatal� sat�r� atla, di�erlerine devam et
                        worksheet.Cells[rowIndex + 2, 1].Value = $"[Veri format� hatas�: Row {rows[rowIndex].RowIndex}]";
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Export s�ras�nda sat�r hatas�: Row {RowIndex}", rowIndex);
                        worksheet.Cells[rowIndex + 2, 1].Value = $"[Hata: Row {rows[rowIndex].RowIndex}]";
                    }
                }

                // Kolonlar� otomatik boyutland�r
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Freeze header row
                worksheet.View.FreezePanes(2, 1);

                // Borders ekle
                var range = worksheet.Cells[1, 1, rows.Count + 1, headers.Count + (exportRequest.IncludeModificationHistory ? 1 : 0)];
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                var result = package.GetAsByteArray();

                _logger.LogInformation("Excel export tamamland�: {ExportedRows} sat�r, {FileSize} bytes", exportedRows, result.Length);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel export hatas�: {FileName}", exportRequest.FileName);
                throw;
            }
        }

        public async Task<List<string>> GetSheetsAsync(string fileName)
        {
            try
            {
                _logger.LogInformation("Sheets al�n�yor: {FileName}", fileName);

                // Input validation
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("Dosya ad� gerekli");
                }

                var excelFile = await _context.ExcelFiles.FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
                if (excelFile == null)
                {
                    throw new FileNotFoundException($"Dosya bulunamad�: {fileName}");
                }

                if (!File.Exists(excelFile.FilePath))
                {
                    _logger.LogWarning("Fiziksel dosya bulunamad�, veritaban�ndan sheet'ler al�n�yor: {FilePath}", excelFile.FilePath);
                    
                    // Fiziksel dosya yoksa veritaban�ndan sheet'leri al
                    var dbSheets = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .Select(r => r.SheetName)
                        .Distinct()
                        .ToListAsync();

                    return dbSheets;
                }

                var sheets = new List<string>();
                using var package = new ExcelPackage(new FileInfo(excelFile.FilePath));
                
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    sheets.Add(worksheet.Name);
                }

                _logger.LogInformation("Sheets al�nd�: {Count} sheet - [{Sheets}]", sheets.Count, string.Join(", ", sheets));

                return sheets;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sheets al�n�rken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<object> GetDataStatisticsAsync(string fileName, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("Data statistics al�n�yor: {FileName}, Sheet: {Sheet}", fileName, sheetName);

                // Input validation
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("Dosya ad� gerekli");
                }

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                var totalRows = await query.CountAsync();
                var modifiedRows = await query.Where(r => r.ModifiedDate != null).CountAsync();
                var sheetsCount = await query.Select(r => r.SheetName).Distinct().CountAsync();
                var lastModified = await query.Where(r => r.ModifiedDate != null).MaxAsync(r => (DateTime?)r.ModifiedDate);
                var firstCreated = await query.MinAsync(r => (DateTime?)r.CreatedDate);

                // Sheet baz�nda istatistikler
                var sheetStats = await query
                    .GroupBy(r => r.SheetName)
                    .Select(g => new
                    {
                        SheetName = g.Key,
                        RowCount = g.Count(),
                        ModifiedRowCount = g.Count(r => r.ModifiedDate != null),
                        LastModified = g.Where(r => r.ModifiedDate != null).Max(r => (DateTime?)r.ModifiedDate),
                        FirstCreated = g.Min(r => r.CreatedDate)
                    })
                    .ToListAsync();

                var statistics = new
                {
                    fileName = fileName,
                    sheetName = sheetName,
                    totalRows = totalRows,
                    modifiedRows = modifiedRows,
                    sheetsCount = sheetsCount,
                    lastModified = lastModified,
                    firstCreated = firstCreated,
                    hasData = totalRows > 0,
                    modificationRate = totalRows > 0 ? (double)modifiedRows / totalRows * 100 : 0,
                    sheetStatistics = sheetStats
                };

                _logger.LogInformation("Data statistics al�nd�: {TotalRows} rows, {ModifiedRows} modified, {SheetsCount} sheets", 
                    totalRows, modifiedRows, sheetsCount);

                return statistics;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Data statistics al�n�rken hata: {FileName}", fileName);
                throw;
            }
        }
    }
}