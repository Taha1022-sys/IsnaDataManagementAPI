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

        public async Task<List<ExcelDataResponseDto>> ReadExcelDataAsync(string fileName, string? sheetName = null)
        {
            try
            {
                var excelFile = await _context.ExcelFiles.FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);
                if (excelFile == null)
                    throw new FileNotFoundException($"Dosya bulunamadý: {fileName}");

                var results = new List<ExcelDataResponseDto>();

                using var package = new ExcelPackage(new FileInfo(excelFile.FilePath));
                var worksheet = sheetName != null ? package.Workbook.Worksheets[sheetName] : package.Workbook.Worksheets.FirstOrDefault();
                
                if (worksheet == null)
                    throw new Exception($"Sheet bulunamadý: {sheetName ?? "Ýlk sheet"}");

                // Header'larý al
                var headers = new List<string>();
                var columnCount = worksheet.Dimension?.Columns ?? 0;
                for (int col = 1; col <= columnCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}");
                }

                // Verileri oku ve veritabanýna kaydet
                var rowCount = worksheet.Dimension?.Rows ?? 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, object>();
                    for (int col = 1; col <= headers.Count; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        rowData[headers[col - 1]] = cellValue ?? "";
                    }

                    var dataRow = new ExcelDataRow
                    {
                        FileName = fileName,
                        SheetName = worksheet.Name,
                        RowIndex = row,
                        RowData = JsonSerializer.Serialize(rowData),
                        CreatedDate = DateTime.UtcNow,
                        Version = 1
                    };

                    _context.ExcelDataRows.Add(dataRow);

                    results.Add(new ExcelDataResponseDto
                    {
                        Id = dataRow.Id,
                        FileName = fileName,
                        SheetName = worksheet.Name,
                        RowIndex = row,
                        Data = rowData,
                        CreatedDate = dataRow.CreatedDate,
                        Version = dataRow.Version
                    });
                }

                await _context.SaveChangesAsync();
                return results;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyasý okunurken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<List<ExcelDataResponseDto>> GetExcelDataAsync(string fileName, string? sheetName = null, int page = 1, int pageSize = 50)
        {
            var query = _context.ExcelDataRows
                .Where(r => r.FileName == fileName && !r.IsDeleted);

            if (!string.IsNullOrEmpty(sheetName))
                query = query.Where(r => r.SheetName == sheetName);

            var rows = await query
                .OrderBy(r => r.RowIndex)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            return rows.Select(r => new ExcelDataResponseDto
            {
                Id = r.Id,
                FileName = r.FileName,
                SheetName = r.SheetName,
                RowIndex = r.RowIndex,
                Data = JsonSerializer.Deserialize<Dictionary<string, object>>(r.RowData) ?? new Dictionary<string, object>(),
                CreatedDate = r.CreatedDate,
                ModifiedDate = r.ModifiedDate,
                Version = r.Version,
                ModifiedBy = r.ModifiedBy
            }).ToList();
        }

        public async Task<ExcelDataResponseDto> UpdateExcelDataAsync(ExcelDataUpdateDto updateDto)
        {
            var dataRow = await _context.ExcelDataRows.FindAsync(updateDto.Id);
            if (dataRow == null)
                throw new Exception($"Veri bulunamadý: {updateDto.Id}");

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

        public async Task<ExcelDataResponseDto> AddExcelRowAsync(string fileName, string sheetName, Dictionary<string, object> rowData, string? addedBy = null)
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
                // Header'larý ekle
                var firstRowData = JsonSerializer.Deserialize<Dictionary<string, object>>(rows.First().RowData) ?? new Dictionary<string, object>();
                var headers = firstRowData.Keys.ToList();

                for (int i = 0; i < headers.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }

                // Verileri ekle
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, object>>(rows[rowIndex].RowData) ?? new Dictionary<string, object>();
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