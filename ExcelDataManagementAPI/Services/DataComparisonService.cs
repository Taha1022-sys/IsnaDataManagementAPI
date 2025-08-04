using Microsoft.EntityFrameworkCore;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Models.DTOs;
using System.Text.Json;

namespace ExcelDataManagementAPI.Services
{
    public class DataComparisonService : IDataComparisonService
    {
        private readonly ExcelDataContext _context;
        private readonly ILogger<DataComparisonService> _logger;

        public DataComparisonService(ExcelDataContext context, ILogger<DataComparisonService> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task<DataComparisonResultDto> CompareFilesAsync(string fileName1, string fileName2, string? sheetName = null)
        {
            try
            {
                var file1Data = await GetFileDataAsync(fileName1, sheetName);
                var file2Data = await GetFileDataAsync(fileName2, sheetName);

                var differences = new List<DataDifferenceDto>();
                var comparisonId = Guid.NewGuid().ToString();

                // File1'deki verilerle File2'deki verileri karþýlaþtýr
                foreach (var row1 in file1Data)
                {
                    var correspondingRow2 = file2Data.FirstOrDefault(r => r.RowIndex == row1.RowIndex);
                    if (correspondingRow2 == null)
                    {
                        // Satýr File2'de yok - silinmiþ
                        differences.Add(new DataDifferenceDto
                        {
                            RowIndex = row1.RowIndex,
                            ColumnName = "EntireRow",
                            OldValue = row1.RowData,
                            NewValue = null,
                            Type = DifferenceType.Deleted
                        });
                    }
                    else
                    {
                        // Satýrlarý karþýlaþtýr
                        var row1Data = JsonSerializer.Deserialize<Dictionary<string, string>>(row1.RowData) ?? new Dictionary<string, string>();
                        var row2Data = JsonSerializer.Deserialize<Dictionary<string, string>>(correspondingRow2.RowData) ?? new Dictionary<string, string>();

                        foreach (var kvp in row1Data)
                        {
                            if (!row2Data.ContainsKey(kvp.Key) || kvp.Value != row2Data[kvp.Key])
                            {
                                differences.Add(new DataDifferenceDto
                                {
                                    RowIndex = row1.RowIndex,
                                    ColumnName = kvp.Key,
                                    OldValue = kvp.Value,
                                    NewValue = row2Data.ContainsKey(kvp.Key) ? row2Data[kvp.Key] : null,
                                    Type = DifferenceType.Modified
                                });
                            }
                        }
                    }
                }

                // File2'de olup File1'de olmayan satýrlarý bul (yeni eklenenler)
                foreach (var row2 in file2Data)
                {
                    if (!file1Data.Any(r => r.RowIndex == row2.RowIndex))
                    {
                        differences.Add(new DataDifferenceDto
                        {
                            RowIndex = row2.RowIndex,
                            ColumnName = "EntireRow",
                            OldValue = null,
                            NewValue = row2.RowData,
                            Type = DifferenceType.Added
                        });
                    }
                }

                var summary = new ComparisonSummaryDto
                {
                    TotalRows = Math.Max(file1Data.Count, file2Data.Count),
                    ModifiedRows = differences.Count(d => d.Type == DifferenceType.Modified),
                    AddedRows = differences.Count(d => d.Type == DifferenceType.Added),
                    DeletedRows = differences.Count(d => d.Type == DifferenceType.Deleted),
                    UnchangedRows = Math.Max(file1Data.Count, file2Data.Count) - differences.Select(d => d.RowIndex).Distinct().Count()
                };

                return new DataComparisonResultDto
                {
                    ComparisonId = comparisonId,
                    File1Name = fileName1,
                    File2Name = fileName2,
                    ComparisonDate = DateTime.UtcNow,
                    Differences = differences,
                    Summary = summary
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya karþýlaþtýrma hatasý: {File1} vs {File2}", fileName1, fileName2);
                throw;
            }
        }

        public async Task<DataComparisonResultDto> CompareVersionsAsync(string fileName, DateTime version1Date, DateTime version2Date, string? sheetName = null)
        {
            try
            {
                var version1Data = await GetFileDataByDateAsync(fileName, version1Date, sheetName);
                var version2Data = await GetFileDataByDateAsync(fileName, version2Date, sheetName);

                // Ayný mantýkla karþýlaþtýr
                var differences = new List<DataDifferenceDto>();
                var comparisonId = Guid.NewGuid().ToString();

                // Basit implementasyon - daha detaylý yapýlabilir
                foreach (var row1 in version1Data)
                {
                    var correspondingRow2 = version2Data.FirstOrDefault(r => r.RowIndex == row1.RowIndex);
                    if (correspondingRow2 != null && row1.RowData != correspondingRow2.RowData)
                    {
                        differences.Add(new DataDifferenceDto
                        {
                            RowIndex = row1.RowIndex,
                            ColumnName = "RowData",
                            OldValue = row1.RowData,
                            NewValue = correspondingRow2.RowData,
                            Type = DifferenceType.Modified
                        });
                    }
                }

                var summary = new ComparisonSummaryDto
                {
                    TotalRows = Math.Max(version1Data.Count, version2Data.Count),
                    ModifiedRows = differences.Count,
                    AddedRows = version2Data.Count - version1Data.Count > 0 ? version2Data.Count - version1Data.Count : 0,
                    DeletedRows = version1Data.Count - version2Data.Count > 0 ? version1Data.Count - version2Data.Count : 0,
                    UnchangedRows = Math.Min(version1Data.Count, version2Data.Count) - differences.Count
                };

                return new DataComparisonResultDto
                {
                    ComparisonId = comparisonId,
                    File1Name = $"{fileName} ({version1Date:yyyy-MM-dd})",
                    File2Name = $"{fileName} ({version2Date:yyyy-MM-dd})",
                    ComparisonDate = DateTime.UtcNow,
                    Differences = differences,
                    Summary = summary
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Versiyon karþýlaþtýrma hatasý: {FileName}", fileName);
                throw;
            }
        }

        public async Task<List<DataDifferenceDto>> GetChangesAsync(string fileName, DateTime? fromDate = null, DateTime? toDate = null, string? sheetName = null)
        {
            try
            {
                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && r.ModifiedDate != null);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                if (fromDate.HasValue)
                    query = query.Where(r => r.ModifiedDate >= fromDate.Value);

                if (toDate.HasValue)
                    query = query.Where(r => r.ModifiedDate <= toDate.Value);

                var changes = await query
                    .OrderByDescending(r => r.ModifiedDate)
                    .ToListAsync();

                return changes.Select(c => new DataDifferenceDto
                {
                    RowIndex = c.RowIndex,
                    ColumnName = "RowData",
                    NewValue = c.RowData,
                    Type = DifferenceType.Modified
                }).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Deðiþiklik geçmiþi alýnýrken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<List<object>> GetChangeHistoryAsync(string fileName, string? sheetName = null)
        {
            try
            {
                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && r.ModifiedDate != null);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                var history = await query
                    .OrderByDescending(r => r.ModifiedDate)
                    .Select(r => new
                    {
                        r.Id,
                        r.RowIndex,
                        r.SheetName,
                        r.ModifiedDate,
                        r.ModifiedBy,
                        r.Version
                    })
                    .ToListAsync();

                return history.Cast<object>().ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya geçmiþi alýnýrken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<List<object>> GetRowHistoryAsync(int rowId)
        {
            try
            {
                var baseRow = await _context.ExcelDataRows.FindAsync(rowId);
                if (baseRow == null)
                    throw new Exception($"Satýr bulunamadý: {rowId}");
                
                // Ayný dosya, sheet ve satýr indexindeki tüm versiyonlarý al
                var history = await _context.ExcelDataRows
                    .Where(r => r.FileName == baseRow.FileName && 
                               r.SheetName == baseRow.SheetName && 
                               r.RowIndex == baseRow.RowIndex)
                    .OrderByDescending(r => r.Version)
                    .Select(r => new
                    {
                        r.Id,
                        r.RowData,
                        r.CreatedDate,
                        r.ModifiedDate,
                        r.ModifiedBy,
                        r.Version
                    })
                    .ToListAsync();

                return history.Cast<object>().ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Satýr geçmiþi alýnýrken hata: {RowId}", rowId);
                throw;
            }
        }

        private async Task<List<Models.ExcelDataRow>> GetFileDataAsync(string fileName, string? sheetName = null)
        {
            var query = _context.ExcelDataRows
                .Where(r => r.FileName == fileName && !r.IsDeleted);

            if (!string.IsNullOrEmpty(sheetName))
                query = query.Where(r => r.SheetName == sheetName);

            return await query.OrderBy(r => r.RowIndex).ToListAsync();
        }

        private async Task<List<Models.ExcelDataRow>> GetFileDataByDateAsync(string fileName, DateTime date, string? sheetName = null)
        {
            var query = _context.ExcelDataRows
                .Where(r => r.FileName == fileName && !r.IsDeleted && 
                           (r.CreatedDate <= date && (r.ModifiedDate == null || r.ModifiedDate <= date)));

            if (!string.IsNullOrEmpty(sheetName))
                query = query.Where(r => r.SheetName == sheetName);

            return await query.OrderBy(r => r.RowIndex).ToListAsync();
        }
    }
}