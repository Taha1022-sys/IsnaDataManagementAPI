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
                _logger.LogInformation("Dosya kar��la�t�rmas� ba�lat�ld�: {File1} vs {File2}, Sheet: {Sheet}", fileName1, fileName2, sheetName);

                // Input validation
                if (string.IsNullOrEmpty(fileName1) || string.IsNullOrEmpty(fileName2))
                {
                    throw new ArgumentException("Dosya adlar� bo� olamaz");
                }

                var file1Data = await GetFileDataAsync(fileName1, sheetName);
                var file2Data = await GetFileDataAsync(fileName2, sheetName);

                _logger.LogInformation("Dosya verileri al�nd�: File1={Count1} rows, File2={Count2} rows", file1Data.Count, file2Data.Count);

                var differences = new List<DataDifferenceDto>();
                var comparisonId = Guid.NewGuid().ToString();

                // File1'deki verilerle File2'deki verileri kar��la�t�r
                foreach (var row1 in file1Data)
                {
                    try
                    {
                        var correspondingRow2 = file2Data.FirstOrDefault(r => r.RowIndex == row1.RowIndex);
                        if (correspondingRow2 == null)
                        {
                            // Sat�r File2'de yok - silinmi�
                            differences.Add(new DataDifferenceDto
                            {
                                RowIndex = row1.RowIndex,
                                ColumnName = "EntireRow",
                                OldValue = GetSafeJsonPreview(row1.RowData),
                                NewValue = null,
                                Type = DifferenceType.Deleted
                            });
                        }
                        else
                        {
                            // Sat�rlar� kar��la�t�r
                            var row1Data = SafeDeserializeJson(row1.RowData);
                            var row2Data = SafeDeserializeJson(correspondingRow2.RowData);

                            CompareRowData(row1Data, row2Data, row1.RowIndex, differences);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Sat�r kar��la�t�rmas�nda hata: RowIndex={RowIndex}", row1.RowIndex);
                        // Hatal� sat�rlar� atla, kar��la�t�rmaya devam et
                    }
                }

                // File2'de olup File1'de olmayan sat�rlar� bul (yeni eklenenler)
                foreach (var row2 in file2Data)
                {
                    try
                    {
                        if (!file1Data.Any(r => r.RowIndex == row2.RowIndex))
                        {
                            differences.Add(new DataDifferenceDto
                            {
                                RowIndex = row2.RowIndex,
                                ColumnName = "EntireRow",
                                OldValue = null,
                                NewValue = GetSafeJsonPreview(row2.RowData),
                                Type = DifferenceType.Added
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Yeni sat�r kontrol�nde hata: RowIndex={RowIndex}", row2.RowIndex);
                    }
                }

                var summary = CreateComparisonSummary(file1Data, file2Data, differences);

                _logger.LogInformation("Kar��la�t�rma tamamland�: {DifferenceCount} fark, {AddedRows} eklenen, {DeletedRows} silinen, {ModifiedRows} de�i�en", 
                    differences.Count, summary.AddedRows, summary.DeletedRows, summary.ModifiedRows);

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
                _logger.LogError(ex, "Dosya kar��la�t�rma hatas�: {File1} vs {File2}", fileName1, fileName2);
                throw;
            }
        }

        public async Task<DataComparisonResultDto> CompareVersionsAsync(string fileName, DateTime version1Date, DateTime version2Date, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("Versiyon kar��la�t�rmas� ba�lat�ld�: {FileName}, {Date1} vs {Date2}", fileName, version1Date, version2Date);

                // Input validation
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("Dosya ad� bo� olamaz");
                }

                if (version1Date >= version2Date)
                {
                    throw new ArgumentException("�kinci versiyon tarihi birinci versiyon tarihinden sonra olmal�");
                }

                var version1Data = await GetFileDataByDateAsync(fileName, version1Date, sheetName);
                var version2Data = await GetFileDataByDateAsync(fileName, version2Date, sheetName);

                _logger.LogInformation("Versiyon verileri al�nd�: Version1={Count1} rows, Version2={Count2} rows", version1Data.Count, version2Data.Count);

                var differences = new List<DataDifferenceDto>();
                var comparisonId = Guid.NewGuid().ToString();

                // Versiyon kar��la�t�rmas� - daha detayl�
                foreach (var row1 in version1Data)
                {
                    try
                    {
                        var correspondingRow2 = version2Data.FirstOrDefault(r => r.RowIndex == row1.RowIndex);
                        if (correspondingRow2 == null)
                        {
                            // Sat�r ikinci versiyonda yok - silinmi�
                            differences.Add(new DataDifferenceDto
                            {
                                RowIndex = row1.RowIndex,
                                ColumnName = "EntireRow",
                                OldValue = GetSafeJsonPreview(row1.RowData),
                                NewValue = null,
                                Type = DifferenceType.Deleted
                            });
                        }
                        else if (row1.RowData != correspondingRow2.RowData)
                        {
                            // Sat�r de�i�mi� - detayl� kar��la�t�r
                            var row1Data = SafeDeserializeJson(row1.RowData);
                            var row2Data = SafeDeserializeJson(correspondingRow2.RowData);

                            CompareRowData(row1Data, row2Data, row1.RowIndex, differences);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Versiyon sat�r kar��la�t�rmas�nda hata: RowIndex={RowIndex}", row1.RowIndex);
                    }
                }

                // Version2'de olup Version1'de olmayan sat�rlar� bul
                foreach (var row2 in version2Data)
                {
                    try
                    {
                        if (!version1Data.Any(r => r.RowIndex == row2.RowIndex))
                        {
                            differences.Add(new DataDifferenceDto
                            {
                                RowIndex = row2.RowIndex,
                                ColumnName = "EntireRow",
                                OldValue = null,
                                NewValue = GetSafeJsonPreview(row2.RowData),
                                Type = DifferenceType.Added
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Versiyon yeni sat�r kontrol�nde hata: RowIndex={RowIndex}", row2.RowIndex);
                    }
                }

                var summary = CreateComparisonSummary(version1Data, version2Data, differences);

                _logger.LogInformation("Versiyon kar��la�t�rmas� tamamland�: {DifferenceCount} fark bulundu", differences.Count);

                return new DataComparisonResultDto
                {
                    ComparisonId = comparisonId,
                    File1Name = $"{fileName} ({version1Date:yyyy-MM-dd HH:mm})",
                    File2Name = $"{fileName} ({version2Date:yyyy-MM-dd HH:mm})",
                    ComparisonDate = DateTime.UtcNow,
                    Differences = differences,
                    Summary = summary
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Versiyon kar��la�t�rma hatas�: {FileName}", fileName);
                throw;
            }
        }

        public async Task<List<DataDifferenceDto>> GetChangesAsync(string fileName, DateTime? fromDate = null, DateTime? toDate = null, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("De�i�iklik listesi istendi: {FileName}, {FromDate} - {ToDate}, Sheet: {Sheet}", fileName, fromDate, toDate, sheetName);

                // Input validation
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("Dosya ad� bo� olamaz");
                }

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && r.ModifiedDate != null && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                if (fromDate.HasValue)
                    query = query.Where(r => r.ModifiedDate >= fromDate.Value);

                if (toDate.HasValue)
                    query = query.Where(r => r.ModifiedDate <= toDate.Value);

                var changes = await query
                    .OrderByDescending(r => r.ModifiedDate)
                    .ToListAsync();

                _logger.LogInformation("De�i�iklik listesi al�nd�: {Count} kay�t", changes.Count);

                return changes.Select(c => new DataDifferenceDto
                {
                    RowIndex = c.RowIndex,
                    ColumnName = "RowData",
                    NewValue = GetSafeJsonPreview(c.RowData),
                    OldValue = null, // Eski veri burada mevcut de�il
                    Type = DifferenceType.Modified
                }).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "De�i�iklik ge�mi�i al�n�rken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<List<object>> GetChangeHistoryAsync(string fileName, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("De�i�iklik ge�mi�i istendi: {FileName}, Sheet: {Sheet}", fileName, sheetName);

                // Input validation
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("Dosya ad� bo� olamaz");
                }

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                var history = await query
                    .OrderByDescending(r => r.ModifiedDate ?? r.CreatedDate)
                    .Select(r => new
                    {
                        r.Id,
                        r.RowIndex,
                        r.SheetName,
                        r.CreatedDate,
                        r.ModifiedDate,
                        r.ModifiedBy,
                        r.Version,
                        HasModification = r.ModifiedDate != null,
                        LastChangeDate = r.ModifiedDate ?? r.CreatedDate,
                        DataPreview = r.RowData.Substring(0, Math.Min(100, r.RowData.Length)) + (r.RowData.Length > 100 ? "..." : "")
                    })
                    .ToListAsync();

                _logger.LogInformation("De�i�iklik ge�mi�i al�nd�: {Count} kay�t", history.Count);

                return history.Cast<object>().ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya ge�mi�i al�n�rken hata: {FileName}", fileName);
                throw;
            }
        }

        public async Task<List<object>> GetRowHistoryAsync(int rowId)
        {
            try
            {
                _logger.LogInformation("Sat�r ge�mi�i istendi: RowId={RowId}", rowId);

                // Input validation
                if (rowId <= 0)
                {
                    throw new ArgumentException("Ge�erli bir sat�r ID'si gerekli");
                }

                var baseRow = await _context.ExcelDataRows.FindAsync(rowId);
                if (baseRow == null)
                {
                    throw new Exception($"Sat�r bulunamad�: {rowId}");
                }
                
                // Ayn� dosya, sheet ve sat�r indexindeki t�m versiyonlar� al
                var history = await _context.ExcelDataRows
                    .Where(r => r.FileName == baseRow.FileName && 
                               r.SheetName == baseRow.SheetName && 
                               r.RowIndex == baseRow.RowIndex &&
                               !r.IsDeleted)
                    .OrderByDescending(r => r.Version)
                    .ThenByDescending(r => r.ModifiedDate ?? r.CreatedDate)
                    .Select(r => new
                    {
                        r.Id,
                        r.RowIndex,
                        r.SheetName,
                        r.FileName,
                        DataPreview = r.RowData.Substring(0, Math.Min(200, r.RowData.Length)) + (r.RowData.Length > 200 ? "..." : ""),
                        r.CreatedDate,
                        r.ModifiedDate,
                        r.ModifiedBy,
                        r.Version,
                        IsCurrentVersion = r.Id == rowId,
                        ChangeDate = r.ModifiedDate ?? r.CreatedDate
                    })
                    .ToListAsync();

                _logger.LogInformation("Sat�r ge�mi�i al�nd�: {Count} versiyon", history.Count);

                return history.Cast<object>().ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sat�r ge�mi�i al�n�rken hata: {RowId}", rowId);
                throw;
            }
        }

        private async Task<List<Models.ExcelDataRow>> GetFileDataAsync(string fileName, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("Dosya verisi al�n�yor: {FileName}, Sheet: {Sheet}", fileName, sheetName);

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                var data = await query.OrderBy(r => r.RowIndex).ToListAsync();

                _logger.LogInformation("Dosya verisi al�nd�: {Count} sat�r", data.Count);

                return data;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya verisi al�n�rken hata: {FileName}", fileName);
                throw;
            }
        }

        private async Task<List<Models.ExcelDataRow>> GetFileDataByDateAsync(string fileName, DateTime date, string? sheetName = null)
        {
            try
            {
                _logger.LogInformation("Tarihe g�re dosya verisi al�n�yor: {FileName}, Date: {Date}, Sheet: {Sheet}", fileName, date, sheetName);

                var query = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted && 
                               r.CreatedDate <= date && 
                               (r.ModifiedDate == null || r.ModifiedDate <= date));

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                var data = await query.OrderBy(r => r.RowIndex).ToListAsync();

                _logger.LogInformation("Tarihe g�re dosya verisi al�nd�: {Count} sat�r", data.Count);

                return data;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Tarihe g�re dosya verisi al�n�rken hata: {FileName}, Date: {Date}", fileName, date);
                throw;
            }
        }

        // Helper methods
        private Dictionary<string, string> SafeDeserializeJson(string jsonData)
        {
            try
            {
                if (string.IsNullOrEmpty(jsonData))
                {
                    return new Dictionary<string, string>();
                }

                var result = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonData);
                return result ?? new Dictionary<string, string>();
            }
            catch (JsonException ex)
            {
                _logger.LogWarning(ex, "JSON deserialization hatas�, bo� dictionary d�nd�r�l�yor: {JsonData}", jsonData?.Substring(0, Math.Min(100, jsonData.Length)));
                return new Dictionary<string, string> { { "Error", "JSON format hatas�" } };
            }
        }

        private string GetSafeJsonPreview(string jsonData)
        {
            if (string.IsNullOrEmpty(jsonData))
                return "";

            if (jsonData.Length <= 200)
                return jsonData;

            return jsonData.Substring(0, 200) + "...";
        }

        private void CompareRowData(Dictionary<string, string> row1Data, Dictionary<string, string> row2Data, int rowIndex, List<DataDifferenceDto> differences)
        {
            // T�m column'lar� kar��la�t�r
            var allColumns = row1Data.Keys.Union(row2Data.Keys);

            foreach (var column in allColumns)
            {
                var value1 = row1Data.ContainsKey(column) ? row1Data[column] : null;
                var value2 = row2Data.ContainsKey(column) ? row2Data[column] : null;

                if (value1 != value2)
                {
                    differences.Add(new DataDifferenceDto
                    {
                        RowIndex = rowIndex,
                        ColumnName = column,
                        OldValue = value1,
                        NewValue = value2,
                        Type = DifferenceType.Modified
                    });
                }
            }
        }

        private ComparisonSummaryDto CreateComparisonSummary(List<Models.ExcelDataRow> file1Data, List<Models.ExcelDataRow> file2Data, List<DataDifferenceDto> differences)
        {
            var addedRows = differences.Count(d => d.Type == DifferenceType.Added && d.ColumnName == "EntireRow");
            var deletedRows = differences.Count(d => d.Type == DifferenceType.Deleted && d.ColumnName == "EntireRow");
            var modifiedRowsCount = differences.Where(d => d.Type == DifferenceType.Modified).Select(d => d.RowIndex).Distinct().Count();

            var totalRows = Math.Max(file1Data.Count, file2Data.Count);
            var unchangedRows = totalRows - addedRows - deletedRows - modifiedRowsCount;

            return new ComparisonSummaryDto
            {
                TotalRows = totalRows,
                ModifiedRows = modifiedRowsCount,
                AddedRows = addedRows,
                DeletedRows = deletedRows,
                UnchangedRows = Math.Max(0, unchangedRows)
            };
        }
    }
}