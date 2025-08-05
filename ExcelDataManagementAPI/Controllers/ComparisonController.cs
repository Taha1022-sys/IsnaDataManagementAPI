using Microsoft.AspNetCore.Mvc;
using ExcelDataManagementAPI.Services;
using ExcelDataManagementAPI.Models.DTOs;

namespace ExcelDataManagementAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ComparisonController : ControllerBase
    {
        private readonly IDataComparisonService _comparisonService;
        private readonly IExcelService _excelService;
        private readonly ILogger<ComparisonController> _logger;

        public ComparisonController(IDataComparisonService comparisonService, IExcelService excelService, ILogger<ComparisonController> logger)
        {
            _comparisonService = comparisonService;
            _excelService = excelService;
            _logger = logger;
        }

        /// <summary>
        /// API test endpoint'i
        /// </summary>
        [HttpGet("test")]
        public IActionResult Test()
        {
            return Ok(new { 
                message = "Comparison API çalýþýyor!", 
                timestamp = DateTime.Now,
                availableOperations = new[]
                {
                    "Dosya karþýlaþtýrma: POST /api/comparison/files",
                    "Manuel dosya karþýlaþtýrma: POST /api/comparison/compare-from-files",
                    "Versiyon karþýlaþtýrma: POST /api/comparison/versions",
                    "Deðiþiklik geçmiþi: GET /api/comparison/history/{fileName}"
                }
            });
        }

        /// <summary>
        /// Bilgisayardan iki Excel dosyasý seçip karþýlaþtýrma - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPost("compare-from-files")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> CompareFromFiles([FromForm] CompareExcelFilesDto request)
        {
            try
            {
                // Input validation
                if (request == null)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Karþýlaþtýrma isteði boþ olamaz" 
                    });
                }

                if (request.File1 == null || request.File1.Length == 0)
                    return BadRequest(new { 
                        success = false, 
                        message = "Lütfen birinci Excel dosyasýný seçin" 
                    });

                if (request.File2 == null || request.File2.Length == 0)
                    return BadRequest(new { 
                        success = false, 
                        message = "Lütfen ikinci Excel dosyasýný seçin" 
                    });

                // Dosya uzantýsý kontrolü
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var file1Extension = Path.GetExtension(request.File1.FileName).ToLowerInvariant();
                var file2Extension = Path.GetExtension(request.File2.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(file1Extension) || !allowedExtensions.Contains(file2Extension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalarý (.xlsx, .xls) desteklenir" 
                    });
                }

                // SheetName'leri normalize et - "undefined" ise null yap
                var sheet1Name = NormalizeSheetName(request.Sheet1Name);
                var sheet2Name = NormalizeSheetName(request.Sheet2Name);

                _logger.LogInformation("Excel dosya karþýlaþtýrmasý baþlatýldý: {File1} vs {File2}", 
                    request.File1.FileName, request.File2.FileName);

                // Her iki dosyayý da yükle
                var uploadedFile1 = await _excelService.UploadExcelFileAsync(request.File1, request.ComparedBy);
                var uploadedFile2 = await _excelService.UploadExcelFileAsync(request.File2, request.ComparedBy);

                _logger.LogInformation("Dosyalar yüklendi: {File1} -> {InternalName1}, {File2} -> {InternalName2}", 
                    request.File1.FileName, uploadedFile1.FileName, request.File2.FileName, uploadedFile2.FileName);

                // Her iki dosyayý da oku
                var file1Data = await _excelService.ReadExcelDataAsync(uploadedFile1.FileName, sheet1Name);
                var file2Data = await _excelService.ReadExcelDataAsync(uploadedFile2.FileName, sheet2Name);

                _logger.LogInformation("Dosya verileri okundu: File1={RowCount1} rows, File2={RowCount2} rows", 
                    file1Data.Count, file2Data.Count);

                // Karþýlaþtýr
                var result = await _comparisonService.CompareFilesAsync(
                    uploadedFile1.FileName, 
                    uploadedFile2.FileName, 
                    sheet1Name ?? sheet2Name);

                _logger.LogInformation("Karþýlaþtýrma tamamlandý: {DifferenceCount} fark bulundu", 
                    result.Differences?.Count ?? 0);

                // Improved response format
                var response = new
                {
                    success = true,
                    data = result,
                    fileInfo = new
                    {
                        file1 = new 
                        { 
                            name = uploadedFile1.FileName, 
                            original = uploadedFile1.OriginalFileName,
                            sheet = sheet1Name,
                            rowCount = file1Data.Count
                        },
                        file2 = new 
                        { 
                            name = uploadedFile2.FileName, 
                            original = uploadedFile2.OriginalFileName,
                            sheet = sheet2Name,
                            rowCount = file2Data.Count
                        }
                    },
                    summary = new
                    {
                        totalDifferences = result.Differences?.Count ?? 0,
                        hasChanges = result.Differences?.Any() == true,
                        summary = result.Summary
                    },
                    message = result.Differences?.Any() == true 
                        ? $"Ýki Excel dosyasý karþýlaþtýrýldý. {result.Differences.Count} fark bulundu."
                        : "Ýki Excel dosyasý karþýlaþtýrýldý. Herhangi bir fark bulunamadý."
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Manuel dosya karþýlaþtýrmasý yapýlýrken hata");
                return StatusCode(500, new { 
                    success = false, 
                    message = "Karþýlaþtýrma yapýlýrken hata oluþtu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Yüklü dosyalar arasýnda karþýlaþtýrma - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPost("files")]
        public async Task<IActionResult> CompareFiles([FromBody] CompareFilesRequestDto compareRequest)
        {
            try
            {
                // Input validation
                if (compareRequest == null)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Karþýlaþtýrma isteði boþ olamaz" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName1))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Birinci dosya adý gerekli" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName2))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Ýkinci dosya adý gerekli" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(compareRequest.SheetName);

                _logger.LogInformation("Yüklü dosya karþýlaþtýrmasý baþlatýldý: {File1} vs {File2}, Sheet: {Sheet}", 
                    compareRequest.FileName1, compareRequest.FileName2, normalizedSheetName);

                var result = await _comparisonService.CompareFilesAsync(
                    compareRequest.FileName1, 
                    compareRequest.FileName2, 
                    normalizedSheetName);

                _logger.LogInformation("Karþýlaþtýrma tamamlandý: {DifferenceCount} fark bulundu", 
                    result.Differences?.Count ?? 0);

                return Ok(new { 
                    success = true, 
                    data = result,
                    summary = new
                    {
                        totalDifferences = result.Differences?.Count ?? 0,
                        hasChanges = result.Differences?.Any() == true
                    },
                    message = result.Differences?.Any() == true 
                        ? $"Dosyalar karþýlaþtýrýldý. {result.Differences.Count} fark bulundu."
                        : "Dosyalar karþýlaþtýrýldý. Herhangi bir fark bulunamadý."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosyalar karþýlaþtýrýlýrken hata: {File1} vs {File2}", 
                    compareRequest?.FileName1, compareRequest?.FileName2);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Dosya karþýlaþtýrmasý yapýlýrken hata oluþtu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Ayný dosyanýn farklý versiyonlarýný karþýlaþtýrma - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPost("versions")]
        public async Task<IActionResult> CompareVersions([FromBody] CompareVersionsRequestDto compareRequest)
        {
            try
            {
                // Input validation
                if (compareRequest == null)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Versiyon karþýlaþtýrma isteði boþ olamaz" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli" 
                    });
                }

                if (compareRequest.Version1Date >= compareRequest.Version2Date)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Ýkinci versiyon tarihi birinci versiyon tarihinden sonra olmalý" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(compareRequest.SheetName);

                _logger.LogInformation("Versiyon karþýlaþtýrmasý baþlatýldý: {FileName}, {Date1} vs {Date2}", 
                    compareRequest.FileName, compareRequest.Version1Date, compareRequest.Version2Date);

                var result = await _comparisonService.CompareVersionsAsync(
                    compareRequest.FileName, 
                    compareRequest.Version1Date, 
                    compareRequest.Version2Date, 
                    normalizedSheetName);

                _logger.LogInformation("Versiyon karþýlaþtýrmasý tamamlandý: {DifferenceCount} fark bulundu", 
                    result.Differences?.Count ?? 0);

                return Ok(new { 
                    success = true, 
                    data = result,
                    versionInfo = new
                    {
                        fileName = compareRequest.FileName,
                        version1Date = compareRequest.Version1Date,
                        version2Date = compareRequest.Version2Date,
                        sheetName = normalizedSheetName
                    },
                    summary = new
                    {
                        totalDifferences = result.Differences?.Count ?? 0,
                        hasChanges = result.Differences?.Any() == true
                    },
                    message = result.Differences?.Any() == true 
                        ? $"Versiyonlar karþýlaþtýrýldý. {result.Differences.Count} fark bulundu."
                        : "Versiyonlar karþýlaþtýrýldý. Herhangi bir fark bulunamadý."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Versiyonlar karþýlaþtýrýlýrken hata: {FileName}", compareRequest?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Versiyon karþýlaþtýrmasý yapýlýrken hata oluþtu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Belirli tarih aralýðýndaki deðiþiklikleri getirme - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpGet("changes/{fileName}")]
        public async Task<IActionResult> GetChanges(string fileName, [FromQuery] DateTime? fromDate = null, [FromQuery] DateTime? toDate = null, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);

                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli" 
                    });
                }

                if (fromDate.HasValue && toDate.HasValue && fromDate > toDate)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Baþlangýç tarihi bitiþ tarihinden önce olmalý" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("Deðiþiklik listesi istendi: {FileName}, {FromDate} - {ToDate}", 
                    fileName, fromDate, toDate);

                var changes = await _comparisonService.GetChangesAsync(fileName, fromDate, toDate, normalizedSheetName);

                return Ok(new { 
                    success = true, 
                    data = changes,
                    fileName = fileName,
                    sheetName = normalizedSheetName,
                    dateRange = new
                    {
                        fromDate = fromDate,
                        toDate = toDate
                    },
                    totalChanges = changes.Count,
                    message = changes.Any() 
                        ? $"{changes.Count} deðiþiklik bulundu" 
                        : "Belirtilen kriterlerde deðiþiklik bulunamadý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Deðiþiklikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Deðiþiklikler getirilirken hata oluþtu: " + ex.Message,
                    fileName = fileName,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Dosyanýn deðiþiklik geçmiþini getirme - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpGet("history/{fileName}")]
        public async Task<IActionResult> GetChangeHistory(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);

                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("Deðiþiklik geçmiþi istendi: {FileName}, Sheet: {SheetName}", fileName, normalizedSheetName);

                var history = await _comparisonService.GetChangeHistoryAsync(fileName, normalizedSheetName);

                return Ok(new { 
                    success = true, 
                    data = history,
                    fileName = fileName,
                    sheetName = normalizedSheetName,
                    totalEntries = history.Count,
                    message = history.Any() 
                        ? $"{history.Count} geçmiþ kaydý bulundu" 
                        : "Bu dosya için deðiþiklik geçmiþi bulunamadý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Deðiþiklik geçmiþi getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Deðiþiklik geçmiþi getirilirken hata oluþtu: " + ex.Message,
                    fileName = fileName,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Belirli bir satýrýn deðiþiklik geçmiþini getirme - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpGet("row-history/{rowId}")]
        public async Task<IActionResult> GetRowHistory(int rowId)
        {
            try
            {
                if (rowId <= 0)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Geçerli bir satýr ID'si gerekli" 
                    });
                }

                _logger.LogInformation("Satýr geçmiþi istendi: RowId={RowId}", rowId);

                var history = await _comparisonService.GetRowHistoryAsync(rowId);

                return Ok(new { 
                    success = true, 
                    data = history,
                    rowId = rowId,
                    totalVersions = history.Count,
                    message = history.Any() 
                        ? $"Satýr için {history.Count} versiyon bulundu" 
                        : "Bu satýr için geçmiþ bulunamadý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Satýr geçmiþi getirilirken hata: {RowId}", rowId);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Satýr geçmiþi getirilirken hata oluþtu: " + ex.Message,
                    rowId = rowId,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Ýki farklý tarihteki dosya durumunu karþýlaþtýrma - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPost("snapshot-compare")]
        public async Task<IActionResult> CompareSnapshots([FromBody] CompareVersionsRequestDto compareRequest)
        {
            try
            {
                // Input validation
                if (compareRequest == null)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Snapshot karþýlaþtýrma isteði boþ olamaz" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli" 
                    });
                }

                if (compareRequest.Version1Date >= compareRequest.Version2Date)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Ýkinci snapshot tarihi birinci snapshot tarihinden sonra olmalý" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(compareRequest.SheetName);

                _logger.LogInformation("Snapshot karþýlaþtýrmasý baþlatýldý: {FileName}, {Date1} vs {Date2}", 
                    compareRequest.FileName, compareRequest.Version1Date, compareRequest.Version2Date);

                var result = await _comparisonService.CompareVersionsAsync(
                    compareRequest.FileName, 
                    compareRequest.Version1Date, 
                    compareRequest.Version2Date, 
                    normalizedSheetName);

                _logger.LogInformation("Snapshot karþýlaþtýrmasý tamamlandý: {DifferenceCount} fark bulundu", 
                    result.Differences?.Count ?? 0);

                return Ok(new { 
                    success = true, 
                    data = result,
                    comparisonType = "snapshot",
                    snapshotInfo = new
                    {
                        fileName = compareRequest.FileName,
                        snapshot1Date = compareRequest.Version1Date,
                        snapshot2Date = compareRequest.Version2Date,
                        sheetName = normalizedSheetName,
                        timeDifference = compareRequest.Version2Date - compareRequest.Version1Date
                    },
                    summary = new
                    {
                        totalDifferences = result.Differences?.Count ?? 0,
                        hasChanges = result.Differences?.Any() == true
                    },
                    message = result.Differences?.Any() == true 
                        ? $"Farklý tarihlerdeki dosya durumlarý karþýlaþtýrýldý. {result.Differences.Count} fark bulundu."
                        : "Farklý tarihlerdeki dosya durumlarý karþýlaþtýrýldý. Herhangi bir fark bulunamadý."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Snapshot karþýlaþtýrmasý yapýlýrken hata: {FileName}", compareRequest?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Snapshot karþýlaþtýrmasý yapýlýrken hata oluþtu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Debug endpoint - Karþýlaþtýrma servisini test etme
        /// </summary>
        [HttpGet("debug/test-comparison/{fileName1}/{fileName2}")]
        public async Task<IActionResult> TestComparison(string fileName1, string fileName2, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode iþlemi
                fileName1 = Uri.UnescapeDataString(fileName1);
                fileName2 = Uri.UnescapeDataString(fileName2);

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("Test karþýlaþtýrmasý baþlatýldý: {File1} vs {File2}", fileName1, fileName2);

                // 1. Dosya verilerini kontrol et
                var file1Data = await _excelService.GetAllExcelDataAsync(fileName1, normalizedSheetName);
                var file2Data = await _excelService.GetAllExcelDataAsync(fileName2, normalizedSheetName);

                // 2. Karþýlaþtýrma servisini test et
                var comparisonResult = await _comparisonService.CompareFilesAsync(fileName1, fileName2, normalizedSheetName);

                return Ok(new
                {
                    success = true,
                    testResults = new
                    {
                        file1Info = new { fileName = fileName1, rowCount = file1Data.Count, hasData = file1Data.Any() },
                        file2Info = new { fileName = fileName2, rowCount = file2Data.Count, hasData = file2Data.Any() },
                        comparisonResult = new
                        {
                            differences = comparisonResult.Differences?.Count ?? 0,
                            summary = comparisonResult.Summary,
                            hasResult = comparisonResult != null
                        }
                    },
                    message = "Test karþýlaþtýrmasý tamamlandý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Test karþýlaþtýrmasýnda hata: {File1} vs {File2}", fileName1, fileName2);
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message,
                    file1 = fileName1,
                    file2 = fileName2
                });
            }
        }

        /// <summary>
        /// Debug endpoint - Comparison servisinin durumunu test etme
        /// </summary>
        [HttpGet("debug/status")]
        public async Task<IActionResult> GetComparisonStatus()
        {
            try
            {
                // Mevcut dosyalarý kontrol et
                var availableFiles = await _excelService.GetExcelFilesAsync();
                
                // Test için sample comparison yapabilir miyiz kontrol et
                var canCompare = availableFiles.Count >= 2;
                
                return Ok(new
                {
                    success = true,
                    comparisonServiceStatus = new
                    {
                        isOperational = true,
                        availableFiles = availableFiles.Select(f => new { f.FileName, f.OriginalFileName, f.IsActive }).ToList(),
                        totalFiles = availableFiles.Count,
                        canPerformComparison = canCompare,
                        supportedOperations = new[]
                        {
                            "File comparison",
                            "Version comparison", 
                            "Change tracking",
                            "History retrieval"
                        }
                    },
                    message = "Comparison servisi çalýþýyor"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Comparison status kontrolünde hata");
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message
                });
            }
        }

        /// <summary>
        /// Debug endpoint - Sample data ile comparison test
        /// </summary>
        [HttpPost("debug/test-sample")]
        public async Task<IActionResult> TestSampleComparison()
        {
            try
            {
                // En az 2 dosya var mý kontrol et
                var files = await _excelService.GetExcelFilesAsync();
                
                if (files.Count < 2)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Test için en az 2 dosya gerekli",
                        availableFiles = files.Count,
                        suggestion = "Önce 2 Excel dosyasý yükleyin"
                    });
                }

                var file1 = files[0];
                var file2 = files[1];

                _logger.LogInformation("Sample comparison test baþlatýldý: {File1} vs {File2}", file1.FileName, file2.FileName);

                // Test comparison yap
                var result = await _comparisonService.CompareFilesAsync(file1.FileName, file2.FileName);

                return Ok(new
                {
                    success = true,
                    testResult = new
                    {
                        comparedFiles = new
                        {
                            file1 = new { file1.FileName, file1.OriginalFileName },
                            file2 = new { file2.FileName, file2.OriginalFileName }
                        },
                        comparisonResult = new
                        {
                            differences = result.Differences?.Count ?? 0,
                            summary = result.Summary,
                            hasResult = result != null,
                            comparisonId = result.ComparisonId
                        }
                    },
                    message = "Sample comparison test baþarýlý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sample comparison test'te hata");
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message
                });
            }
        }

        /// <summary>
        /// SheetName'i normalize eder - "undefined" string'ini null'a çevirir
        /// </summary>
        private string? NormalizeSheetName(string? sheetName)
        {
            if (string.IsNullOrEmpty(sheetName) || 
                string.Equals(sheetName, "undefined", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(sheetName, "null", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            return sheetName.Trim();
        }
    }
}