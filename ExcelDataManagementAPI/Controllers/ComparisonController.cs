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
                message = "Comparison API �al���yor!", 
                timestamp = DateTime.Now,
                availableOperations = new[]
                {
                    "Dosya kar��la�t�rma: POST /api/comparison/files",
                    "Manuel dosya kar��la�t�rma: POST /api/comparison/compare-from-files",
                    "Versiyon kar��la�t�rma: POST /api/comparison/versions",
                    "De�i�iklik ge�mi�i: GET /api/comparison/history/{fileName}"
                }
            });
        }

        /// <summary>
        /// Bilgisayardan iki Excel dosyas� se�ip kar��la�t�rma - �Y�LE�T�R�LD�
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
                        message = "Kar��la�t�rma iste�i bo� olamaz" 
                    });
                }

                if (request.File1 == null || request.File1.Length == 0)
                    return BadRequest(new { 
                        success = false, 
                        message = "L�tfen birinci Excel dosyas�n� se�in" 
                    });

                if (request.File2 == null || request.File2.Length == 0)
                    return BadRequest(new { 
                        success = false, 
                        message = "L�tfen ikinci Excel dosyas�n� se�in" 
                    });

                // Dosya uzant�s� kontrol�
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var file1Extension = Path.GetExtension(request.File1.FileName).ToLowerInvariant();
                var file2Extension = Path.GetExtension(request.File2.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(file1Extension) || !allowedExtensions.Contains(file2Extension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalar� (.xlsx, .xls) desteklenir" 
                    });
                }

                // SheetName'leri normalize et - "undefined" ise null yap
                var sheet1Name = NormalizeSheetName(request.Sheet1Name);
                var sheet2Name = NormalizeSheetName(request.Sheet2Name);

                _logger.LogInformation("Excel dosya kar��la�t�rmas� ba�lat�ld�: {File1} vs {File2}", 
                    request.File1.FileName, request.File2.FileName);

                // Her iki dosyay� da y�kle
                var uploadedFile1 = await _excelService.UploadExcelFileAsync(request.File1, request.ComparedBy);
                var uploadedFile2 = await _excelService.UploadExcelFileAsync(request.File2, request.ComparedBy);

                _logger.LogInformation("Dosyalar y�klendi: {File1} -> {InternalName1}, {File2} -> {InternalName2}", 
                    request.File1.FileName, uploadedFile1.FileName, request.File2.FileName, uploadedFile2.FileName);

                // Her iki dosyay� da oku
                var file1Data = await _excelService.ReadExcelDataAsync(uploadedFile1.FileName, sheet1Name);
                var file2Data = await _excelService.ReadExcelDataAsync(uploadedFile2.FileName, sheet2Name);

                _logger.LogInformation("Dosya verileri okundu: File1={RowCount1} rows, File2={RowCount2} rows", 
                    file1Data.Count, file2Data.Count);

                // Kar��la�t�r
                var result = await _comparisonService.CompareFilesAsync(
                    uploadedFile1.FileName, 
                    uploadedFile2.FileName, 
                    sheet1Name ?? sheet2Name);

                _logger.LogInformation("Kar��la�t�rma tamamland�: {DifferenceCount} fark bulundu", 
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
                        ? $"�ki Excel dosyas� kar��la�t�r�ld�. {result.Differences.Count} fark bulundu."
                        : "�ki Excel dosyas� kar��la�t�r�ld�. Herhangi bir fark bulunamad�."
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Manuel dosya kar��la�t�rmas� yap�l�rken hata");
                return StatusCode(500, new { 
                    success = false, 
                    message = "Kar��la�t�rma yap�l�rken hata olu�tu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Y�kl� dosyalar aras�nda kar��la�t�rma - �Y�LE�T�R�LD�
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
                        message = "Kar��la�t�rma iste�i bo� olamaz" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName1))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Birinci dosya ad� gerekli" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName2))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "�kinci dosya ad� gerekli" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(compareRequest.SheetName);

                _logger.LogInformation("Y�kl� dosya kar��la�t�rmas� ba�lat�ld�: {File1} vs {File2}, Sheet: {Sheet}", 
                    compareRequest.FileName1, compareRequest.FileName2, normalizedSheetName);

                var result = await _comparisonService.CompareFilesAsync(
                    compareRequest.FileName1, 
                    compareRequest.FileName2, 
                    normalizedSheetName);

                _logger.LogInformation("Kar��la�t�rma tamamland�: {DifferenceCount} fark bulundu", 
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
                        ? $"Dosyalar kar��la�t�r�ld�. {result.Differences.Count} fark bulundu."
                        : "Dosyalar kar��la�t�r�ld�. Herhangi bir fark bulunamad�."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosyalar kar��la�t�r�l�rken hata: {File1} vs {File2}", 
                    compareRequest?.FileName1, compareRequest?.FileName2);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Dosya kar��la�t�rmas� yap�l�rken hata olu�tu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Ayn� dosyan�n farkl� versiyonlar�n� kar��la�t�rma - �Y�LE�T�R�LD�
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
                        message = "Versiyon kar��la�t�rma iste�i bo� olamaz" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� gerekli" 
                    });
                }

                if (compareRequest.Version1Date >= compareRequest.Version2Date)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "�kinci versiyon tarihi birinci versiyon tarihinden sonra olmal�" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(compareRequest.SheetName);

                _logger.LogInformation("Versiyon kar��la�t�rmas� ba�lat�ld�: {FileName}, {Date1} vs {Date2}", 
                    compareRequest.FileName, compareRequest.Version1Date, compareRequest.Version2Date);

                var result = await _comparisonService.CompareVersionsAsync(
                    compareRequest.FileName, 
                    compareRequest.Version1Date, 
                    compareRequest.Version2Date, 
                    normalizedSheetName);

                _logger.LogInformation("Versiyon kar��la�t�rmas� tamamland�: {DifferenceCount} fark bulundu", 
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
                        ? $"Versiyonlar kar��la�t�r�ld�. {result.Differences.Count} fark bulundu."
                        : "Versiyonlar kar��la�t�r�ld�. Herhangi bir fark bulunamad�."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Versiyonlar kar��la�t�r�l�rken hata: {FileName}", compareRequest?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Versiyon kar��la�t�rmas� yap�l�rken hata olu�tu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Belirli tarih aral���ndaki de�i�iklikleri getirme - �Y�LE�T�R�LD�
        /// </summary>
        [HttpGet("changes/{fileName}")]
        public async Task<IActionResult> GetChanges(string fileName, [FromQuery] DateTime? fromDate = null, [FromQuery] DateTime? toDate = null, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);

                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� gerekli" 
                    });
                }

                if (fromDate.HasValue && toDate.HasValue && fromDate > toDate)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Ba�lang�� tarihi biti� tarihinden �nce olmal�" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("De�i�iklik listesi istendi: {FileName}, {FromDate} - {ToDate}", 
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
                        ? $"{changes.Count} de�i�iklik bulundu" 
                        : "Belirtilen kriterlerde de�i�iklik bulunamad�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "De�i�iklikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "De�i�iklikler getirilirken hata olu�tu: " + ex.Message,
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
        /// Dosyan�n de�i�iklik ge�mi�ini getirme - �Y�LE�T�R�LD�
        /// </summary>
        [HttpGet("history/{fileName}")]
        public async Task<IActionResult> GetChangeHistory(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);

                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� gerekli" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("De�i�iklik ge�mi�i istendi: {FileName}, Sheet: {SheetName}", fileName, normalizedSheetName);

                var history = await _comparisonService.GetChangeHistoryAsync(fileName, normalizedSheetName);

                return Ok(new { 
                    success = true, 
                    data = history,
                    fileName = fileName,
                    sheetName = normalizedSheetName,
                    totalEntries = history.Count,
                    message = history.Any() 
                        ? $"{history.Count} ge�mi� kayd� bulundu" 
                        : "Bu dosya i�in de�i�iklik ge�mi�i bulunamad�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "De�i�iklik ge�mi�i getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "De�i�iklik ge�mi�i getirilirken hata olu�tu: " + ex.Message,
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
        /// Belirli bir sat�r�n de�i�iklik ge�mi�ini getirme - �Y�LE�T�R�LD�
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
                        message = "Ge�erli bir sat�r ID'si gerekli" 
                    });
                }

                _logger.LogInformation("Sat�r ge�mi�i istendi: RowId={RowId}", rowId);

                var history = await _comparisonService.GetRowHistoryAsync(rowId);

                return Ok(new { 
                    success = true, 
                    data = history,
                    rowId = rowId,
                    totalVersions = history.Count,
                    message = history.Any() 
                        ? $"Sat�r i�in {history.Count} versiyon bulundu" 
                        : "Bu sat�r i�in ge�mi� bulunamad�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sat�r ge�mi�i getirilirken hata: {RowId}", rowId);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Sat�r ge�mi�i getirilirken hata olu�tu: " + ex.Message,
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
        /// �ki farkl� tarihteki dosya durumunu kar��la�t�rma - �Y�LE�T�R�LD�
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
                        message = "Snapshot kar��la�t�rma iste�i bo� olamaz" 
                    });
                }

                if (string.IsNullOrEmpty(compareRequest.FileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� gerekli" 
                    });
                }

                if (compareRequest.Version1Date >= compareRequest.Version2Date)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "�kinci snapshot tarihi birinci snapshot tarihinden sonra olmal�" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(compareRequest.SheetName);

                _logger.LogInformation("Snapshot kar��la�t�rmas� ba�lat�ld�: {FileName}, {Date1} vs {Date2}", 
                    compareRequest.FileName, compareRequest.Version1Date, compareRequest.Version2Date);

                var result = await _comparisonService.CompareVersionsAsync(
                    compareRequest.FileName, 
                    compareRequest.Version1Date, 
                    compareRequest.Version2Date, 
                    normalizedSheetName);

                _logger.LogInformation("Snapshot kar��la�t�rmas� tamamland�: {DifferenceCount} fark bulundu", 
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
                        ? $"Farkl� tarihlerdeki dosya durumlar� kar��la�t�r�ld�. {result.Differences.Count} fark bulundu."
                        : "Farkl� tarihlerdeki dosya durumlar� kar��la�t�r�ld�. Herhangi bir fark bulunamad�."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Snapshot kar��la�t�rmas� yap�l�rken hata: {FileName}", compareRequest?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Snapshot kar��la�t�rmas� yap�l�rken hata olu�tu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Debug endpoint - Kar��la�t�rma servisini test etme
        /// </summary>
        [HttpGet("debug/test-comparison/{fileName1}/{fileName2}")]
        public async Task<IActionResult> TestComparison(string fileName1, string fileName2, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName1 = Uri.UnescapeDataString(fileName1);
                fileName2 = Uri.UnescapeDataString(fileName2);

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("Test kar��la�t�rmas� ba�lat�ld�: {File1} vs {File2}", fileName1, fileName2);

                // 1. Dosya verilerini kontrol et
                var file1Data = await _excelService.GetAllExcelDataAsync(fileName1, normalizedSheetName);
                var file2Data = await _excelService.GetAllExcelDataAsync(fileName2, normalizedSheetName);

                // 2. Kar��la�t�rma servisini test et
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
                    message = "Test kar��la�t�rmas� tamamland�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Test kar��la�t�rmas�nda hata: {File1} vs {File2}", fileName1, fileName2);
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
                // Mevcut dosyalar� kontrol et
                var availableFiles = await _excelService.GetExcelFilesAsync();
                
                // Test i�in sample comparison yapabilir miyiz kontrol et
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
                    message = "Comparison servisi �al���yor"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Comparison status kontrol�nde hata");
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
                // En az 2 dosya var m� kontrol et
                var files = await _excelService.GetExcelFilesAsync();
                
                if (files.Count < 2)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Test i�in en az 2 dosya gerekli",
                        availableFiles = files.Count,
                        suggestion = "�nce 2 Excel dosyas� y�kleyin"
                    });
                }

                var file1 = files[0];
                var file2 = files[1];

                _logger.LogInformation("Sample comparison test ba�lat�ld�: {File1} vs {File2}", file1.FileName, file2.FileName);

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
                    message = "Sample comparison test ba�ar�l�"
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
        /// SheetName'i normalize eder - "undefined" string'ini null'a �evirir
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