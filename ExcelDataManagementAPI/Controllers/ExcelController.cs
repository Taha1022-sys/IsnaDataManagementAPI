using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Services;
using ExcelDataManagementAPI.Models.DTOs;
using System.Text.Json;

namespace ExcelDataManagementAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelService _excelService;
        private readonly ExcelDataContext _context;
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(IExcelService excelService, ExcelDataContext context, ILogger<ExcelController> logger)
        {
            _excelService = excelService;
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// API test endpoint'i
        /// </summary>
        [HttpGet("test")]
        public IActionResult Test()
        {
            return Ok(new { 
                message = "Excel API çalýþýyor!", 
                timestamp = DateTime.Now,
                availableOperations = new[]
                {
                    "Excel dosyasý yükleme: POST /api/excel/upload",
                    "Dosya listesi: GET /api/excel/files", 
                    "Dosya silme: DELETE /api/excel/files/{fileName}",
                    "Manuel dosya seçip okuma: POST /api/excel/read-from-file",
                    "Manuel dosya seçip güncelleme: POST /api/excel/update-from-file",
                    "Veri düzenleme: PUT /api/excel/data",
                    "Tüm verileri getirme: GET /api/excel/data/{fileName}/all",
                    "Sayfalý veri getirme: GET /api/excel/data/{fileName}?page=1&pageSize=50"
                },
                debugEndpoints = new[]
                {
                    "Veritabaný durumu: GET /api/excel/debug/database-status",
                    "Dosya durumu: GET /api/excel/files/{fileName}/status",
                    "Veri sorgusu test: GET /api/excel/debug/test-data-query/{fileName}",
                    "Veri akýþ testi: GET /api/excel/debug/data-flow-test/{fileName}"
                }
            });
        }

        /// <summary>
        /// Debug endpoint - Veritabaný durumunu kontrol etme
        /// </summary>
        [HttpGet("debug/database-status")]
        public async Task<IActionResult> GetDatabaseStatus()
        {
            try
            {
                var filesCount = await _excelService.GetExcelFilesAsync();
                var totalDataRows = await _context.ExcelDataRows.CountAsync();
                var activeDataRows = await _context.ExcelDataRows.Where(r => !r.IsDeleted).CountAsync();
                var deletedDataRows = await _context.ExcelDataRows.Where(r => r.IsDeleted).CountAsync();

                // Her dosya için detaylý bilgi
                var fileDetails = new List<object>();
                foreach (var file in filesCount)
                {
                    var fileDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == file.FileName && !r.IsDeleted)
                        .CountAsync();

                    var sheets = await _context.ExcelDataRows
                        .Where(r => r.FileName == file.FileName && !r.IsDeleted)
                        .Select(r => r.SheetName)
                        .Distinct()
                        .ToListAsync();

                    fileDetails.Add(new
                    {
                        file.FileName,
                        file.OriginalFileName,
                        file.IsActive,
                        file.UploadDate,
                        file.FileSize,
                        DataRowCount = fileDataCount,
                        AvailableSheets = sheets,
                        PhysicalFileExists = !string.IsNullOrEmpty(file.FilePath) && System.IO.File.Exists(file.FilePath),
                        FilePath = file.FilePath
                    });
                }

                return Ok(new
                {
                    success = true,
                    databaseStatus = new
                    {
                        totalFiles = filesCount.Count,
                        activeFiles = filesCount.Where(f => f.IsActive).Count(),
                        totalDataRows = totalDataRows,
                        activeDataRows = activeDataRows,
                        deletedDataRows = deletedDataRows,
                        files = fileDetails
                    },
                    message = "Veritabaný durumu baþarýyla alýndý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veritabaný durumu kontrol edilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Yüklü Excel dosyalarýnýn listesi
        /// </summary>
        [HttpGet("files")]
        public async Task<IActionResult> GetFiles()
        {
            try
            {
                var files = await _excelService.GetExcelFilesAsync();
                return Ok(new { 
                    success = true, 
                    data = files,
                    count = files.Count,
                    message = "Yüklü Excel dosyalarý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosyalar getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Excel dosyasý yükleme - Bilgisayardan dosya seçimi
        /// </summary>
        [HttpPost("upload")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> Upload(IFormFile file, [FromForm] string? uploadedBy = null)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest(new { 
                        success = false, 
                        message = "Lütfen bilgisayarýnýzdan bir Excel dosyasý seçin (.xlsx veya .xls)" 
                    });

                // Dosya uzantýsý kontrolü
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalarý (.xlsx, .xls) desteklenir" 
                    });
                }

                _logger.LogInformation("Dosya yükleniyor: OriginalName={OriginalName}, Size={Size}", file.FileName, file.Length);

                var result = await _excelService.UploadExcelFileAsync(file, uploadedBy);
                
                _logger.LogInformation("Dosya yüklendi: OriginalName={OriginalName} -> GeneratedName={GeneratedName}", 
                    result.OriginalFileName, result.FileName);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    // Frontend için hem gerçek hem orijinal dosya adýný ver
                    fileName = result.FileName,
                    originalFileName = result.OriginalFileName,
                    message = "Dosya baþarýyla yüklendi! Þimdi iþlem yapmak için dosyayý kullanabilirsiniz.",
                    nextSteps = new 
                    {
                        // Gerçek dosya adýný kullan
                        readData = $"/api/excel/read/{result.FileName}",
                        getData = $"/api/excel/data/{result.FileName}",
                        getSheets = $"/api/excel/sheets/{result.FileName}",
                        // Debug endpoint'leri
                        checkStatus = $"/api/excel/files/{result.FileName}/status",
                        recentUploads = "/api/excel/debug/recent-uploads"
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya yüklenirken hata: {FileName}", file?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message,
                    originalFileName = file?.FileName
                });
            }
        }

        /// <summary>
        /// Bilgisayardan Excel dosyasý seçip direkt okuma
        /// </summary>
        [HttpPost("read-from-file")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> ReadFromFile([FromForm] ManualFileSelectionDto request)
        {
            try
            {
                if (request.ExcelFile == null || request.ExcelFile.Length == 0)
                    return BadRequest(new { 
                        success = false, 
                        message = "Lütfen bilgisayarýnýzdan bir Excel dosyasý seçin" 
                    });

                // Dosya uzantýsý kontrolü
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(request.ExcelFile.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalarý (.xlsx, .xls) desteklenir" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(request.SheetName);

                _logger.LogInformation("ReadFromFile baþlatýldý: OriginalName={OriginalName}, Sheet={Sheet}", 
                    request.ExcelFile.FileName, normalizedSheetName);

                // Önce dosyayý yükle
                var uploadedFile = await _excelService.UploadExcelFileAsync(request.ExcelFile, request.ProcessedBy);
                
                _logger.LogInformation("Dosya yüklendi: {OriginalName} -> {GeneratedName}", 
                    uploadedFile.OriginalFileName, uploadedFile.FileName);

                // Sonra oku
                var data = await _excelService.ReadExcelDataAsync(uploadedFile.FileName, normalizedSheetName);
                
                _logger.LogInformation("Dosya okundu: {Count} satýr", data.Count);
                
                return Ok(new { 
                    success = true, 
                    data = data,
                    fileName = uploadedFile.FileName,
                    originalFileName = uploadedFile.OriginalFileName,
                    sheetName = normalizedSheetName,
                    totalRows = data.Count,
                    message = "Excel dosyasý baþarýyla okundu ve veritabanýna aktarýldý",
                    debug = new
                    {
                        uploadedFileName = uploadedFile.FileName,
                        originalFileName = uploadedFile.OriginalFileName,
                        processedSheet = normalizedSheetName
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya okunurken hata: {FileName}", request?.ExcelFile?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message,
                    originalFileName = request?.ExcelFile?.FileName
                });
            }
        }

        /// <summary>
        /// Bilgisayardan Excel dosyasý seçip güncelleme
        /// </summary>
        [HttpPost("update-from-file")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> UpdateFromFile([FromForm] UpdateExcelFileDto request)
        {
            try
            {
                if (request.ExcelFile == null || request.ExcelFile.Length == 0)
                    return BadRequest(new { 
                        success = false, 
                        message = "Lütfen bilgisayarýnýzdan bir Excel dosyasý seçin" 
                    });

                // Dosya uzantýsý kontrolü
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(request.ExcelFile.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalarý (.xlsx, .xls) desteklenir" 
                    });
                }

                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(request.SheetName);

                // Önce dosyayý yükle
                var uploadedFile = await _excelService.UploadExcelFileAsync(request.ExcelFile, request.UpdatedBy);
                
                // Dosyayý oku
                var data = await _excelService.ReadExcelDataAsync(uploadedFile.FileName, normalizedSheetName);
                
                return Ok(new { 
                    success = true, 
                    uploadedFile = uploadedFile,
                    data = data,
                    totalRows = data.Count,
                    message = "Excel dosyasý yüklendi ve okundu. Þimdi güncellemeler yapabilirsiniz.",
                    availableOperations = new 
                    {
                        getData = $"/api/excel/data/{uploadedFile.FileName}",
                        updateData = "/api/excel/data",
                        addRow = "/api/excel/data",
                        deleteRow = "/api/excel/data/{id}",
                        export = "/api/excel/export"
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya güncellenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Yüklü dosyadan veri okuma ve veritabanýna aktarma
        /// </summary>
        [HttpPost("read/{fileName}")]
        public async Task<IActionResult> ReadExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);
                
                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("ReadExcelData çaðrýldý: FileName={FileName}, SheetName={SheetName}", fileName, normalizedSheetName);

                // Önce tam eþleþme ara
                var exactMatch = await _context.ExcelFiles
                    .FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);

                if (exactMatch == null)
                {
                    // Tam eþleþme yoksa, orijinal dosya adýyla ara
                    var originalMatch = await _context.ExcelFiles
                        .FirstOrDefaultAsync(f => f.OriginalFileName == fileName && f.IsActive);

                    if (originalMatch != null)
                    {
                        _logger.LogInformation("Orijinal dosya adýyla eþleþme bulundu: {OriginalName} -> {FileName}", fileName, originalMatch.FileName);
                        fileName = originalMatch.FileName; // Gerçek dosya adýný kullan
                    }
                    else
                    {
                        // Partial match dene (dosya adý uzantýsý olmadan)
                        var fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                        var partialMatch = await _context.ExcelFiles
                            .Where(f => f.IsActive && 
                                       (f.FileName.Contains(fileNameWithoutExt) || 
                                        f.OriginalFileName.Contains(fileNameWithoutExt)))
                            .OrderByDescending(f => f.UploadDate)
                            .FirstOrDefaultAsync();

                        if (partialMatch != null)
                        {
                            _logger.LogInformation("Kýsmi eþleþme bulundu: {RequestedName} -> {FileName}", fileName, partialMatch.FileName);
                            fileName = partialMatch.FileName; // Gerçek dosya adýný kullan
                        }
                        else
                        {
                            _logger.LogWarning("Dosya bulunamadý: {RequestedFileName}", fileName);
                            
                            var availableFiles = await _context.ExcelFiles
                                .Where(f => f.IsActive)
                                .Select(f => new { f.FileName, f.OriginalFileName, f.UploadDate })
                                .OrderByDescending(f => f.UploadDate)
                                .ToListAsync();

                            return NotFound(new
                            {
                                success = false,
                                message = $"Dosya bulunamadý: {fileName}",
                                requestedFileName = fileName,
                                availableFiles = availableFiles,
                                suggestion = "Lütfen mevcut dosyalardan birini seçin veya dosyayý yeniden yükleyin"
                            });
                        }
                    }
                }

                var data = await _excelService.ReadExcelDataAsync(fileName, normalizedSheetName);
                return Ok(new { 
                    success = true, 
                    data = data,
                    fileName = fileName,
                    sheetName = normalizedSheetName,
                    totalRows = data.Count,
                    message = "Excel verisi baþarýyla okundu ve veritabanýna aktarýldý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyasý okunurken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message,
                    requestedFileName = fileName
                });
            }
        }

        /// <summary>
        /// Dosyadan veri getirme (sayfalama ile) - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpGet("data/{fileName}")]
        public async Task<IActionResult> GetExcelData(string fileName, [FromQuery] string? sheetName = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);
                
                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("GetExcelData çaðrýldý: FileName={FileName}, SheetName={SheetName}, Page={Page}, PageSize={PageSize}", 
                    fileName, normalizedSheetName, page, pageSize);

                // Parametre validasyonu
                if (page < 1) page = 1;
                if (pageSize < 1 || pageSize > 1000) pageSize = 50;

                // Önce dosyanýn var olup olmadýðýný kontrol et
                var fileExists = await _context.ExcelFiles
                    .AnyAsync(f => f.FileName == fileName && f.IsActive);

                if (!fileExists)
                {
                    _logger.LogWarning("Dosya bulunamadý: {FileName}", fileName);
                    return NotFound(new
                    {
                        success = false,
                        message = "Belirtilen dosya bulunamadý. Dosyanýn yüklendiðinden ve aktif olduðundan emin olun.",
                        fileName = fileName,
                        availableFiles = await _context.ExcelFiles
                            .Where(f => f.IsActive)
                            .Select(f => new { f.FileName, f.OriginalFileName })
                            .ToListAsync()
                    });
                }

                // Verileri service'den getir
                var data = await _excelService.GetExcelDataAsync(fileName, normalizedSheetName, page, pageSize);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, normalizedSheetName);

                // Debug bilgisi
                _logger.LogInformation("Service'den dönen veri sayýsý: {Count}", data?.Count ?? 0);

                // Eðer veri yoksa ancak dosya varsa, dosyanýn okunup okunmadýðýný kontrol et
                if (data == null || data.Count == 0)
                {
                    var totalDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .CountAsync();

                    if (totalDataCount == 0)
                    {
                        // Dosya var ama veri yok - okunmamýþ olabilir
                        var availableSheets = await _excelService.GetSheetsAsync(fileName);
                        
                        _logger.LogWarning("Dosya bulundu ancak veri yok: {FileName}", fileName);
                        return Ok(new
                        {
                            success = true,
                            data = new List<object>(),
                            fileName = fileName,
                            sheetName = normalizedSheetName,
                            page = page,
                            pageSize = pageSize,
                            totalRows = 0,
                            totalPages = 0,
                            statistics = statistics,
                            hasData = false,
                            availableSheets = availableSheets,
                            message = "Dosya bulundu ancak henüz okunmamýþ. Önce dosyayý okuyun.",
                            suggestedActions = new
                            {
                                readFile = $"/api/excel/read/{fileName}",
                                readWithSheet = availableSheets.Any() ? $"/api/excel/read/{fileName}?sheetName={availableSheets.First()}" : null
                            }
                        });
                    }
                    
                    // Sheet belirtilmiþse ve o sheet'te veri yoksa
                    if (!string.IsNullOrEmpty(normalizedSheetName))
                    {
                        var sheetExists = await _context.ExcelDataRows
                            .AnyAsync(r => r.FileName == fileName && r.SheetName == normalizedSheetName && !r.IsDeleted);

                        if (!sheetExists)
                        {
                            var availableSheets = await _context.ExcelDataRows
                                .Where(r => r.FileName == fileName && !r.IsDeleted)
                                .Select(r => r.SheetName)
                                .Distinct()
                                .ToListAsync();

                            return NotFound(new
                            {
                                success = false,
                                message = $"Belirtilen sayfa '{normalizedSheetName}' bulunamadý.",
                                fileName = fileName,
                                requestedSheet = normalizedSheetName,
                                availableSheets = availableSheets,
                                suggestion = availableSheets.Any() ? $"Mevcut sayfalardan birini seçin: {string.Join(", ", availableSheets)}" : "Bu dosyada henüz veri bulunmuyor."
                            });
                        }
                    }
                }

                // Baþarýlý response
                var totalRowsForPagination = statistics != null ? (int)(statistics.GetType().GetProperty("totalRows")?.GetValue(statistics) ?? 0) : 0;
                var totalPages = totalRowsForPagination > 0 ? (int)Math.Ceiling((double)totalRowsForPagination / pageSize) : 0;
                
                var response = new
                {
                    success = true,
                    data = data ?? new List<ExcelDataResponseDto>(),
                    fileName = fileName,
                    sheetName = normalizedSheetName,
                    page = page,
                    pageSize = pageSize,
                    totalRows = totalRowsForPagination,
                    totalPages = totalPages,
                    currentPageRowCount = data?.Count ?? 0,
                    hasData = data != null && data.Count > 0,
                    statistics = statistics,
                    pagination = new
                    {
                        currentPage = page,
                        pageSize = pageSize,
                        totalPages = totalPages,
                        totalRows = totalRowsForPagination,
                        hasNext = page < totalPages,
                        hasPrevious = page > 1
                    },
                    message = data != null && data.Count > 0 
                        ? $"Sayfa {page}/{totalPages} - {data.Count} kayýt gösteriliyor" 
                        : "Bu sayfada gösterilecek veri bulunamadý"
                };

                _logger.LogInformation("Response hazýrlandý: TotalRows={TotalRows}, PageCount={PageCount}", 
                    totalRowsForPagination, data?.Count ?? 0);

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}, Page: {Page}", fileName, sheetName, page);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veriler getirilirken bir hata oluþtu: " + ex.Message,
                    fileName = fileName,
                    sheetName = sheetName,
                    page = page,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message,
                        stackTrace = ex.StackTrace
                    }
                });
            }
        }

        /// <summary>
        /// Dosyadan tüm verileri getirme (sayfalama olmadan) - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpGet("data/{fileName}/all")]
        public async Task<IActionResult> GetAllExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);
                
                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);

                _logger.LogInformation("GetAllExcelData çaðrýldý: FileName={FileName}, SheetName={SheetName}", fileName, normalizedSheetName);

                // Önce dosyanýn var olup olmadýðýný kontrol et
                var fileExists = await _context.ExcelFiles
                    .AnyAsync(f => f.FileName == fileName && f.IsActive);

                if (!fileExists)
                {
                    _logger.LogWarning("Dosya bulunamadý: {FileName}", fileName);
                    return NotFound(new
                    {
                        success = false,
                        message = "Belirtilen dosya bulunamadý. Dosyanýn yüklendiðinden ve aktif olduðundan emin olun.",
                        fileName = fileName,
                        availableFiles = await _context.ExcelFiles
                            .Where(f => f.IsActive)
                            .Select(f => new { f.FileName, f.OriginalFileName })
                            .ToListAsync()
                    });
                }

                var data = await _excelService.GetAllExcelDataAsync(fileName, normalizedSheetName);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, normalizedSheetName);

                // Debug bilgisi
                _logger.LogInformation("Service'den dönen tüm veri sayýsý: {Count}", data?.Count ?? 0);

                // Eðer veri yoksa detaylý bilgi ver
                if (data == null || data.Count == 0)
                {
                    var totalDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .CountAsync();

                    if (totalDataCount == 0)
                    {
                        // Dosya var ama veri yok - okunmamýþ olabilir
                        var availableSheets = await _excelService.GetSheetsAsync(fileName);
                        
                        _logger.LogWarning("Dosya bulundu ancak veri yok: {FileName}", fileName);
                        return Ok(new
                        {
                            success = true,
                            data = new List<object>(),
                            fileName = fileName,
                            sheetName = normalizedSheetName,
                            totalRows = 0,
                            statistics = statistics,
                            hasData = false,
                            availableSheets = availableSheets,
                            message = "Dosya bulundu ancak henüz okunmamýþ. Önce dosyayý okuyun.",
                            suggestedActions = new
                            {
                                readFile = $"/api/excel/read/{fileName}",
                                readWithSheet = availableSheets.Any() ? $"/api/excel/read/{fileName}?sheetName={availableSheets.First()}" : null
                            }
                        });
                    }
                    
                    // Sheet belirtilmiþse ve o sheet'te veri yoksa
                    if (!string.IsNullOrEmpty(normalizedSheetName))
                    {
                        var availableSheets = await _context.ExcelDataRows
                            .Where(r => r.FileName == fileName && !r.IsDeleted)
                            .Select(r => r.SheetName)
                            .Distinct()
                            .ToListAsync();

                        return NotFound(new
                        {
                            success = false,
                            message = $"Belirtilen sayfa '{normalizedSheetName}' bulunamadý.",
                            fileName = fileName,
                            requestedSheet = normalizedSheetName,
                            availableSheets = availableSheets,
                            suggestion = availableSheets.Any() ? $"Mevcut sayfalardan birini seçin: {string.Join(", ", availableSheets)}" : "Bu dosyada henüz veri bulunmuyor."
                        });
                    }
                }
               
                return Ok(new { 
                    success = true, 
                    data = data ?? new List<ExcelDataResponseDto>(), 
                    fileName = fileName,
                    sheetName = normalizedSheetName,
                    totalRows = data?.Count ?? 0,
                    statistics = statistics,
                    hasData = data != null && data.Count > 0,
                    message = data != null && data.Count > 0 ? $"Toplam {data.Count} kayýt getirildi" : "Bu dosya/sayfa için veri bulunamadý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Tüm Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}", fileName, sheetName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veriler getirilirken bir hata oluþtu: " + ex.Message,
                    fileName = fileName,
                    sheetName = sheetName,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message,
                        stackTrace = ex.StackTrace
                    }
                });
            }
        }

        /// <summary>
        /// Excel verisini güncelleme - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPut("data")]
        public async Task<IActionResult> UpdateExcelData([FromBody] ExcelDataUpdateDto updateDto)
        {
            try
            {
                if (updateDto == null || updateDto.Id <= 0)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Geçersiz güncelleme verisi. ID zorunludur." 
                    });
                }

                if (updateDto.Data == null || !updateDto.Data.Any())
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Güncellenecek veri boþ olamaz." 
                    });
                }

                _logger.LogInformation("UpdateExcelData çaðrýldý: Id={Id}, DataCount={DataCount}", 
                    updateDto.Id, updateDto.Data?.Count ?? 0);

                var result = await _excelService.UpdateExcelDataAsync(updateDto);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    message = "Veri baþarýyla güncellendi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri güncellenirken hata: {Id}", updateDto?.Id);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veri güncellenirken hata oluþtu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Toplu Excel verisini güncelleme - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPut("data/bulk")]
        public async Task<IActionResult> BulkUpdateExcelData([FromBody] BulkUpdateDto bulkUpdateDto)
        {
            try
            {
                if (bulkUpdateDto == null || bulkUpdateDto.Updates == null || !bulkUpdateDto.Updates.Any())
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Toplu güncelleme için en az bir güncelleme verisi gerekli." 
                    });
                }

                _logger.LogInformation("BulkUpdateExcelData çaðrýldý: UpdateCount={Count}", bulkUpdateDto.Updates.Count);

                var results = await _excelService.BulkUpdateExcelDataAsync(bulkUpdateDto);
                
                return Ok(new { 
                    success = true, 
                    data = results, 
                    count = results.Count,
                    message = $"{results.Count} kayýt baþarýyla güncellendi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Toplu veri güncellenirken hata");
                return StatusCode(500, new { 
                    success = false, 
                    message = "Toplu veri güncellenirken hata oluþtu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Yeni bir satýr ekleme - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPost("data")]
        public async Task<IActionResult> AddExcelRow([FromBody] AddRowRequestDto addRowDto)
        {
            try
            {
                if (addRowDto == null)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Satýr ekleme verisi gerekli." 
                    });
                }

                if (string.IsNullOrEmpty(addRowDto.FileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý zorunludur." 
                    });
                }

                if (string.IsNullOrEmpty(addRowDto.SheetName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sheet adý zorunludur." 
                    });
                }

                if (addRowDto.RowData == null || !addRowDto.RowData.Any())
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Satýr verisi boþ olamaz." 
                    });
                }

                _logger.LogInformation("AddExcelRow çaðrýldý: FileName={FileName}, SheetName={SheetName}, DataCount={DataCount}", 
                    addRowDto.FileName, addRowDto.SheetName, addRowDto.RowData?.Count ?? 0);

                var result = await _excelService.AddExcelRowAsync(addRowDto.FileName, addRowDto.SheetName, addRowDto.RowData, addRowDto.AddedBy);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    message = "Yeni satýr baþarýyla eklendi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Yeni satýr eklenirken hata: {FileName}", addRowDto?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Yeni satýr eklenirken hata oluþtu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Veri silme - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpDelete("data/{id}")]
        public async Task<IActionResult> DeleteExcelData(int id, [FromQuery] string? deletedBy = null)
        {
            try
            {
                if (id <= 0)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Geçerli bir ID gerekli." 
                    });
                }

                _logger.LogInformation("DeleteExcelData çaðrýldý: Id={Id}, DeletedBy={DeletedBy}", id, deletedBy);

                var result = await _excelService.DeleteExcelDataAsync(id, deletedBy);
                
                return Ok(new { 
                    success = result, 
                    message = result ? "Veri baþarýyla silindi" : "Veri bulunamadý",
                    id = id
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri silinirken hata: {Id}", id);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veri silinirken hata oluþtu: " + ex.Message,
                    id = id,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Excel verilerini belirtilen kritere göre dýþa aktarma - ÝYÝLEÞTÝRÝLDÝ
        /// </summary>
        [HttpPost("export")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExcelExportRequestDto exportRequest)
        {
            try
            {
                if (exportRequest == null || string.IsNullOrEmpty(exportRequest.FileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Export için dosya adý gerekli." 
                    });
                }

                _logger.LogInformation("Excel export isteði alýndý: {FileName}, Sheet: {Sheet}, RowIds: {RowIdCount}, IncludeHistory: {IncludeModificationHistory}", 
                    exportRequest.FileName, exportRequest.SheetName, exportRequest.RowIds?.Count ?? 0, exportRequest.IncludeModificationHistory);

                var fileBytes = await _excelService.ExportToExcelAsync(exportRequest);
                
                if (fileBytes == null || fileBytes.Length == 0)
                {
                    return StatusCode(500, new { 
                        success = false, 
                        message = "Excel dosyasý oluþturulamadý" 
                    });
                }

                var fileName = $"{exportRequest.FileName}_export_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                
                _logger.LogInformation("Excel export baþarýlý: {FileName}, {FileSize} bytes", fileName, fileBytes.Length);

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel export edilirken hata: {FileName}", exportRequest?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Excel export edilirken hata oluþtu: " + ex.Message,
                    fileName = exportRequest?.FileName,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Dosyadaki sayfalarý (sheet'leri) getirme
        /// </summary>
        [HttpGet("sheets/{fileName}")]
        public async Task<IActionResult> GetSheets(string fileName)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);
                
                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli." 
                    });
                }

                var sheets = await _excelService.GetSheetsAsync(fileName);
                return Ok(new { 
                    success = true, 
                    data = sheets,
                    fileName = fileName,
                    count = sheets.Count,
                    message = "Dosyadaki sheet'ler baþarýyla getirildi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sheet'ler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Sheet'ler getirilirken hata oluþtu: " + ex.Message,
                    fileName = fileName
                });
            }
        }

        /// <summary>
        /// Veritabanýndaki verilerin istatistiklerini alma
        /// </summary>
        [HttpGet("statistics/{fileName}")]
        public async Task<IActionResult> GetDataStatistics(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);
                
                // SheetName'i normalize et - "undefined" ise null yap
                var normalizedSheetName = NormalizeSheetName(sheetName);
                
                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli." 
                    });
                }

                var statistics = await _excelService.GetDataStatisticsAsync(fileName, normalizedSheetName);
                return Ok(new { 
                    success = true, 
                    data = statistics,
                    fileName = fileName,
                    sheetName = normalizedSheetName,
                    message = "Ýstatistikler baþarýyla getirildi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Ýstatistikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Ýstatistikler getirilirken hata oluþtu: " + ex.Message,
                    fileName = fileName
                });
            }
        }

        /// <summary>
        /// Excel dosyasý silme
        /// </summary>
        [HttpDelete("files/{fileName}")]
        public async Task<IActionResult> DeleteFile(string fileName, [FromQuery] string? deletedBy = null)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);
                
                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli." 
                    });
                }

                var result = await _excelService.DeleteExcelFileAsync(fileName, deletedBy);
                if (result)
                {
                    return Ok(new { 
                        success = true, 
                        message = "Dosya baþarýyla silindi",
                        fileName = fileName
                    });
                }
                else
                {
                    return NotFound(new { 
                        success = false, 
                        message = "Dosya bulunamadý",
                        fileName = fileName
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya silinirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Dosya silinirken hata oluþtu: " + ex.Message,
                    fileName = fileName
                });
            }
        }

        /// <summary>
        /// Debug endpoint - Dosya durumunu kontrol etme
        /// </summary>
        [HttpGet("files/{fileName}/status")]
        public async Task<IActionResult> GetFileStatus(string fileName)
        {
            try
            {
                // URL decode iþlemi
                fileName = Uri.UnescapeDataString(fileName);

                var file = await _context.ExcelFiles
                    .FirstOrDefaultAsync(f => f.FileName == fileName);

                if (file == null)
                {
                    return NotFound(new
                    {
                        success = false,
                        message = "Dosya bulunamadý",
                        fileName = fileName,
                        availableFiles = await _context.ExcelFiles
                            .Where(f => f.IsActive)
                            .Select(f => new { f.FileName, f.OriginalFileName })
                            .ToListAsync()
                    });
                }

                var dataRowsCount = await _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted)
                    .CountAsync();

                var sheets = await _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted)
                    .Select(r => r.SheetName)
                    .Distinct()
                    .ToListAsync();

                // Physical file kontrolü
                bool physicalFileExists = !string.IsNullOrEmpty(file.FilePath) && System.IO.File.Exists(file.FilePath);

                // Eðer dosya Excel formatýndaysa sheet'leri de al
                List<string> availableSheets = new List<string>();
                if (physicalFileExists)
                {
                    try
                    {
                        availableSheets = await _excelService.GetSheetsAsync(fileName);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Dosya sheet'leri alýnýrken hata: {FileName}", fileName);
                    }
                }

                return Ok(new
                {
                    success = true,
                    fileStatus = new
                    {
                        fileName = file.FileName,
                        originalFileName = file.OriginalFileName,
                        isActive = file.IsActive,
                        uploadDate = file.UploadDate,
                        fileSize = file.FileSize,
                        hasData = dataRowsCount > 0,
                        totalDataRows = dataRowsCount,
                        availableSheets = sheets,
                        physicalFileExists = physicalFileExists,
                        filePath = file.FilePath,
                        sheetsInFile = availableSheets,
                        needsReading = dataRowsCount == 0 && physicalFileExists
                    },
                    message = "Dosya durumu baþarýyla alýndý",
                    recommendations = GetFileRecommendations(file.IsActive, physicalFileExists, dataRowsCount > 0, availableSheets.Any())
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya durumu kontrol edilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Dosya durumu kontrol edilirken hata oluþtu: " + ex.Message,
                    fileName = fileName
                });
            }
        }

        private object GetFileRecommendations(bool isActive, bool physicalExists, bool hasData, bool hasSheets)
        {
            var recommendations = new List<string>();

            if (!isActive)
                recommendations.Add("Dosya aktif deðil. Dosyayý yeniden yükleyin.");
            else if (!physicalExists)
                recommendations.Add("Fiziksel dosya bulunamadý. Dosyayý yeniden yükleyin.");
            else if (!hasData)
                recommendations.Add("Dosya var ancak veri yok. Dosyayý okuyun.");
            else if (!hasSheets)
                recommendations.Add("Veritabanýnda sheet bilgisi yok. Dosyayý yeniden okuyun.");
            else
                recommendations.Add("Dosya durumu normal.");

            return recommendations;
        }

        /// <summary>
        /// Debug endpoint - Deneme amaçlý
        /// </summary>
        [HttpGet("debug/test")]
        public IActionResult DebugTest()
        {
            return Ok(new { success = true, message = "Debug test baþarýlý" });
        }

        /// <summary>
        /// Debug endpoint - Dosya adý problemlerini tespit ve düzelt
        /// </summary>
        [HttpGet("debug/fix-filename-issues")]
        public async Task<IActionResult> FixFilenameIssues()
        {
            try
            {
                var issues = new List<object>();
                var fixes = new List<object>();

                // Tüm dosyalarý kontrol et
                var allFiles = await _context.ExcelFiles.ToListAsync();
                
                foreach (var file in allFiles)
                {
                    var fileIssues = new List<string>();
                    var originalFileName = file.FileName;
                    
                    // Çift uzantý kontrolü
                    if (file.FileName.EndsWith(".xlsx.xlsx") || file.FileName.EndsWith(".xls.xls"))
                    {
                        fileIssues.Add("Çift uzantý problemi");
                        
                        // Düzeltme yap
                        var correctedName = file.FileName.Replace(".xlsx.xlsx", ".xlsx").Replace(".xls.xls", ".xls");
                        
                        // Fiziksel dosyayý yeniden adlandýr
                        if (!string.IsNullOrEmpty(file.FilePath) && System.IO.File.Exists(file.FilePath))
                        {
                            var directory = Path.GetDirectoryName(file.FilePath);
                            if (!string.IsNullOrEmpty(directory))
                            {
                                var newFilePath = Path.Combine(directory, correctedName);
                                
                                if (!System.IO.File.Exists(newFilePath))
                                {
                                    System.IO.File.Move(file.FilePath, newFilePath);
                                    file.FilePath = newFilePath;
                                    file.FileName = correctedName;
                                    
                                    fixes.Add(new
                                    {
                                        originalFileName = originalFileName,
                                        correctedFileName = correctedName,
                                        originalPath = file.FilePath,
                                        correctedPath = newFilePath,
                                        action = "Dosya yeniden adlandýrýldý"
                                    });
                                }
                            }
                        }
                        else
                        {
                            // Sadece veritabanýnda düzelt
                            file.FileName = correctedName;
                            fixes.Add(new
                            {
                                originalFileName = originalFileName,
                                correctedFileName = correctedName,
                                action = "Sadece veritabanýnda düzeltildi (fiziksel dosya yok)"
                            });
                        }
                    }
                    
                    // Fiziksel dosya var mý kontrolü
                    if (!string.IsNullOrEmpty(file.FilePath) && !System.IO.File.Exists(file.FilePath))
                    {
                        fileIssues.Add("Fiziksel dosya bulunamadý");
                    }
                    
                    // Dosya adýnda geçersiz karakterler var mý kontrolü
                    var invalidChars = Path.GetInvalidFileNameChars();
                    if (file.FileName.Any(c => invalidChars.Contains(c)))
                    {
                        fileIssues.Add("Geçersiz karakter içeriyor");
                    }
                    
                    if (fileIssues.Any())
                    {
                        issues.Add(new
                        {
                            fileName = originalFileName,
                            originalFileName = file.OriginalFileName,
                            filePath = file.FilePath,
                            isActive = file.IsActive,
                            issues = fileIssues
                        });
                    }
                }

                // Deðiþiklikleri kaydet
                if (fixes.Any())
                {
                    await _context.SaveChangesAsync();
                }

                return Ok(new
                {
                    success = true,
                    summary = new
                    {
                        totalFiles = allFiles.Count,
                        filesWithIssues = issues.Count,
                        fixesApplied = fixes.Count
                    },
                    issues = issues,
                    fixes = fixes,
                    message = fixes.Any() ? $"{fixes.Count} dosya problemi düzeltildi" : "Hiçbir problem bulunamadý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya adý problemleri düzeltilirken hata");
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message
                });
            }
        }

        /// <summary>
        /// Debug endpoint - Son yüklenen dosyalarý listele
        /// </summary>
        [HttpGet("debug/recent-uploads")]
        public async Task<IActionResult> GetRecentUploads()
        {
            try
            {
                var recentFiles = await _context.ExcelFiles
                    .Where(f => f.IsActive)
                    .OrderByDescending(f => f.UploadDate)
                    .Take(10)
                    .Select(f => new
                    {
                        f.FileName,
                        f.OriginalFileName,
                        f.UploadDate,
                        f.FileSize,
                        f.IsActive,
                        PhysicalFileExists = !string.IsNullOrEmpty(f.FilePath) && System.IO.File.Exists(f.FilePath),
                        f.FilePath,
                        HasData = _context.ExcelDataRows.Any(r => r.FileName == f.FileName && !r.IsDeleted)
                    })
                    .ToListAsync();

                return Ok(new
                {
                    success = true,
                    recentUploads = recentFiles,
                    message = "Son yüklenen dosyalar"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Son yüklemeler alýnýrken hata");
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