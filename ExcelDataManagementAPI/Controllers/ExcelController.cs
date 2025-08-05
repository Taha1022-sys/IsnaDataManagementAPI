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

                var result = await _excelService.UploadExcelFileAsync(file, uploadedBy);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    message = "Dosya baþarýyla yüklendi! Þimdi iþlem yapmak için dosyayý kullanabilirsiniz.",
                    nextSteps = new 
                    {
                        readData = $"/api/excel/read/{result.FileName}",
                        getData = $"/api/excel/data/{result.FileName}",
                        getSheets = $"/api/excel/sheets/{result.FileName}"
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya yüklenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
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

                // Önce dosyayý yükle
                var uploadedFile = await _excelService.UploadExcelFileAsync(request.ExcelFile, request.ProcessedBy);
                
                // Sonra oku
                var data = await _excelService.ReadExcelDataAsync(uploadedFile.FileName, request.SheetName);
                
                return Ok(new { 
                    success = true, 
                    data = data,
                    fileName = uploadedFile.FileName,
                    originalFileName = uploadedFile.OriginalFileName,
                    sheetName = request.SheetName,
                    totalRows = data.Count,
                    message = "Excel dosyasý baþarýyla okundu ve veritabanýna aktarýldý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya okunurken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
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

                // Önce dosyayý yükle
                var uploadedFile = await _excelService.UploadExcelFileAsync(request.ExcelFile, request.UpdatedBy);
                
                // Dosyayý oku
                var data = await _excelService.ReadExcelDataAsync(uploadedFile.FileName, request.SheetName);
                
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
                
                var data = await _excelService.ReadExcelDataAsync(fileName, sheetName);
                return Ok(new { 
                    success = true, 
                    data = data,
                    fileName = fileName,
                    sheetName = sheetName,
                    totalRows = data.Count,
                    message = "Excel verisi baþarýyla okundu ve veritabanýna aktarýldý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyasý okunurken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
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
                
                _logger.LogInformation("GetExcelData çaðrýldý: FileName={FileName}, SheetName={SheetName}, Page={Page}, PageSize={PageSize}", 
                    fileName, sheetName, page, pageSize);

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
                var data = await _excelService.GetExcelDataAsync(fileName, sheetName, page, pageSize);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);

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
                            sheetName = sheetName,
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
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        var sheetExists = await _context.ExcelDataRows
                            .AnyAsync(r => r.FileName == fileName && r.SheetName == sheetName && !r.IsDeleted);

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
                                message = $"Belirtilen sayfa '{sheetName}' bulunamadý.",
                                fileName = fileName,
                                requestedSheet = sheetName,
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
                    sheetName = sheetName,
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
                
                _logger.LogInformation("GetAllExcelData çaðrýldý: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

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

                var data = await _excelService.GetAllExcelDataAsync(fileName, sheetName);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);

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
                            sheetName = sheetName,
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
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        var availableSheets = await _context.ExcelDataRows
                            .Where(r => r.FileName == fileName && !r.IsDeleted)
                            .Select(r => r.SheetName)
                            .Distinct()
                            .ToListAsync();

                        return NotFound(new
                        {
                            success = false,
                            message = $"Belirtilen sayfa '{sheetName}' bulunamadý.",
                            fileName = fileName,
                            requestedSheet = sheetName,
                            availableSheets = availableSheets,
                            suggestion = availableSheets.Any() ? $"Mevcut sayfalardan birini seçin: {string.Join(", ", availableSheets)}" : "Bu dosyada henüz veri bulunmuyor."
                        });
                    }
                }
                
                return Ok(new { 
                    success = true, 
                    data = data ?? new List<ExcelDataResponseDto>(), 
                    fileName = fileName,
                    sheetName = sheetName,
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

                _logger.LogInformation("Excel export isteði alýndý: {FileName}, Sheet: {Sheet}, RowIds: {RowIdCount}, IncludeHistory: {IncludeHistory}", 
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
                
                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya adý gerekli." 
                    });
                }

                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);
                return Ok(new { 
                    success = true, 
                    data = statistics,
                    fileName = fileName,
                    sheetName = sheetName,
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
    }
}