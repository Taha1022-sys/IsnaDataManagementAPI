using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Services;
using ExcelDataManagementAPI.Models.DTOs;

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
                    "Tek sheet okuma: POST /api/excel/read/{fileName}?sheetName=SheetAdý",
                    "TÜM sheet'leri okuma: POST /api/excel/read/{fileName} (sheetName belirtmeyin)",
                    "Zorunlu tüm sheet okuma: POST /api/excel/read-all-sheets/{fileName}",
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
                },
                newFeatures = new[]
                {
                    "? Artýk tüm sheet'ler otomatik okunuyor!",
                    "? sheetName belirtmezseniz tüm sheet'ler iþlenir",
                    "? Her sheet'in verileri ayrý ayrý veritabanýnda saklanýyor",
                    "? Frontend'de artýk tüm sheet'lerin verileri görünecek"
                }
            });
        }

        [HttpGet("debug/database-status")]
        public async Task<IActionResult> GetDatabaseStatus()
        {
            try
            {
                var filesCount = await _excelService.GetExcelFilesAsync();
                var totalDataRows = await _context.ExcelDataRows.CountAsync();
                var activeDataRows = await _context.ExcelDataRows.Where(r => !r.IsDeleted).CountAsync();
                var deletedDataRows = await _context.ExcelDataRows.Where(r => r.IsDeleted).CountAsync();

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

        [HttpGet("files")]
        public async Task<IActionResult> GetFiles()
        {
            try
            {
                var files = await _excelService.GetExcelFilesAsync();
                return Ok(new { 
                    success = true, 
                    data = files,
                    message = "Yüklü Excel dosyalarý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosyalar getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

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

                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(request.ExcelFile.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalarý (.xlsx, .xls) desteklenir" 
                    });
                }

                var uploadedFile = await _excelService.UploadExcelFileAsync(request.ExcelFile, request.ProcessedBy);
                
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

                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(request.ExcelFile.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalarý (.xlsx, .xls) desteklenir" 
                    });
                }

                var uploadedFile = await _excelService.UploadExcelFileAsync(request.ExcelFile, request.UpdatedBy);
                
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

        [HttpPost("read/{fileName}")]
        public async Task<IActionResult> ReadExcelData(string fileName, [FromQuery] string? sheetName = null, [FromQuery] bool forceReread = false)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("ReadExcelData çaðrýldý: FileName={FileName}, SheetName={SheetName}, ForceReread={ForceReread}", fileName, sheetName, forceReread);

                if (!forceReread)
                {
                    var query = _context.ExcelDataRows.Where(r => r.FileName == fileName && !r.IsDeleted);
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        query = query.Where(r => r.SheetName == sheetName);
                    }

                    var existingDataCount = await query.CountAsync();

                    if (existingDataCount > 0)
                    {
                        _logger.LogInformation("Veriler zaten mevcut, veritabanýndan döndürülüyor: {FileName}, Sheet: {SheetName}", fileName, sheetName);
                        
                        var existingData = await _excelService.GetAllExcelDataAsync(fileName, sheetName);
                        
                        object additionalInfo = new { };
                        if (string.IsNullOrEmpty(sheetName))
                        {
                            var sheetGroups = existingData.GroupBy(d => d.SheetName)
                                .ToDictionary(g => g.Key, g => g.Count());
                            additionalInfo = new
                            {
                                totalSheets = sheetGroups.Count,
                                sheetSummary = sheetGroups,
                                processedSheets = sheetGroups.Keys.ToArray()
                            };
                        }

                        return Ok(new { 
                            success = true, 
                            data = existingData,
                            fileName = fileName,
                            sheetName = sheetName,
                            totalRows = existingData.Count,
                            message = "Mevcut veriler döndürüldü (Excel dosyasý yeniden okunmadý)",
                            isFromDatabase = true,
                            additionalInfo = additionalInfo
                        });
                    }
                }

                var data = await _excelService.ReadExcelDataAsync(fileName, sheetName);
                
                object resultAdditionalInfo = new { };
                if (string.IsNullOrEmpty(sheetName))
                {
                    var sheetGroups = data.GroupBy(d => d.SheetName)
                        .ToDictionary(g => g.Key, g => g.Count());
                    resultAdditionalInfo = new
                    {
                        totalSheets = sheetGroups.Count,
                        sheetSummary = sheetGroups,
                        processedSheets = sheetGroups.Keys.ToArray()
                    };
                }

                return Ok(new { 
                    success = true, 
                    data = data,
                    fileName = fileName,
                    sheetName = sheetName,
                    totalRows = data.Count,
                    message = forceReread ? "Excel dosyasý zorla yeniden okundu" : 
                             string.IsNullOrEmpty(sheetName) ? "Excel dosyasýndaki tüm sheet'ler baþarýyla okundu" : "Excel sheet'i baþarýyla okundu ve veritabanýna aktarýldý",
                    isFromDatabase = false,
                    additionalInfo = resultAdditionalInfo
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyasý okunurken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("read-all-sheets/{fileName}")]
        public async Task<IActionResult> ReadAllSheetsFromExcel(string fileName, [FromQuery] string? processedBy = null)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("ReadAllSheetsFromExcel çaðrýldý: FileName={FileName}", fileName);

                var allSheetsData = await _excelService.ReadAllSheetsFromExcelAsync(fileName);
                
                var totalRows = allSheetsData.Values.Sum(sheetData => sheetData.Count);
                var sheetSummary = allSheetsData.ToDictionary(
                    kvp => kvp.Key, 
                    kvp => kvp.Value.Count
                );

                return Ok(new { 
                    success = true, 
                    data = allSheetsData,
                    fileName = fileName,
                    totalSheets = allSheetsData.Count,
                    totalRows = totalRows,
                    sheetSummary = sheetSummary,
                    message = $"Dosyadaki {allSheetsData.Count} sheet baþarýyla okundu ve veritabanýna aktarýldý. Toplam {totalRows} kayýt.",
                    processedSheets = allSheetsData.Keys.ToArray()
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Tüm sheet'ler okunurken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("force-reread/{fileName}")]
        public async Task<IActionResult> ForceRereadExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogWarning("FORCE REREAD çaðrýldý - mevcut deðiþiklikler kaybolacak: {FileName}", fileName);

                var data = await _excelService.ReadExcelDataAsync(fileName, sheetName);
                return Ok(new { 
                    success = true, 
                    data = data,
                    fileName = fileName,
                    sheetName = sheetName,
                    totalRows = data.Count,
                    message = "?? Excel dosyasý zorla yeniden okundu - önceki deðiþiklikler kayboldu!",
                    warning = "Bu iþlem mevcut tüm deðiþiklikleri sildi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyasý zorla okunurken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("data/{fileName}")]
        public async Task<IActionResult> GetExcelData(string fileName, [FromQuery] string? sheetName = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("Veri getirilmeye çalýþýlýyor: FileName={FileName}, SheetName={SheetName}, Page={Page}", fileName, sheetName, page);

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
                            .Select(f => f.FileName)
                            .ToListAsync(),
                        suggestion = "Dosyayý yeniden yükleyin veya listeden doðru dosyayý seçin"
                    });
                }

                var data = await _excelService.GetExcelDataAsync(fileName, sheetName, page, pageSize);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);

                if (data.Count == 0)
                {
                    var totalDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .CountAsync();

                    if (totalDataCount == 0)
                    {
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
                            statistics = statistics,
                            hasData = false,
                            availableSheets = availableSheets,
                            message = "Dosya bulundu ancak henüz okunmamýþ. Dosyayý okumak için read endpoint'ini kullanýn.",
                            suggestedActions = new
                            {
                                readFile = $"/api/excel/read/{fileName}",
                                readWithSheet = availableSheets.Any() ? $"/api/excel/read/{fileName}?sheetName={availableSheets.First()}" : null,
                                note = "read endpoint'i mevcut verileri kontrol eder ve gerekirse okur"
                            }
                        });
                    }
                    
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

                return Ok(new { 
                    success = true, 
                    data = data, 
                    fileName = fileName,
                    sheetName = sheetName,
                    page = page, 
                    pageSize = pageSize,
                    statistics = statistics,
                    hasData = data.Count > 0,
                    message = data.Count > 0 ? $"Sayfa {page} - {data.Count} kayýt gösteriliyor" : "Bu sayfada gösterilecek veri bulunamadý",
                    dataSource = "database" 
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}", fileName, sheetName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("data/{fileName}/all")]
        public async Task<IActionResult> GetAllExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("Tüm veriler getirilmeye çalýþýlýyor: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

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
                            .Select(f => f.FileName)
                            .ToListAsync()
                    });
                }

                var data = await _excelService.GetAllExcelDataAsync(fileName, sheetName);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);

                if (data.Count == 0)
                {
                    var totalDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .CountAsync();

                    if (totalDataCount == 0)
                    {
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
                            message = "Dosya bulundu ancak henüz okunmamýþ. Dosyayý okumak için read endpoint'ini kullanýn.",
                            suggestedActions = new
                            {
                                readFile = $"/api/excel/read/{fileName}",
                                readWithSheet = availableSheets.Any() ? $"/api/excel/read/{fileName}?sheetName={availableSheets.First()}" : null,
                                note = "read endpoint'i mevcut verileri kontrol eder ve gerekirse okur"
                            }
                        });
                    }
                    
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
                    data = data, 
                    fileName = fileName,
                    sheetName = sheetName,
                    totalRows = data.Count,
                    statistics = statistics,
                    hasData = data.Count > 0,
                    message = data.Count > 0 ? $"Toplam {data.Count} kayýt getirildi" : "Bu dosya/sayfa için veri bulunamadý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}", fileName, sheetName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veriler getirilirken bir hata oluþtu: " + ex.Message,
                    fileName = fileName,
                    sheetName = sheetName
                });
            }
        }

        [HttpPut("data")]
        public async Task<IActionResult> UpdateExcelData([FromBody] ExcelDataUpdateDto updateDto)
        {
            try
            {
                if (updateDto.Id <= 0)
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Geçerli bir ID belirtmeniz gerekiyor",
                        providedId = updateDto.Id
                    });
                }

                if (updateDto.Data == null || !updateDto.Data.Any())
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Güncellenecek veri boþ olamaz"
                    });
                }

                _logger.LogInformation("Veri güncelleme isteði: ID={Id}, ModifiedBy={ModifiedBy}", updateDto.Id, updateDto.ModifiedBy);

                var result = await _excelService.UpdateExcelDataAsync(updateDto, HttpContext);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    message = "Veri baþarýyla güncellendi",
                    updatedFields = updateDto.Data.Keys.ToArray(),
                    version = result.Version,
                    modifiedDate = result.ModifiedDate
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri güncellenirken hata: ID={Id}, ModifiedBy={ModifiedBy}", updateDto.Id, updateDto.ModifiedBy);
                
                if (ex.Message.Contains("bulunamadý"))
                {
                    return NotFound(new { 
                        success = false, 
                        message = "Güncellenecek veri bulunamadý. Veri silinmiþ olabilir.",
                        id = updateDto.Id
                    });
                }
                else if (ex.Message.Contains("silinmiþ"))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Bu veri silinmiþ durumda ve güncellenemez.",
                        id = updateDto.Id
                    });
                }
                else if (ex.Message.Contains("eþzamanlýlýk"))
                {
                    return Conflict(new { 
                        success = false, 
                        message = "Bu veri baþka bir kullanýcý tarafýndan deðiþtirilmiþ. Sayfayý yenileyip tekrar deneyin.",
                        id = updateDto.Id
                    });
                }
                
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veri güncellenirken beklenmeyen bir hata oluþtu: " + ex.Message,
                    id = updateDto.Id
                });
            }
        }

        [HttpPut("data/bulk")]
        public async Task<IActionResult> BulkUpdateExcelData([FromBody] BulkUpdateDto bulkUpdateDto)
        {
            try
            {
                var results = await _excelService.BulkUpdateExcelDataAsync(bulkUpdateDto, HttpContext);
                return Ok(new { success = true, data = results, count = results.Count });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Toplu veri güncellenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("data")]
        public async Task<IActionResult> AddExcelRow([FromBody] AddRowRequestDto addRowDto)
        {
            try
            {
                var result = await _excelService.AddExcelRowAsync(addRowDto.FileName, addRowDto.SheetName, addRowDto.RowData, addRowDto.AddedBy, HttpContext);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Yeni satýr eklenirken hata: {FileName}", addRowDto.FileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpDelete("data/{id}")]
        public async Task<IActionResult> DeleteExcelData(int id, [FromQuery] string? deletedBy = null)
        {
            try
            {
                var result = await _excelService.DeleteExcelDataAsync(id, deletedBy, HttpContext);
                return Ok(new { success = result, message = result ? "Veri silindi" : "Veri bulunamadý" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri silinirken hata: {Id}", id);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("recent-changes")]
        public async Task<IActionResult> GetRecentChanges(
            [FromQuery] int hours = 24, 
            [FromQuery] string? fileName = null,
            [FromQuery] string? sheetName = null,
            [FromQuery] int limit = 100)
        {
            try
            {
                var query = _context.GerceklesenRaporlar.AsQueryable();

                var fromDate = DateTime.UtcNow.AddHours(-hours);
                query = query.Where(r => r.ChangeDate >= fromDate);

                if (!string.IsNullOrEmpty(fileName))
                    query = query.Where(r => r.FileName == fileName);

                if (!string.IsNullOrEmpty(sheetName))
                    query = query.Where(r => r.SheetName == sheetName);

                var recentChanges = await query
                    .OrderByDescending(r => r.ChangeDate)
                    .Take(limit)
                    .Select(r => new
                    {
                        r.Id,
                        r.FileName,
                        r.SheetName,
                        r.RowIndex,
                        r.OperationType,
                        r.ChangeDate,
                        r.ModifiedBy,
                        r.UserIP,
                        r.ChangeReason,
                        MinutesAgo = EF.Functions.DateDiffMinute(r.ChangeDate, DateTime.UtcNow),
                        OldValuePreview = r.OldValue != null && r.OldValue.Length > 200 ? 
                                        r.OldValue.Substring(0, 200) + "..." : r.OldValue,
                        NewValuePreview = r.NewValue != null && r.NewValue.Length > 200 ? 
                                        r.NewValue.Substring(0, 200) + "..." : r.NewValue,
                        r.ChangedColumns,
                        r.IsSuccess,
                        r.ErrorMessage
                    })
                    .ToListAsync();

                return Ok(new
                {
                    success = true,
                    data = recentChanges,
                    summary = new
                    {
                        totalChanges = recentChanges.Count,
                        timeRange = $"Son {hours} saat",
                        filterFileName = fileName,
                        filterSheetName = sheetName,
                        queryTime = DateTime.UtcNow,
                        operationTypes = recentChanges.GroupBy(x => x.OperationType)
                                                   .ToDictionary(g => g.Key, g => g.Count())
                    },
                    message = $"Son {hours} saatteki {recentChanges.Count} deðiþiklik"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Son deðiþiklikler getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("today-changes")]
        public async Task<IActionResult> GetTodayChanges()
        {
            try
            {
                var today = DateTime.Today;
                var tomorrow = today.AddDays(1);

                var todayChanges = await _context.GerceklesenRaporlar
                    .Where(r => r.ChangeDate >= today && r.ChangeDate < tomorrow)
                    .OrderByDescending(r => r.ChangeDate)
                    .Select(r => new
                    {
                        r.Id,
                        r.FileName,
                        r.SheetName,
                        r.RowIndex,
                        r.OperationType,
                        r.ChangeDate,
                        r.ModifiedBy,
                        r.UserIP,
                        r.ChangeReason,
                        OldValuePreview = r.OldValue != null && r.OldValue.Length > 150 ? 
                                        r.OldValue.Substring(0, 150) + "..." : r.OldValue,
                        NewValuePreview = r.NewValue != null && r.NewValue.Length > 150 ? 
                                        r.NewValue.Substring(0, 150) + "..." : r.NewValue,
                        r.ChangedColumns,
                        r.IsSuccess
                    })
                    .ToListAsync();

                var summary = todayChanges
                    .GroupBy(r => r.OperationType)
                    .ToDictionary(g => g.Key, g => g.Count());

                return Ok(new
                {
                    success = true,
                    data = todayChanges,
                    summary = new
                    {
                        date = today.ToString("yyyy-MM-dd"),
                        totalChanges = todayChanges.Count,
                        operationTypes = summary,
                        uniqueFiles = todayChanges.Select(r => r.FileName).Distinct().Count(),
                        uniqueUsers = todayChanges.Where(r => r.ModifiedBy != null)
                                                .Select(r => r.ModifiedBy).Distinct().Count(),
                        queryTime = DateTime.UtcNow
                    },
                    message = $"Bugünkü toplam {todayChanges.Count} deðiþiklik"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Bugünkü deðiþiklikler getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("most-active-users")]
        public async Task<IActionResult> GetMostActiveUsers([FromQuery] int days = 7)
        {
            try
            {
                var fromDate = DateTime.UtcNow.AddDays(-days);

                var activeUsers = await _context.GerceklesenRaporlar
                    .Where(r => r.ChangeDate >= fromDate && r.ModifiedBy != null)
                    .GroupBy(r => r.ModifiedBy)
                    .Select(g => new
                    {
                        User = g.Key,
                        TotalChanges = g.Count(),
                        LastActivity = g.Max(r => r.ChangeDate),
                        FirstActivity = g.Min(r => r.ChangeDate),
                        FilesModified = g.Select(r => r.FileName).Distinct().Count(),
                        SheetsModified = g.Select(r => r.SheetName).Distinct().Count(),
                        OperationTypes = g.GroupBy(x => x.OperationType).ToDictionary(og => og.Key, og => og.Count())
                    })
                    .OrderByDescending(u => u.TotalChanges)
                    .ToListAsync();

                return Ok(new
                {
                    success = true,
                    data = activeUsers,
                    summary = new
                    {
                        totalUsers = activeUsers.Count,
                        timeRange = $"Son {days} gün",
                        totalChanges = activeUsers.Sum(u => u.TotalChanges),
                        queryTime = DateTime.UtcNow
                    },
                    message = $"Son {days} gündeki en aktif kullanýcýlar"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Aktif kullanýcýlar getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("export")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExcelExportRequestDto exportRequest)
        {
            try
            {
                var fileBytes = await _excelService.ExportToExcelAsync(exportRequest);
                var fileName = $"{exportRequest.FileName}_export_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel export edilirken hata: {FileName}", exportRequest.FileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("sheets/{fileName}")]
        public async Task<IActionResult> GetSheets(string fileName)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
                var sheets = await _excelService.GetSheetsAsync(fileName);
                return Ok(new { 
                    success = true, 
                    data = sheets,
                    fileName = fileName,
                    message = "Dosyadaki sheet'ler"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sheet'ler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("statistics/{fileName}")]
        public async Task<IActionResult> GetDataStatistics(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);
                return Ok(new { success = true, data = statistics });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Ýstatistikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpDelete("files/{fileName}")]
        public async Task<IActionResult> DeleteFile(string fileName, [FromQuery] string? deletedBy = null)
        {
            try
            {
                fileName = Uri.UnescapeDataString(fileName);
                
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
                    message = ex.Message,
                    fileName = fileName
                });
            }
        }

        [HttpGet("current-status")]
        public async Task<IActionResult> GetCurrentDatabaseStatus()
        {
            try
            {
                var status = new
                {
                    timestamp = DateTime.UtcNow,
                    totalFiles = await _context.ExcelFiles.CountAsync(),
                    activeFiles = await _context.ExcelFiles.CountAsync(f => f.IsActive),
                    totalDataRows = await _context.ExcelDataRows.CountAsync(),
                    activeDataRows = await _context.ExcelDataRows.CountAsync(r => !r.IsDeleted),
                    deletedDataRows = await _context.ExcelDataRows.CountAsync(r => r.IsDeleted),
                    recentFileUploads = await _context.ExcelFiles
                        .OrderByDescending(f => f.UploadDate)
                        .Take(5)
                        .Select(f => new
                        {
                            f.FileName,
                            f.OriginalFileName,
                            f.IsActive,
                            f.UploadDate,
                            f.FileSize
                        })
                        .ToListAsync(),
                    recentActivity = await _context.GerceklesenRaporlar
                        .OrderByDescending(r => r.ChangeDate)
                        .Take(5)
                        .Select(r => new
                        {
                            r.Id,
                            r.FileName,
                            r.OperationType,
                            r.ChangeDate,
                            r.ModifiedBy,
                            r.IsSuccess
                        })
                        .ToListAsync()
                };

                return Ok(new { success = true, data = status });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veritabaný durumu alýnýrken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("dashboard")]
        public async Task<IActionResult> GetDashboardData()
        {
            try
            {
                var dashboardData = new
                {
                    overview = new
                    {
                        totalFiles = await _context.ExcelFiles.CountAsync(),
                        activeFiles = await _context.ExcelFiles.CountAsync(f => f.IsActive),
                        totalDataRows = await _context.ExcelDataRows.CountAsync(),
                        activeDataRows = await _context.ExcelDataRows.CountAsync(r => !r.IsDeleted),
                        totalChangesToday = await _context.GerceklesenRaporlar
                            .CountAsync(g => g.ChangeDate.Date == DateTime.Today),
                        totalStorageMB = Math.Round(await _context.ExcelFiles.SumAsync(f => f.FileSize) / 1024.0 / 1024.0, 2)
                    },
                    recentActivity = await _context.GerceklesenRaporlar
                        .OrderByDescending(g => g.ChangeDate)
                        .Take(10)
                        .Select(g => new
                        {
                            g.Id,
                            g.FileName,
                            g.OperationType,
                            g.ModifiedBy,
                            g.ChangeDate,
                            g.IsSuccess,
                            MinutesAgo = EF.Functions.DateDiffMinute(g.ChangeDate, DateTime.UtcNow)
                        })
                        .ToListAsync(),
                    todayStats = new
                    {
                        creates = await _context.GerceklesenRaporlar.CountAsync(r => r.OperationType == "Create" && r.ChangeDate.Date == DateTime.Today),
                        updates = await _context.GerceklesenRaporlar.CountAsync(r => r.OperationType == "Update" && r.ChangeDate.Date == DateTime.Today),
                        deletes = await _context.GerceklesenRaporlar.CountAsync(r => r.OperationType == "Delete" && r.ChangeDate.Date == DateTime.Today)
                    },
                    fileStatistics = await _context.ExcelFiles
                        .Select(f => new 
                        {
                            f.FileName,
                            f.OriginalFileName,
                            f.IsActive,
                            f.UploadDate,
                            f.FileSize,
                            DataRowCount = _context.ExcelDataRows.Count(r => r.FileName == f.FileName && !r.IsDeleted),
                            LastModified = _context.GerceklesenRaporlar.Where(r => r.FileName == f.FileName)
                                                                      .OrderByDescending(r => r.ChangeDate)
                                                                      .Select(r => r.ChangeDate)
                                                                      .FirstOrDefault(),
                            CreatedBy = _context.GerceklesenRaporlar.Where(r => r.FileName == f.FileName)
                                                                     .Select(r => r.ModifiedBy)
                                                                     .FirstOrDefault()
                        })
                        .ToListAsync()
                };

                return Ok(new { success = true, data = dashboardData });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dashboard verileri getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}