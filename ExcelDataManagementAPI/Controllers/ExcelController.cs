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
                message = "Excel API �al���yor!", 
                timestamp = DateTime.Now,
                availableOperations = new[]
                {
                    "Excel dosyas� y�kleme: POST /api/excel/upload",
                    "Dosya listesi: GET /api/excel/files", 
                    "Dosya silme: DELETE /api/excel/files/{fileName}",
                    "Manuel dosya se�ip okuma: POST /api/excel/read-from-file",
                    "Manuel dosya se�ip g�ncelleme: POST /api/excel/update-from-file",
                    "Veri d�zenleme: PUT /api/excel/data",
                    "T�m verileri getirme: GET /api/excel/data/{fileName}/all",
                    "Sayfal� veri getirme: GET /api/excel/data/{fileName}?page=1&pageSize=50"
                },
                debugEndpoints = new[]
                {
                    "Veritaban� durumu: GET /api/excel/debug/database-status",
                    "Dosya durumu: GET /api/excel/files/{fileName}/status",
                    "Veri sorgusu test: GET /api/excel/debug/test-data-query/{fileName}",
                    "Veri ak�� testi: GET /api/excel/debug/data-flow-test/{fileName}"
                }
            });
        }

        /// <summary>
        /// Debug endpoint - Veritaban� durumunu kontrol etme
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

                // Her dosya i�in detayl� bilgi
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
                    message = "Veritaban� durumu ba�ar�yla al�nd�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veritaban� durumu kontrol edilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Y�kl� Excel dosyalar�n�n listesi
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
                    message = "Y�kl� Excel dosyalar�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosyalar getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Excel dosyas� y�kleme - Bilgisayardan dosya se�imi
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
                        message = "L�tfen bilgisayar�n�zdan bir Excel dosyas� se�in (.xlsx veya .xls)" 
                    });

                // Dosya uzant�s� kontrol�
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalar� (.xlsx, .xls) desteklenir" 
                    });
                }

                var result = await _excelService.UploadExcelFileAsync(file, uploadedBy);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    message = "Dosya ba�ar�yla y�klendi! �imdi i�lem yapmak i�in dosyay� kullanabilirsiniz.",
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
                _logger.LogError(ex, "Dosya y�klenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Bilgisayardan Excel dosyas� se�ip direkt okuma
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
                        message = "L�tfen bilgisayar�n�zdan bir Excel dosyas� se�in" 
                    });

                // Dosya uzant�s� kontrol�
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(request.ExcelFile.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalar� (.xlsx, .xls) desteklenir" 
                    });
                }

                // �nce dosyay� y�kle
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
                    message = "Excel dosyas� ba�ar�yla okundu ve veritaban�na aktar�ld�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya okunurken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Bilgisayardan Excel dosyas� se�ip g�ncelleme
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
                        message = "L�tfen bilgisayar�n�zdan bir Excel dosyas� se�in" 
                    });

                // Dosya uzant�s� kontrol�
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(request.ExcelFile.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sadece Excel dosyalar� (.xlsx, .xls) desteklenir" 
                    });
                }

                // �nce dosyay� y�kle
                var uploadedFile = await _excelService.UploadExcelFileAsync(request.ExcelFile, request.UpdatedBy);
                
                // Dosyay� oku
                var data = await _excelService.ReadExcelDataAsync(uploadedFile.FileName, request.SheetName);
                
                return Ok(new { 
                    success = true, 
                    uploadedFile = uploadedFile,
                    data = data,
                    totalRows = data.Count,
                    message = "Excel dosyas� y�klendi ve okundu. �imdi g�ncellemeler yapabilirsiniz.",
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
                _logger.LogError(ex, "Dosya g�ncellenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Y�kl� dosyadan veri okuma ve veritaban�na aktarma
        /// </summary>
        [HttpPost("read/{fileName}")]
        public async Task<IActionResult> ReadExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                var data = await _excelService.ReadExcelDataAsync(fileName, sheetName);
                return Ok(new { 
                    success = true, 
                    data = data,
                    fileName = fileName,
                    sheetName = sheetName,
                    totalRows = data.Count,
                    message = "Excel verisi ba�ar�yla okundu ve veritaban�na aktar�ld�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyas� okunurken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Dosyadan veri getirme (sayfalama ile) - �Y�LE�T�R�LD�
        /// </summary>
        [HttpGet("data/{fileName}")]
        public async Task<IActionResult> GetExcelData(string fileName, [FromQuery] string? sheetName = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("GetExcelData �a�r�ld�: FileName={FileName}, SheetName={SheetName}, Page={Page}, PageSize={PageSize}", 
                    fileName, sheetName, page, pageSize);

                // Parametre validasyonu
                if (page < 1) page = 1;
                if (pageSize < 1 || pageSize > 1000) pageSize = 50;

                // �nce dosyan�n var olup olmad���n� kontrol et
                var fileExists = await _context.ExcelFiles
                    .AnyAsync(f => f.FileName == fileName && f.IsActive);

                if (!fileExists)
                {
                    _logger.LogWarning("Dosya bulunamad�: {FileName}", fileName);
                    return NotFound(new
                    {
                        success = false,
                        message = "Belirtilen dosya bulunamad�. Dosyan�n y�klendi�inden ve aktif oldu�undan emin olun.",
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
                _logger.LogInformation("Service'den d�nen veri say�s�: {Count}", data?.Count ?? 0);

                // E�er veri yoksa ancak dosya varsa, dosyan�n okunup okunmad���n� kontrol et
                if (data == null || data.Count == 0)
                {
                    var totalDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .CountAsync();

                    if (totalDataCount == 0)
                    {
                        // Dosya var ama veri yok - okunmam�� olabilir
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
                            message = "Dosya bulundu ancak hen�z okunmam��. �nce dosyay� okuyun.",
                            suggestedActions = new
                            {
                                readFile = $"/api/excel/read/{fileName}",
                                readWithSheet = availableSheets.Any() ? $"/api/excel/read/{fileName}?sheetName={availableSheets.First()}" : null
                            }
                        });
                    }
                    
                    // Sheet belirtilmi�se ve o sheet'te veri yoksa
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
                                message = $"Belirtilen sayfa '{sheetName}' bulunamad�.",
                                fileName = fileName,
                                requestedSheet = sheetName,
                                availableSheets = availableSheets,
                                suggestion = availableSheets.Any() ? $"Mevcut sayfalardan birini se�in: {string.Join(", ", availableSheets)}" : "Bu dosyada hen�z veri bulunmuyor."
                            });
                        }
                    }
                }

                // Ba�ar�l� response
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
                        ? $"Sayfa {page}/{totalPages} - {data.Count} kay�t g�steriliyor" 
                        : "Bu sayfada g�sterilecek veri bulunamad�"
                };

                _logger.LogInformation("Response haz�rland�: TotalRows={TotalRows}, PageCount={PageCount}", 
                    totalRowsForPagination, data?.Count ?? 0);

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}, Page: {Page}", fileName, sheetName, page);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veriler getirilirken bir hata olu�tu: " + ex.Message,
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
        /// Dosyadan t�m verileri getirme (sayfalama olmadan) - �Y�LE�T�R�LD�
        /// </summary>
        [HttpGet("data/{fileName}/all")]
        public async Task<IActionResult> GetAllExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("GetAllExcelData �a�r�ld�: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

                // �nce dosyan�n var olup olmad���n� kontrol et
                var fileExists = await _context.ExcelFiles
                    .AnyAsync(f => f.FileName == fileName && f.IsActive);

                if (!fileExists)
                {
                    _logger.LogWarning("Dosya bulunamad�: {FileName}", fileName);
                    return NotFound(new
                    {
                        success = false,
                        message = "Belirtilen dosya bulunamad�. Dosyan�n y�klendi�inden ve aktif oldu�undan emin olun.",
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
                _logger.LogInformation("Service'den d�nen t�m veri say�s�: {Count}", data?.Count ?? 0);

                // E�er veri yoksa detayl� bilgi ver
                if (data == null || data.Count == 0)
                {
                    var totalDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && !r.IsDeleted)
                        .CountAsync();

                    if (totalDataCount == 0)
                    {
                        // Dosya var ama veri yok - okunmam�� olabilir
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
                            message = "Dosya bulundu ancak hen�z okunmam��. �nce dosyay� okuyun.",
                            suggestedActions = new
                            {
                                readFile = $"/api/excel/read/{fileName}",
                                readWithSheet = availableSheets.Any() ? $"/api/excel/read/{fileName}?sheetName={availableSheets.First()}" : null
                            }
                        });
                    }
                    
                    // Sheet belirtilmi�se ve o sheet'te veri yoksa
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
                            message = $"Belirtilen sayfa '{sheetName}' bulunamad�.",
                            fileName = fileName,
                            requestedSheet = sheetName,
                            availableSheets = availableSheets,
                            suggestion = availableSheets.Any() ? $"Mevcut sayfalardan birini se�in: {string.Join(", ", availableSheets)}" : "Bu dosyada hen�z veri bulunmuyor."
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
                    message = data != null && data.Count > 0 ? $"Toplam {data.Count} kay�t getirildi" : "Bu dosya/sayfa i�in veri bulunamad�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "T�m Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}", fileName, sheetName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veriler getirilirken bir hata olu�tu: " + ex.Message,
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
        /// Excel verisini g�ncelleme - �Y�LE�T�R�LD�
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
                        message = "Ge�ersiz g�ncelleme verisi. ID zorunludur." 
                    });
                }

                if (updateDto.Data == null || !updateDto.Data.Any())
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "G�ncellenecek veri bo� olamaz." 
                    });
                }

                _logger.LogInformation("UpdateExcelData �a�r�ld�: Id={Id}, DataCount={DataCount}", 
                    updateDto.Id, updateDto.Data?.Count ?? 0);

                var result = await _excelService.UpdateExcelDataAsync(updateDto);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    message = "Veri ba�ar�yla g�ncellendi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri g�ncellenirken hata: {Id}", updateDto?.Id);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veri g�ncellenirken hata olu�tu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Toplu Excel verisini g�ncelleme - �Y�LE�T�R�LD�
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
                        message = "Toplu g�ncelleme i�in en az bir g�ncelleme verisi gerekli." 
                    });
                }

                _logger.LogInformation("BulkUpdateExcelData �a�r�ld�: UpdateCount={Count}", bulkUpdateDto.Updates.Count);

                var results = await _excelService.BulkUpdateExcelDataAsync(bulkUpdateDto);
                
                return Ok(new { 
                    success = true, 
                    data = results, 
                    count = results.Count,
                    message = $"{results.Count} kay�t ba�ar�yla g�ncellendi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Toplu veri g�ncellenirken hata");
                return StatusCode(500, new { 
                    success = false, 
                    message = "Toplu veri g�ncellenirken hata olu�tu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Yeni bir sat�r ekleme - �Y�LE�T�R�LD�
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
                        message = "Sat�r ekleme verisi gerekli." 
                    });
                }

                if (string.IsNullOrEmpty(addRowDto.FileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� zorunludur." 
                    });
                }

                if (string.IsNullOrEmpty(addRowDto.SheetName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sheet ad� zorunludur." 
                    });
                }

                if (addRowDto.RowData == null || !addRowDto.RowData.Any())
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Sat�r verisi bo� olamaz." 
                    });
                }

                _logger.LogInformation("AddExcelRow �a�r�ld�: FileName={FileName}, SheetName={SheetName}, DataCount={DataCount}", 
                    addRowDto.FileName, addRowDto.SheetName, addRowDto.RowData?.Count ?? 0);

                var result = await _excelService.AddExcelRowAsync(addRowDto.FileName, addRowDto.SheetName, addRowDto.RowData, addRowDto.AddedBy);
                
                return Ok(new { 
                    success = true, 
                    data = result,
                    message = "Yeni sat�r ba�ar�yla eklendi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Yeni sat�r eklenirken hata: {FileName}", addRowDto?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Yeni sat�r eklenirken hata olu�tu: " + ex.Message,
                    error = new
                    {
                        type = ex.GetType().Name,
                        message = ex.Message
                    }
                });
            }
        }

        /// <summary>
        /// Veri silme - �Y�LE�T�R�LD�
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
                        message = "Ge�erli bir ID gerekli." 
                    });
                }

                _logger.LogInformation("DeleteExcelData �a�r�ld�: Id={Id}, DeletedBy={DeletedBy}", id, deletedBy);

                var result = await _excelService.DeleteExcelDataAsync(id, deletedBy);
                
                return Ok(new { 
                    success = result, 
                    message = result ? "Veri ba�ar�yla silindi" : "Veri bulunamad�",
                    id = id
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri silinirken hata: {Id}", id);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veri silinirken hata olu�tu: " + ex.Message,
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
        /// Excel verilerini belirtilen kritere g�re d��a aktarma - �Y�LE�T�R�LD�
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
                        message = "Export i�in dosya ad� gerekli." 
                    });
                }

                _logger.LogInformation("Excel export iste�i al�nd�: {FileName}, Sheet: {Sheet}, RowIds: {RowIdCount}, IncludeHistory: {IncludeHistory}", 
                    exportRequest.FileName, exportRequest.SheetName, exportRequest.RowIds?.Count ?? 0, exportRequest.IncludeModificationHistory);

                var fileBytes = await _excelService.ExportToExcelAsync(exportRequest);
                
                if (fileBytes == null || fileBytes.Length == 0)
                {
                    return StatusCode(500, new { 
                        success = false, 
                        message = "Excel dosyas� olu�turulamad�" 
                    });
                }

                var fileName = $"{exportRequest.FileName}_export_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                
                _logger.LogInformation("Excel export ba�ar�l�: {FileName}, {FileSize} bytes", fileName, fileBytes.Length);

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel export edilirken hata: {FileName}", exportRequest?.FileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Excel export edilirken hata olu�tu: " + ex.Message,
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
        /// Dosyadaki sayfalar� (sheet'leri) getirme
        /// </summary>
        [HttpGet("sheets/{fileName}")]
        public async Task<IActionResult> GetSheets(string fileName)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� gerekli." 
                    });
                }

                var sheets = await _excelService.GetSheetsAsync(fileName);
                return Ok(new { 
                    success = true, 
                    data = sheets,
                    fileName = fileName,
                    count = sheets.Count,
                    message = "Dosyadaki sheet'ler ba�ar�yla getirildi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sheet'ler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Sheet'ler getirilirken hata olu�tu: " + ex.Message,
                    fileName = fileName
                });
            }
        }

        /// <summary>
        /// Veritaban�ndaki verilerin istatistiklerini alma
        /// </summary>
        [HttpGet("statistics/{fileName}")]
        public async Task<IActionResult> GetDataStatistics(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� gerekli." 
                    });
                }

                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);
                return Ok(new { 
                    success = true, 
                    data = statistics,
                    fileName = fileName,
                    sheetName = sheetName,
                    message = "�statistikler ba�ar�yla getirildi"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "�statistikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "�statistikler getirilirken hata olu�tu: " + ex.Message,
                    fileName = fileName
                });
            }
        }

        /// <summary>
        /// Excel dosyas� silme
        /// </summary>
        [HttpDelete("files/{fileName}")]
        public async Task<IActionResult> DeleteFile(string fileName, [FromQuery] string? deletedBy = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                if (string.IsNullOrEmpty(fileName))
                {
                    return BadRequest(new { 
                        success = false, 
                        message = "Dosya ad� gerekli." 
                    });
                }

                var result = await _excelService.DeleteExcelFileAsync(fileName, deletedBy);
                if (result)
                {
                    return Ok(new { 
                        success = true, 
                        message = "Dosya ba�ar�yla silindi",
                        fileName = fileName
                    });
                }
                else
                {
                    return NotFound(new { 
                        success = false, 
                        message = "Dosya bulunamad�",
                        fileName = fileName
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya silinirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Dosya silinirken hata olu�tu: " + ex.Message,
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
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);

                var file = await _context.ExcelFiles
                    .FirstOrDefaultAsync(f => f.FileName == fileName);

                if (file == null)
                {
                    return NotFound(new
                    {
                        success = false,
                        message = "Dosya bulunamad�",
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

                // Physical file kontrol�
                bool physicalFileExists = !string.IsNullOrEmpty(file.FilePath) && System.IO.File.Exists(file.FilePath);

                // E�er dosya Excel format�ndaysa sheet'leri de al
                List<string> availableSheets = new List<string>();
                if (physicalFileExists)
                {
                    try
                    {
                        availableSheets = await _excelService.GetSheetsAsync(fileName);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Dosya sheet'leri al�n�rken hata: {FileName}", fileName);
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
                    message = "Dosya durumu ba�ar�yla al�nd�",
                    recommendations = GetFileRecommendations(file.IsActive, physicalFileExists, dataRowsCount > 0, availableSheets.Any())
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya durumu kontrol edilirken hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Dosya durumu kontrol edilirken hata olu�tu: " + ex.Message,
                    fileName = fileName
                });
            }
        }

        private object GetFileRecommendations(bool isActive, bool physicalExists, bool hasData, bool hasSheets)
        {
            var recommendations = new List<string>();

            if (!isActive)
                recommendations.Add("Dosya aktif de�il. Dosyay� yeniden y�kleyin.");
            else if (!physicalExists)
                recommendations.Add("Fiziksel dosya bulunamad�. Dosyay� yeniden y�kleyin.");
            else if (!hasData)
                recommendations.Add("Dosya var ancak veri yok. Dosyay� okuyun.");
            else if (!hasSheets)
                recommendations.Add("Veritaban�nda sheet bilgisi yok. Dosyay� yeniden okuyun.");
            else
                recommendations.Add("Dosya durumu normal.");

            return recommendations;
        }

        /// <summary>
        /// Debug endpoint - Deneme ama�l�
        /// </summary>
        [HttpGet("debug/test")]
        public IActionResult DebugTest()
        {
            return Ok(new { success = true, message = "Debug test ba�ar�l�" });
        }
    }
}