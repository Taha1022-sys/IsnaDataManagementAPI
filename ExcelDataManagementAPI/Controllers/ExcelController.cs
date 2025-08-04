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
        /// Dosyadan veri getirme (sayfalama ile)
        /// </summary>
        [HttpGet("data/{fileName}")]
        public async Task<IActionResult> GetExcelData(string fileName, [FromQuery] string? sheetName = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("Veri getirilmeye �al���l�yor: FileName={FileName}, SheetName={SheetName}, Page={Page}", fileName, sheetName, page);

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
                            .Select(f => f.FileName)
                            .ToListAsync()
                    });
                }

                // Verileri getir
                var data = await _excelService.GetExcelDataAsync(fileName, sheetName, page, pageSize);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);

                // E�er veri yoksa ancak dosya varsa, dosyan�n okunup okunmad���n� kontrol et
                if (data.Count == 0)
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
                
                return Ok(new { 
                    success = true, 
                    data = data, 
                    fileName = fileName,
                    sheetName = sheetName,
                    page = page, 
                    pageSize = pageSize,
                    statistics = statistics,
                    hasData = data.Count > 0,
                    message = data.Count > 0 ? $"Sayfa {page} - {data.Count} kay�t g�steriliyor" : "Bu sayfada g�sterilecek veri bulunamad�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}", fileName, sheetName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veriler getirilirken bir hata olu�tu: " + ex.Message,
                    fileName = fileName,
                    sheetName = sheetName
                });
            }
        }

        /// <summary>
        /// Dosyadan t�m verileri getirme (sayfalama olmadan)
        /// </summary>
        [HttpGet("data/{fileName}/all")]
        public async Task<IActionResult> GetAllExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);
                
                _logger.LogInformation("T�m veriler getirilmeye �al���l�yor: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

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
                            .Select(f => f.FileName)
                            .ToListAsync()
                    });
                }

                var data = await _excelService.GetAllExcelDataAsync(fileName, sheetName);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);

                // E�er veri yoksa detayl� bilgi ver
                if (data.Count == 0)
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
                    data = data, 
                    fileName = fileName,
                    sheetName = sheetName,
                    totalRows = data.Count,
                    statistics = statistics,
                    hasData = data.Count > 0,
                    message = data.Count > 0 ? $"Toplam {data.Count} kay�t getirildi" : "Bu dosya/sayfa i�in veri bulunamad�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}, Sheet: {SheetName}", fileName, sheetName);
                return StatusCode(500, new { 
                    success = false, 
                    message = "Veriler getirilirken bir hata olu�tu: " + ex.Message,
                    fileName = fileName,
                    sheetName = sheetName
                });
            }
        }

        /// <summary>
        /// Excel verisini g�ncelleme
        /// </summary>
        [HttpPut("data")]
        public async Task<IActionResult> UpdateExcelData([FromBody] ExcelDataUpdateDto updateDto)
        {
            try
            {
                var result = await _excelService.UpdateExcelDataAsync(updateDto);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri g�ncellenirken hata: {Id}", updateDto.Id);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Toplu Excel verisini g�ncelleme
        /// </summary>
        [HttpPut("data/bulk")]
        public async Task<IActionResult> BulkUpdateExcelData([FromBody] BulkUpdateDto bulkUpdateDto)
        {
            try
            {
                var results = await _excelService.BulkUpdateExcelDataAsync(bulkUpdateDto);
                return Ok(new { success = true, data = results, count = results.Count });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Toplu veri g�ncellenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Yeni bir sat�r ekleme
        /// </summary>
        [HttpPost("data")]
        public async Task<IActionResult> AddExcelRow([FromBody] AddRowRequestDto addRowDto)
        {
            try
            {
                var result = await _excelService.AddExcelRowAsync(addRowDto.FileName, addRowDto.SheetName, addRowDto.RowData, addRowDto.AddedBy);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Yeni sat�r eklenirken hata: {FileName}", addRowDto.FileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Veri silme
        /// </summary>
        [HttpDelete("data/{id}")]
        public async Task<IActionResult> DeleteExcelData(int id, [FromQuery] string? deletedBy = null)
        {
            try
            {
                var result = await _excelService.DeleteExcelDataAsync(id, deletedBy);
                return Ok(new { success = result, message = result ? "Veri silindi" : "Veri bulunamad�" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri silinirken hata: {Id}", id);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Excel verilerini belirtilen kritere g�re d��a aktarma
        /// </summary>
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
                
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);
                return Ok(new { success = true, data = statistics });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "�statistikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
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
                    message = ex.Message,
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

        /// <summary>
        /// Debug endpoint - Belirli dosya i�in veri sorgusu test etme
        /// </summary>
        [HttpGet("debug/test-data-query/{fileName}")]
        public async Task<IActionResult> TestDataQuery(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);

                _logger.LogInformation("Test data query ba�lat�ld�: FileName={FileName}, SheetName={SheetName}", fileName, sheetName);

                // 1. Dosya kontrol�
                var file = await _context.ExcelFiles
                    .FirstOrDefaultAsync(f => f.FileName == fileName);

                var fileInfo = new
                {
                    exists = file != null,
                    isActive = file?.IsActive ?? false,
                    fileName = file?.FileName,
                    originalFileName = file?.OriginalFileName,
                    uploadDate = file?.UploadDate,
                    physicalFileExists = file != null && !string.IsNullOrEmpty(file.FilePath) && System.IO.File.Exists(file.FilePath)
                };

                // 2. Veri kontrol�
                var baseQuery = _context.ExcelDataRows
                    .Where(r => r.FileName == fileName);

                var allDataCount = await baseQuery.CountAsync();
                var activeDataCount = await baseQuery.Where(r => !r.IsDeleted).CountAsync();
                var deletedDataCount = await baseQuery.Where(r => r.IsDeleted).CountAsync();

                var dataInfo = new
                {
                    totalDataRows = allDataCount,
                    activeDataRows = activeDataCount,
                    deletedDataRows = deletedDataCount
                };

                // 3. Sheet kontrol�
                var allSheets = await baseQuery
                    .Where(r => !r.IsDeleted)
                    .Select(r => r.SheetName)
                    .Distinct()
                    .ToListAsync();

                var sheetInfo = new
                {
                    requestedSheet = sheetName,
                    availableSheets = allSheets,
                    requestedSheetExists = !string.IsNullOrEmpty(sheetName) && allSheets.Contains(sheetName),
                    sheetDataCount = !string.IsNullOrEmpty(sheetName) 
                        ? await baseQuery.Where(r => r.SheetName == sheetName && !r.IsDeleted).CountAsync()
                        : 0
                };

                // 4. Sample data
                var sampleData = await baseQuery
                    .Where(r => !r.IsDeleted)
                    .Take(3)
                    .Select(r => new
                    {
                        r.Id,
                        r.FileName,
                        r.SheetName,
                        r.RowIndex,
                        r.CreatedDate,
                        r.IsDeleted,
                        DataPreview = r.RowData.Substring(0, Math.Min(r.RowData.Length, 100)) + (r.RowData.Length > 100 ? "..." : "")
                    })
                    .ToListAsync();

                return Ok(new
                {
                    success = true,
                    testResults = new
                    {
                        queryParameters = new { fileName, sheetName },
                        fileInfo,
                        dataInfo,
                        sheetInfo,
                        sampleData,
                        diagnosis = GetDiagnosis(fileInfo, dataInfo, sheetInfo)
                    },
                    message = "Test sorgusu tamamland�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Test data query'de hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message,
                    fileName = fileName,
                    sheetName = sheetName
                });
            }
        }

        /// <summary>
        /// Debug endpoint - Veri ak���n� test etme
        /// </summary>
        [HttpGet("debug/data-flow-test/{fileName}")]
        public async Task<IActionResult> TestDataFlow(string fileName, [FromQuery] string? sheetName = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 10)
        {
            try
            {
                // URL decode i�lemi
                fileName = Uri.UnescapeDataString(fileName);

                var steps = new List<object>();

                // Ad�m 1: Dosya kontrol�
                var fileCheck = await _context.ExcelFiles
                    .FirstOrDefaultAsync(f => f.FileName == fileName && f.IsActive);

                steps.Add(new
                {
                    step = 1,
                    description = "Dosya kontrol�",
                    success = fileCheck != null,
                    result = fileCheck != null ? new { fileCheck.FileName, fileCheck.OriginalFileName, fileCheck.IsActive } : null
                });

                if (fileCheck == null)
                {
                    return Ok(new { success = false, steps, message = "Dosya bulunamad�" });
                }

                // Ad�m 2: Veri say�s� kontrol�
                var dataCount = await _context.ExcelDataRows
                    .Where(r => r.FileName == fileName && !r.IsDeleted)
                    .CountAsync();

                steps.Add(new
                {
                    step = 2,
                    description = "Aktif veri say�s� kontrol�",
                    success = dataCount > 0,
                    result = new { dataCount }
                });

                // Ad�m 3: Sheet kontrol� (e�er belirtildiyse)
                if (!string.IsNullOrEmpty(sheetName))
                {
                    var sheetDataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == fileName && r.SheetName == sheetName && !r.IsDeleted)
                        .CountAsync();

                    steps.Add(new
                    {
                        step = 3,
                        description = $"Sheet '{sheetName}' veri kontrol�",
                        success = sheetDataCount > 0,
                        result = new { sheetName, sheetDataCount }
                    });
                }

                // Ad�m 4: Service metodunu test et
                try
                {
                    var serviceResult = await _excelService.GetExcelDataAsync(fileName, sheetName, page, pageSize);
                    steps.Add(new
                    {
                        step = 4,
                        description = "ExcelService.GetExcelDataAsync �a�r�s�",
                        success = true,
                        result = new { returnedCount = serviceResult.Count, hasData = serviceResult.Any() }
                    });

                    // Ad�m 5: �lk kay�tlar�n �rne�i
                    if (serviceResult.Any())
                    {
                        var sample = serviceResult.Take(2).Select(r => new
                        {
                            r.Id,
                            r.FileName,
                            r.SheetName,
                            r.RowIndex,
                            ColumnCount = r.Data?.Count ?? 0,
                            FirstColumns = r.Data?.Take(3).ToDictionary(kvp => kvp.Key, kvp => kvp.Value) ?? new Dictionary<string, string>()
                        });

                        steps.Add(new
                        {
                            step = 5,
                            description = "Veri �rnekleri",
                            success = true,
                            result = sample
                        });
                    }

                    return Ok(new
                    {
                        success = true,
                        steps,
                        finalResult = new
                        {
                            fileName,
                            sheetName,
                            page,
                            pageSize,
                            resultCount = serviceResult.Count,
                            hasData = serviceResult.Any()
                        },
                        message = "Veri ak�� testi tamamland�"
                    });
                }
                catch (Exception serviceEx)
                {
                    steps.Add(new
                    {
                        step = 4,
                        description = "ExcelService.GetExcelDataAsync �a�r�s�",
                        success = false,
                        error = serviceEx.Message
                    });

                    return Ok(new { success = false, steps, message = "Service katman�nda hata" });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Data flow test'te hata: {FileName}", fileName);
                return StatusCode(500, new { 
                    success = false, 
                    message = ex.Message,
                    fileName = fileName
                });
            }
        }

        private object GetDiagnosis(dynamic fileInfo, dynamic dataInfo, dynamic sheetInfo)
        {
            var issues = new List<string>();
            var recommendations = new List<string>();

            if (!fileInfo.exists)
            {
                issues.Add("Dosya veritaban�nda bulunamad�");
                recommendations.Add("Dosyay� yeniden y�kleyin");
            }
            else if (!fileInfo.isActive)
            {
                issues.Add("Dosya aktif de�il");
                recommendations.Add("Dosya silinmi� olabilir, yeniden y�kleyin");
            }
            else if (!fileInfo.physicalFileExists)
            {
                issues.Add("Fiziksel dosya bulunamad�");
                recommendations.Add("Dosya sistemi dosyas� silinmi�, yeniden y�kleyin");
            }
            else if (dataInfo.activeDataRows == 0)
            {
                if (dataInfo.totalDataRows == 0)
                {
                    issues.Add("Dosya hi� okunmam��");
                    recommendations.Add("POST /api/excel/read/{fileName} endpoint'ini kullanarak dosyay� okuyun");
                }
                else
                {
                    issues.Add("T�m veriler silinmi� durumda");
                    recommendations.Add("Dosyay� yeniden okuyun");
                }
            }
            else if (!string.IsNullOrEmpty((string)sheetInfo.requestedSheet) && !sheetInfo.requestedSheetExists)
            {
                issues.Add($"�stenen sheet '{sheetInfo.requestedSheet}' bulunamad�");
                recommendations.Add($"Mevcut sheet'lerden birini kullan�n: {string.Join(", ", sheetInfo.availableSheets)}");
            }
            else if (!string.IsNullOrEmpty((string)sheetInfo.requestedSheet) && sheetInfo.sheetDataCount == 0)
            {
                issues.Add($"�stenen sheet '{sheetInfo.requestedSheet}' var ama veri yok");
                recommendations.Add("Dosyay� yeniden okuyun veya ba�ka bir sheet se�in");
            }

            if (issues.Count == 0)
            {
                issues.Add("Belirgin bir sorun tespit edilmedi");
                recommendations.Add("Frontend'deki parametreleri kontrol edin");
            }

            return new { issues, recommendations };
        }
     }
}