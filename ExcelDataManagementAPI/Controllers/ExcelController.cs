using Microsoft.AspNetCore.Mvc;
using ExcelDataManagementAPI.Services;
using ExcelDataManagementAPI.Models.DTOs;

namespace ExcelDataManagementAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelService _excelService;
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(IExcelService excelService, ILogger<ExcelController> logger)
        {
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
                message = "Excel API çalýþýyor!", 
                timestamp = DateTime.Now,
                availableOperations = new[]
                {
                    "Excel dosyasý yükleme: POST /api/excel/upload",
                    "Dosya listesi: GET /api/excel/files", 
                    "Manuel dosya seçip okuma: POST /api/excel/read-from-file",
                    "Manuel dosya seçip güncelleme: POST /api/excel/update-from-file",
                    "Veri düzenleme: PUT /api/excel/data"
                }
            });
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
        /// Dosyadan veri getirme (sayfalama ile)
        /// </summary>
        [HttpGet("data/{fileName}")]
        public async Task<IActionResult> GetExcelData(string fileName, [FromQuery] string? sheetName = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        {
            try
            {
                var data = await _excelService.GetExcelDataAsync(fileName, sheetName, page, pageSize);
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);
                
                return Ok(new { 
                    success = true, 
                    data = data, 
                    fileName = fileName,
                    sheetName = sheetName,
                    page = page, 
                    pageSize = pageSize,
                    statistics = statistics,
                    message = $"Sayfa {page} - {data.Count} kayýt gösteriliyor"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Excel verisini güncelleme
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
                _logger.LogError(ex, "Veri güncellenirken hata: {Id}", updateDto.Id);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Toplu Excel verisini güncelleme
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
                _logger.LogError(ex, "Toplu veri güncellenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Yeni bir satýr ekleme
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
                _logger.LogError(ex, "Yeni satýr eklenirken hata: {FileName}", addRowDto.FileName);
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
                return Ok(new { success = result, message = result ? "Veri silindi" : "Veri bulunamadý" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Veri silinirken hata: {Id}", id);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Excel verilerini belirtilen kritere göre dýþa aktarma
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
        /// Dosyadaki sayfalarý (sheet'leri) getirme
        /// </summary>
        [HttpGet("sheets/{fileName}")]
        public async Task<IActionResult> GetSheets(string fileName)
        {
            try
            {
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
        /// Veritabanýndaki verilerin istatistiklerini alma
        /// </summary>
        [HttpGet("statistics/{fileName}")]
        public async Task<IActionResult> GetDataStatistics(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                var statistics = await _excelService.GetDataStatisticsAsync(fileName, sheetName);
                return Ok(new { success = true, data = statistics });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Ýstatistikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
     }
}