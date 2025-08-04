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
        /// Bilgisayardan iki Excel dosyasý seçip karþýlaþtýrma
        /// </summary>
        [HttpPost("compare-from-files")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> CompareFromFiles([FromForm] CompareExcelFilesDto request)
        {
            try
            {
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

                // Her iki dosyayý da yükle
                var uploadedFile1 = await _excelService.UploadExcelFileAsync(request.File1, request.ComparedBy);
                var uploadedFile2 = await _excelService.UploadExcelFileAsync(request.File2, request.ComparedBy);

                // Her iki dosyayý da oku
                await _excelService.ReadExcelDataAsync(uploadedFile1.FileName, request.Sheet1Name);
                await _excelService.ReadExcelDataAsync(uploadedFile2.FileName, request.Sheet2Name);

                // Karþýlaþtýr
                var result = await _comparisonService.CompareFilesAsync(
                    uploadedFile1.FileName, 
                    uploadedFile2.FileName, 
                    request.Sheet1Name ?? request.Sheet2Name);

                return Ok(new { 
                    success = true, 
                    data = result,
                    file1 = new { name = uploadedFile1.FileName, original = uploadedFile1.OriginalFileName },
                    file2 = new { name = uploadedFile2.FileName, original = uploadedFile2.OriginalFileName },
                    message = "Ýki Excel dosyasý baþarýyla karþýlaþtýrýldý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Manuel dosya karþýlaþtýrmasý yapýlýrken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Yüklü dosyalar arasýnda karþýlaþtýrma
        /// </summary>
        [HttpPost("files")]
        public async Task<IActionResult> CompareFiles([FromBody] CompareFilesRequestDto compareRequest)
        {
            try
            {
                var result = await _comparisonService.CompareFilesAsync(compareRequest.FileName1, compareRequest.FileName2, compareRequest.SheetName);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosyalar karþýlaþtýrýlýrken hata: {File1} vs {File2}", compareRequest.FileName1, compareRequest.FileName2);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Ayný dosyanýn farklý versiyonlarýný karþýlaþtýrma
        /// </summary>
        [HttpPost("versions")]
        public async Task<IActionResult> CompareVersions([FromBody] CompareVersionsRequestDto compareRequest)
        {
            try
            {
                var result = await _comparisonService.CompareVersionsAsync(compareRequest.FileName, compareRequest.Version1Date, compareRequest.Version2Date, compareRequest.SheetName);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Versiyonlar karþýlaþtýrýlýrken hata: {FileName}", compareRequest.FileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Belirli tarih aralýðýndaki deðiþiklikleri getirme
        /// </summary>
        [HttpGet("changes/{fileName}")]
        public async Task<IActionResult> GetChanges(string fileName, [FromQuery] DateTime? fromDate = null, [FromQuery] DateTime? toDate = null, [FromQuery] string? sheetName = null)
        {
            try
            {
                var changes = await _comparisonService.GetChangesAsync(fileName, fromDate, toDate, sheetName);
                return Ok(new { success = true, data = changes });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Deðiþiklikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Dosyanýn deðiþiklik geçmiþini getirme
        /// </summary>
        [HttpGet("history/{fileName}")]
        public async Task<IActionResult> GetChangeHistory(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                var history = await _comparisonService.GetChangeHistoryAsync(fileName, sheetName);
                return Ok(new { success = true, data = history });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Deðiþiklik geçmiþi getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Belirli bir satýrýn deðiþiklik geçmiþini getirme
        /// </summary>
        [HttpGet("row-history/{rowId}")]
        public async Task<IActionResult> GetRowHistory(int rowId)
        {
            try
            {
                var history = await _comparisonService.GetRowHistoryAsync(rowId);
                return Ok(new { success = true, data = history });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Satýr geçmiþi getirilirken hata: {RowId}", rowId);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Ýki farklý tarihteki dosya durumunu karþýlaþtýrma
        /// </summary>
        [HttpPost("snapshot-compare")]
        public async Task<IActionResult> CompareSnapshots([FromBody] CompareVersionsRequestDto compareRequest)
        {
            try
            {
                var result = await _comparisonService.CompareVersionsAsync(
                    compareRequest.FileName, 
                    compareRequest.Version1Date, 
                    compareRequest.Version2Date, 
                    compareRequest.SheetName);

                return Ok(new { 
                    success = true, 
                    data = result,
                    comparisonType = "snapshot",
                    message = "Farklý tarihlerdeki dosya durumlarý karþýlaþtýrýldý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Snapshot karþýlaþtýrmasý yapýlýrken hata: {FileName}", compareRequest.FileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}