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

        [HttpPost("compare-from-files")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> CompareFromFiles([FromForm] CompareExcelFilesDto request)
        {
            try
            {
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

                var uploadedFile1 = await _excelService.UploadExcelFileAsync(request.File1, request.ComparedBy);
                var uploadedFile2 = await _excelService.UploadExcelFileAsync(request.File2, request.ComparedBy);

                await _excelService.ReadExcelDataAsync(uploadedFile1.FileName, request.Sheet1Name);
                await _excelService.ReadExcelDataAsync(uploadedFile2.FileName, request.Sheet2Name);

                var result = await _comparisonService.CompareFilesAsync(
                    uploadedFile1.FileName, 
                    uploadedFile2.FileName, 
                    request.Sheet1Name ?? request.Sheet2Name);

                return Ok(new { 
                    success = true, 
                    data = result,
                    file1 = new { name = uploadedFile1.FileName, original = uploadedFile1.OriginalFileName },
                    file2 = new { name = uploadedFile2.FileName, original = uploadedFile2.OriginalFileName },
                   
                    message = "�ki Excel dosyas� ba�ar�yla kar��la�t�r�ld�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Manuel dosya kar��la�t�rmas� yap�l�rken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

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
                _logger.LogError(ex, "Dosyalar kar��la�t�r�l�rken hata: {File1} vs {File2}", compareRequest.FileName1, compareRequest.FileName2);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
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
                _logger.LogError(ex, "Versiyonlar kar��la�t�r�l�rken hata: {FileName}", compareRequest.FileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

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
                _logger.LogError(ex, "De�i�iklikler getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

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
                _logger.LogError(ex, "De�i�iklik ge�mi�i getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

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
                _logger.LogError(ex, "Sat�r ge�mi�i getirilirken hata: {RowId}", rowId);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

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
                    message = "Farkl� tarihlerdeki dosya durumlar� kar��la�t�r�ld�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Snapshot kar��la�t�rmas� yap�l�rken hata: {FileName}", compareRequest.FileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}