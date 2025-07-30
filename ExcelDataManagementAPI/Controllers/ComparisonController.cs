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
        private readonly ILogger<ComparisonController> _logger;

        public ComparisonController(IDataComparisonService comparisonService, ILogger<ComparisonController> logger)
        {
            _comparisonService = comparisonService;
            _logger = logger;
        }

        [HttpGet("test")]
        public IActionResult Test()
        {
            return Ok(new { message = "Comparison API çalýþýyor!", timestamp = DateTime.Now });
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
                _logger.LogError(ex, "Dosyalar karþýlaþtýrýlýrken hata: {File1} vs {File2}", compareRequest.FileName1, compareRequest.FileName2);
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
                _logger.LogError(ex, "Versiyonlar karþýlaþtýrýlýrken hata: {FileName}", compareRequest.FileName);
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
                _logger.LogError(ex, "Deðiþiklikler getirilirken hata: {FileName}", fileName);
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
                _logger.LogError(ex, "Deðiþiklik geçmiþi getirilirken hata: {FileName}", fileName);
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
                _logger.LogError(ex, "Satýr geçmiþi getirilirken hata: {RowId}", rowId);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}