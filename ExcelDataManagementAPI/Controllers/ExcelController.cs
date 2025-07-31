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

        [HttpGet("test")]
        public IActionResult Test()
        {
            return Ok(new { message = "Excel API çalýþýyor!", timestamp = DateTime.Now });
        }

        [HttpGet("files")]
        public async Task<IActionResult> GetFiles()
        {
            try
            {
                var files = await _excelService.GetExcelFilesAsync();
                return Ok(new { success = true, data = files });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosyalar getirilirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("upload")]
        public async Task<IActionResult> Upload()
        {
            try
            {
                var file = Request.Form.Files.FirstOrDefault();
                var uploadedBy = Request.Form["uploadedBy"].ToString();

                if (file == null || file.Length == 0)
                    return BadRequest(new { success = false, message = "Dosya seçiniz" });

                var result = await _excelService.UploadExcelFileAsync(file, uploadedBy);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya yüklenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("read/{fileName}")]
        public async Task<IActionResult> ReadExcelData(string fileName, [FromQuery] string? sheetName = null)
        {
            try
            {
                var data = await _excelService.ReadExcelDataAsync(fileName, sheetName);
                return Ok(new { success = true, data = data });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel dosyasý okunurken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("data/{fileName}")]
        public async Task<IActionResult> GetExcelData(string fileName, [FromQuery] string? sheetName = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        {
            try
            {
                var data = await _excelService.GetExcelDataAsync(fileName, sheetName, page, pageSize);
                return Ok(new { success = true, data = data, page = page, pageSize = pageSize });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel verileri getirilirken hata: {FileName}", fileName);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

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
                var sheets = await _excelService.GetSheetsAsync(fileName);
                return Ok(new { success = true, data = sheets });
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