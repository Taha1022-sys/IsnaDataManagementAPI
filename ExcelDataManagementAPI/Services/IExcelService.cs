using ExcelDataManagementAPI.Models;
using ExcelDataManagementAPI.Models.DTOs;

namespace ExcelDataManagementAPI.Services
{
    public interface IExcelService
    {
        Task<ExcelFile> UploadExcelFileAsync(IFormFile file, string? uploadedBy = null);
        Task<List<ExcelFile>> GetExcelFilesAsync();
        Task<List<ExcelDataResponseDto>> ReadExcelDataAsync(string fileName, string? sheetName = null);
        Task<Dictionary<string, List<ExcelDataResponseDto>>> ReadAllSheetsFromExcelAsync(string fileName);
        Task<List<ExcelDataResponseDto>> GetExcelDataAsync(string fileName, string? sheetName = null, int page = 1, int pageSize = 50);
        Task<List<ExcelDataResponseDto>> GetAllExcelDataAsync(string fileName, string? sheetName = null);
        Task<ExcelDataResponseDto> UpdateExcelDataAsync(ExcelDataUpdateDto updateDto, HttpContext? httpContext = null);
        Task<List<ExcelDataResponseDto>> BulkUpdateExcelDataAsync(BulkUpdateDto bulkUpdateDto, HttpContext? httpContext = null);
        Task<bool> DeleteExcelDataAsync(int id, string? deletedBy = null, HttpContext? httpContext = null);
        Task<bool> DeleteExcelFileAsync(string fileName, string? deletedBy = null, HttpContext? httpContext = null);
        Task<ExcelDataResponseDto> AddExcelRowAsync(string fileName, string sheetName, Dictionary<string, string> rowData, string? addedBy = null, HttpContext? httpContext = null);
        Task<byte[]> ExportToExcelAsync(ExcelExportRequestDto exportRequest);
        Task<List<string>> GetSheetsAsync(string fileName);
        Task<object> GetDataStatisticsAsync(string fileName, string? sheetName = null);
    }
}