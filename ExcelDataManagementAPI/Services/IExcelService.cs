using ExcelDataManagementAPI.Models;
using ExcelDataManagementAPI.Models.DTOs;

namespace ExcelDataManagementAPI.Services
{
    public interface IExcelService
    {
        Task<ExcelFile> UploadExcelFileAsync(IFormFile file, string? uploadedBy = null);
        Task<List<ExcelDataResponseDto>> ReadExcelDataAsync(string fileName, string? sheetName = null);
        
        Task<List<ExcelDataResponseDto>> GetExcelDataAsync(string fileName, string? sheetName = null, int page = 1, int pageSize = 50);
        
        Task<ExcelDataResponseDto> UpdateExcelDataAsync(ExcelDataUpdateDto updateDto);
        
        Task<List<ExcelDataResponseDto>> BulkUpdateExcelDataAsync(BulkUpdateDto bulkUpdateDto);
        
        Task<bool> DeleteExcelDataAsync(int id, string? deletedBy = null);
        
        Task<ExcelDataResponseDto> AddExcelRowAsync(string fileName, string sheetName, Dictionary<string, string> rowData, string? addedBy = null);
        
        Task<byte[]> ExportToExcelAsync(ExcelExportRequestDto exportRequest);
        
        Task<List<ExcelFile>> GetExcelFilesAsync();
        
        Task<List<string>> GetSheetsAsync(string fileName);
        
        Task<object> GetDataStatisticsAsync(string fileName, string? sheetName = null);
    }
}