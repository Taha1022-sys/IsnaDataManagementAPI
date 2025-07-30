using ExcelDataManagementAPI.Models.DTOs;

namespace ExcelDataManagementAPI.Services
{
    public interface IDataComparisonService
    {
        Task<DataComparisonResultDto> CompareFilesAsync(string fileName1, string fileName2, string? sheetName = null);
        
        Task<DataComparisonResultDto> CompareVersionsAsync(string fileName, DateTime version1Date, DateTime version2Date, string? sheetName = null);
        
        Task<List<DataDifferenceDto>> GetChangesAsync(string fileName, DateTime? fromDate = null, DateTime? toDate = null, string? sheetName = null);
        
        Task<List<object>> GetChangeHistoryAsync(string fileName, string? sheetName = null);
        
        Task<List<object>> GetRowHistoryAsync(int rowId);
    }
}