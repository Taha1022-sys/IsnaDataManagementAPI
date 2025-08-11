using System.Text.Json.Serialization;

namespace ExcelDataManagementAPI.Models.DTOs
{
    public class FileUploadDto
    {
        public IFormFile File { get; set; } = null!;
        public string? UploadedBy { get; set; }
    }
    
    public class ExcelDataResponseDto
    {
        public int Id { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public int RowIndex { get; set; }
        
        [JsonPropertyName("data")]
        public Dictionary<string, string> Data { get; set; } = new();
        
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public int Version { get; set; }
        public string? ModifiedBy { get; set; }
    }
    
    public class ExcelDataUpdateDto
    {
        public int Id { get; set; }
        
        [JsonPropertyName("data")]
        public Dictionary<string, string> Data { get; set; } = new();
        
        public string? ModifiedBy { get; set; }
    }
    
    public class DataComparisonResultDto
    {
        public string ComparisonId { get; set; } = string.Empty;
        public string File1Name { get; set; } = string.Empty;
        public string File2Name { get; set; } = string.Empty;
        public DateTime ComparisonDate { get; set; }
        public List<DataDifferenceDto> Differences { get; set; } = new();
        public ComparisonSummaryDto Summary { get; set; } = new();
    }
    
    public class DataDifferenceDto
    {
        public int RowIndex { get; set; }
        public string ColumnName { get; set; } = string.Empty;
        public string? OldValue { get; set; }
        public string? NewValue { get; set; }
        public DifferenceType Type { get; set; }
    }
    
    public class ComparisonSummaryDto
    {
        public int TotalRows { get; set; }
        public int ModifiedRows { get; set; }
        public int AddedRows { get; set; }
        public int DeletedRows { get; set; }
        public int UnchangedRows { get; set; }
    }
    
    public enum DifferenceType
    {
        Modified,
        Added,
        Deleted
    }
    
    public class BulkUpdateDto
    {
        public List<ExcelDataUpdateDto> Updates { get; set; } = new();
        public string? ModifiedBy { get; set; }
    }
    
    public class ExcelExportRequestDto
    {
        public string FileName { get; set; } = string.Empty;
        public string? SheetName { get; set; }
        public List<int>? RowIds { get; set; }
        public bool IncludeModificationHistory { get; set; } = false;
    }

    public class CompareFilesRequestDto
    {
        public string FileName1 { get; set; } = string.Empty;
        public string FileName2 { get; set; } = string.Empty;
        public string? SheetName { get; set; }
    }

    public class CompareVersionsRequestDto
    {
        public string FileName { get; set; } = string.Empty;
        public DateTime Version1Date { get; set; }
        public DateTime Version2Date { get; set; }
        public string? SheetName { get; set; }
    }

    public class AddRowRequestDto
    {
        public string FileName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        
        [JsonPropertyName("rowData")]
        public Dictionary<string, string> RowData { get; set; } = new();
        
        public string? AddedBy { get; set; }
    }

    public class ManualFileSelectionDto
    {
        public IFormFile ExcelFile { get; set; } = null!;
        public string Operation { get; set; } = string.Empty; 
        public string? SheetName { get; set; }
        public string? ProcessedBy { get; set; }
    }

    public class CompareExcelFilesDto
    {
        public IFormFile File1 { get; set; } = null!;
        public IFormFile File2 { get; set; } = null!;
        public string? Sheet1Name { get; set; }
        public string? Sheet2Name { get; set; }
        public string? ComparedBy { get; set; }
    }

    public class UpdateExcelFileDto
    {
        public IFormFile ExcelFile { get; set; } = null!;
        
        [JsonPropertyName("updateData")]
        public Dictionary<string, string>? UpdateData { get; set; }
        
        public string? SheetName { get; set; }
        public string? UpdatedBy { get; set; }
    }
}