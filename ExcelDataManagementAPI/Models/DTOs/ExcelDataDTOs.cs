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
        public Dictionary<string, object> Data { get; set; } = new();
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public int Version { get; set; }
        public string? ModifiedBy { get; set; }
    }
    
    public class ExcelDataUpdateDto
    {
        public int Id { get; set; }
        public Dictionary<string, object> Data { get; set; } = new();
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
        public object? OldValue { get; set; }
        public object? NewValue { get; set; }
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
        public Dictionary<string, object> RowData { get; set; } = new();
        public string? AddedBy { get; set; }
    }
}