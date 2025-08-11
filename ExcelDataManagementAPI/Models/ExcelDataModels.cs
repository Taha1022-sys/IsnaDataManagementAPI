using System.ComponentModel.DataAnnotations;

namespace ExcelDataManagementAPI.Models
{
    public class ExcelDataRow
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        [MaxLength(255)]
        public string FileName { get; set; } = string.Empty;
        
        [Required]
        [MaxLength(255)]
        public string SheetName { get; set; } = string.Empty;
        
        public int RowIndex { get; set; }
        
        [Required]
        public string RowData { get; set; } = string.Empty;
        
        public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
        
        public DateTime? ModifiedDate { get; set; }
        
        public bool IsDeleted { get; set; } = false;
        
        public int Version { get; set; } = 1;
        
        [MaxLength(255)]
        public string? ModifiedBy { get; set; }
    }
    
    public class ExcelFile
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        [MaxLength(255)]
        public string FileName { get; set; } = string.Empty;
        
        [MaxLength(255)]
        public string OriginalFileName { get; set; } = string.Empty;
        
        [MaxLength(500)]
        public string FilePath { get; set; } = string.Empty;
        
        public long FileSize { get; set; }
        
        public DateTime UploadDate { get; set; } = DateTime.UtcNow;
        
        [MaxLength(255)]
        public string? UploadedBy { get; set; }
        
        public bool IsActive { get; set; } = true;
    }

    public class GerceklesenRaporlar
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        [MaxLength(255)]
        public string FileName { get; set; } = string.Empty;
        
        [Required]
        [MaxLength(255)]
        public string SheetName { get; set; } = string.Empty;
        
        public int RowIndex { get; set; }
        
        public int OriginalRowId { get; set; }
        
        [Required]
        [MaxLength(50)]
        public string OperationType { get; set; } = string.Empty;
        
        public string? OldValue { get; set; }
        
        public string? NewValue { get; set; }
        
        [MaxLength(255)]
        public string? ModifiedBy { get; set; }
        
        public DateTime ChangeDate { get; set; } = DateTime.UtcNow;
        
        [MaxLength(50)]
        public string? UserIP { get; set; }
        
        [MaxLength(500)]
        public string? UserAgent { get; set; }
        
        [MaxLength(1000)]
        public string? ChangeReason { get; set; }
        
        public string? ChangedColumns { get; set; }
        
        public bool IsSuccess { get; set; } = true;
        
        [MaxLength(1000)]
        public string? ErrorMessage { get; set; }
    }
}