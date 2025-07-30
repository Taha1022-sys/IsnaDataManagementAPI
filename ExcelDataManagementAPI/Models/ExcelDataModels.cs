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
}