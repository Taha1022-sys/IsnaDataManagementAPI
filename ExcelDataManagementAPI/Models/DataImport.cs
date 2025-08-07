using System.ComponentModel.DataAnnotations;

namespace ExcelDataManagementAPI.Models
{
    public class DataImport
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        [MaxLength(255)]
        public string FileName { get; set; } = string.Empty;
        
        [MaxLength(500)]
        public string? FilePath { get; set; }
        
        public long FileSize { get; set; }
        
        [Required]
        [MaxLength(50)]
        public string Status { get; set; } = "Pending"; // Pending, Processing, Completed, Failed
        
        [MaxLength(1000)]
        public string? ErrorMessage { get; set; }
        
        public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
        
        public DateTime? StartedDate { get; set; }
        
        public DateTime? CompletedDate { get; set; }
        
        [MaxLength(255)]
        public string? CreatedBy { get; set; }
        
        public int RecordsProcessed { get; set; } = 0;
        
        public int RecordsTotal { get; set; } = 0;
        
        [MaxLength(255)]
        public string? SSISPackageName { get; set; }
        
        [MaxLength(1000)]
        public string? ProcessingLog { get; set; }
    }
    
    public enum ImportStatus
    {
        Pending,
        Processing,
        Completed,
        Failed
    }
}