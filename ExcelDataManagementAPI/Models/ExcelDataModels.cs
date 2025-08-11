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

    /// <summary>
    /// Web sitesindeki tüm deðiþiklikleri kaydeden ana audit tablosu
    /// </summary>
    public class GerceklesenRaporlar
    {
        [Key]
        public int Id { get; set; }
        
        /// <summary>
        /// Deðiþiklik yapýlan Excel dosyasýnýn adý
        /// </summary>
        [Required]
        [MaxLength(255)]
        public string FileName { get; set; } = string.Empty;
        
        /// <summary>
        /// Deðiþiklik yapýlan sheet adý
        /// </summary>
        [Required]
        [MaxLength(255)]
        public string SheetName { get; set; } = string.Empty;
        
        /// <summary>
        /// Deðiþiklik yapýlan satýr indexi
        /// </summary>
        public int RowIndex { get; set; }
        
        /// <summary>
        /// Orijinal ExcelDataRow tablosundaki kayýt ID'si
        /// </summary>
        public int OriginalRowId { get; set; }
        
        /// <summary>
        /// Ýþlem tipi: CREATE, UPDATE, DELETE
        /// </summary>
        [Required]
        [MaxLength(50)]
        public string OperationType { get; set; } = string.Empty;
        
        /// <summary>
        /// Deðiþiklikten önceki veri (JSON formatýnda)
        /// </summary>
        public string? OldValue { get; set; }
        
        /// <summary>
        /// Deðiþiklikten sonraki veri (JSON formatýnda)
        /// </summary>
        public string? NewValue { get; set; }
        
        /// <summary>
        /// Deðiþiklik yapan kullanýcý
        /// </summary>
        [MaxLength(255)]
        public string? ModifiedBy { get; set; }
        
        /// <summary>
        /// Deðiþiklik tarihi
        /// </summary>
        public DateTime ChangeDate { get; set; } = DateTime.UtcNow;
        
        /// <summary>
        /// Kullanýcýnýn IP adresi
        /// </summary>
        [MaxLength(50)]
        public string? UserIP { get; set; }
        
        /// <summary>
        /// Tarayýcý bilgisi
        /// </summary>
        [MaxLength(500)]
        public string? UserAgent { get; set; }
        
        /// <summary>
        /// Deðiþiklik sebebi/açýklamasý
        /// </summary>
        [MaxLength(1000)]
        public string? ChangeReason { get; set; }
        
        /// <summary>
        /// Hangi hücreler deðiþmiþ (JSON array formatýnda)
        /// </summary>
        public string? ChangedColumns { get; set; }
        
        /// <summary>
        /// Ýþlem baþarýlý mý?
        /// </summary>
        public bool IsSuccess { get; set; } = true;
        
        /// <summary>
        /// Hata mesajý (varsa)
        /// </summary>
        [MaxLength(1000)]
        public string? ErrorMessage { get; set; }
    }
}