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
    /// Web sitesindeki t�m de�i�iklikleri kaydeden ana audit tablosu
    /// </summary>
    public class GerceklesenRaporlar
    {
        [Key]
        public int Id { get; set; }
        
        /// <summary>
        /// De�i�iklik yap�lan Excel dosyas�n�n ad�
        /// </summary>
        [Required]
        [MaxLength(255)]
        public string FileName { get; set; } = string.Empty;
        
        /// <summary>
        /// De�i�iklik yap�lan sheet ad�
        /// </summary>
        [Required]
        [MaxLength(255)]
        public string SheetName { get; set; } = string.Empty;
        
        /// <summary>
        /// De�i�iklik yap�lan sat�r indexi
        /// </summary>
        public int RowIndex { get; set; }
        
        /// <summary>
        /// Orijinal ExcelDataRow tablosundaki kay�t ID'si
        /// </summary>
        public int OriginalRowId { get; set; }
        
        /// <summary>
        /// ��lem tipi: CREATE, UPDATE, DELETE
        /// </summary>
        [Required]
        [MaxLength(50)]
        public string OperationType { get; set; } = string.Empty;
        
        /// <summary>
        /// De�i�iklikten �nceki veri (JSON format�nda)
        /// </summary>
        public string? OldValue { get; set; }
        
        /// <summary>
        /// De�i�iklikten sonraki veri (JSON format�nda)
        /// </summary>
        public string? NewValue { get; set; }
        
        /// <summary>
        /// De�i�iklik yapan kullan�c�
        /// </summary>
        [MaxLength(255)]
        public string? ModifiedBy { get; set; }
        
        /// <summary>
        /// De�i�iklik tarihi
        /// </summary>
        public DateTime ChangeDate { get; set; } = DateTime.UtcNow;
        
        /// <summary>
        /// Kullan�c�n�n IP adresi
        /// </summary>
        [MaxLength(50)]
        public string? UserIP { get; set; }
        
        /// <summary>
        /// Taray�c� bilgisi
        /// </summary>
        [MaxLength(500)]
        public string? UserAgent { get; set; }
        
        /// <summary>
        /// De�i�iklik sebebi/a��klamas�
        /// </summary>
        [MaxLength(1000)]
        public string? ChangeReason { get; set; }
        
        /// <summary>
        /// Hangi h�creler de�i�mi� (JSON array format�nda)
        /// </summary>
        public string? ChangedColumns { get; set; }
        
        /// <summary>
        /// ��lem ba�ar�l� m�?
        /// </summary>
        public bool IsSuccess { get; set; } = true;
        
        /// <summary>
        /// Hata mesaj� (varsa)
        /// </summary>
        [MaxLength(1000)]
        public string? ErrorMessage { get; set; }
    }
}