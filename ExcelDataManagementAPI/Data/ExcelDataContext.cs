using Microsoft.EntityFrameworkCore;
using ExcelDataManagementAPI.Models;

namespace ExcelDataManagementAPI.Data
{
    public class ExcelDataContext : DbContext
    {
        public ExcelDataContext(DbContextOptions<ExcelDataContext> options) : base(options)
        {
        }
        
        public DbSet<ExcelFile> ExcelFiles { get; set; }
        public DbSet<ExcelDataRow> ExcelDataRows { get; set; }
        public DbSet<GerceklesenRaporlar> GerceklesenRaporlar { get; set; }
        
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            
            modelBuilder.Entity<ExcelFile>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.FileName).IsRequired().HasMaxLength(255);
                entity.Property(e => e.OriginalFileName).HasMaxLength(255);
                entity.Property(e => e.FilePath).HasMaxLength(500);
                entity.Property(e => e.UploadedBy).HasMaxLength(255);
                entity.HasIndex(e => e.FileName);
            });

             
            modelBuilder.Entity<ExcelDataRow>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.FileName).IsRequired().HasMaxLength(255);
                entity.Property(e => e.SheetName).IsRequired().HasMaxLength(255);
                entity.Property(e => e.RowData).IsRequired().HasColumnType("nvarchar(max)");
                entity.Property(e => e.ModifiedBy).HasMaxLength(255);
                
                entity.HasIndex(e => new { e.FileName, e.SheetName });
                entity.HasIndex(e => e.RowIndex);
                entity.HasIndex(e => e.CreatedDate);
                entity.HasIndex(e => e.IsDeleted);
                entity.HasIndex(e => e.ModifiedDate);
            });

            modelBuilder.Entity<GerceklesenRaporlar>(entity =>
            {
                entity.ToTable("GerceklesenRaporlar"); 
                entity.HasKey(e => e.Id);
                
                entity.Property(e => e.FileName).IsRequired().HasMaxLength(255);
                entity.Property(e => e.SheetName).IsRequired().HasMaxLength(255);
                entity.Property(e => e.OperationType).IsRequired().HasMaxLength(50);
                entity.Property(e => e.ModifiedBy).HasMaxLength(255);
                entity.Property(e => e.UserIP).HasMaxLength(50);
                entity.Property(e => e.UserAgent).HasMaxLength(500);
                entity.Property(e => e.ChangeReason).HasMaxLength(1000);
                entity.Property(e => e.ErrorMessage).HasMaxLength(1000);
                
                entity.Property(e => e.OldValue).HasColumnType("nvarchar(max)");
                entity.Property(e => e.NewValue).HasColumnType("nvarchar(max)");
                entity.Property(e => e.ChangedColumns).HasColumnType("nvarchar(max)");
                
                entity.HasIndex(e => e.FileName);
                entity.HasIndex(e => e.SheetName);
                entity.HasIndex(e => e.OperationType);
                entity.HasIndex(e => e.ChangeDate);
                entity.HasIndex(e => e.ModifiedBy);
                entity.HasIndex(e => e.OriginalRowId);
                entity.HasIndex(e => new { e.FileName, e.SheetName, e.ChangeDate });
            });
        }
    }
}