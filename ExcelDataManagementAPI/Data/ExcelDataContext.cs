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
        public DbSet<DataImport> DataImports { get; set; }
        
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
            
            modelBuilder.Entity<DataImport>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.FileName).IsRequired().HasMaxLength(255);
                entity.Property(e => e.FilePath).HasMaxLength(500);
                entity.Property(e => e.Status).IsRequired().HasMaxLength(50);
                entity.Property(e => e.ErrorMessage).HasMaxLength(1000);
                entity.Property(e => e.CreatedBy).HasMaxLength(255);
                entity.Property(e => e.SSISPackageName).HasMaxLength(255);
                entity.Property(e => e.ProcessingLog).HasMaxLength(1000);
                
                entity.HasIndex(e => e.Status);
                entity.HasIndex(e => e.CreatedDate);
                entity.HasIndex(e => e.FileName);
            });
        }
    }
}