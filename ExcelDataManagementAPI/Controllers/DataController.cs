using Microsoft.AspNetCore.Mvc;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Models;
using ExcelDataManagementAPI.Services;
using Microsoft.EntityFrameworkCore;

namespace ExcelDataManagementAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DataController : ControllerBase
    {
        private readonly ExcelDataContext _context;
        private readonly ISSISService _ssisService;
        private readonly ILogger<DataController> _logger;
        
        public DataController(ExcelDataContext context, ISSISService ssisService, ILogger<DataController> logger)
        {
            _context = context;
            _ssisService = ssisService;
            _logger = logger;
        }
        
        /// <summary>
        /// Uploads a file and creates a data import record
        /// </summary>
        [HttpPost("upload")]
        public async Task<ActionResult<DataImportResponseDto>> UploadFile(IFormFile file, [FromForm] string? createdBy = null)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("Dosya seçilmedi.");
                }
                
                // Validate file type
                var allowedExtensions = new[] { ".xlsx", ".xls", ".csv" };
                var fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest("Sadece Excel (.xlsx, .xls) ve CSV (.csv) dosyaları kabul edilir.");
                }
                
                // Create uploads directory if it doesn't exist
                var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "uploads");
                if (!Directory.Exists(uploadsDir))
                {
                    Directory.CreateDirectory(uploadsDir);
                }
                
                // Generate unique filename
                var fileName = $"{Guid.NewGuid()}_{file.FileName}";
                var filePath = Path.Combine(uploadsDir, fileName);
                
                // Save file
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }
                
                // Create DataImport record
                var dataImport = new DataImport
                {
                    FileName = file.FileName,
                    FilePath = filePath,
                    FileSize = file.Length,
                    Status = "Pending",
                    CreatedBy = createdBy,
                    CreatedDate = DateTime.UtcNow
                };
                
                _context.DataImports.Add(dataImport);
                await _context.SaveChangesAsync();
                
                _logger.LogInformation($"File uploaded successfully: {file.FileName}, DataImport ID: {dataImport.Id}");
                
                return Ok(new DataImportResponseDto
                {
                    Id = dataImport.Id,
                    FileName = dataImport.FileName,
                    FileSize = dataImport.FileSize,
                    Status = dataImport.Status,
                    CreatedDate = dataImport.CreatedDate,
                    CreatedBy = dataImport.CreatedBy,
                    Message = "Dosya başarıyla yüklendi."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading file");
                return StatusCode(500, $"Dosya yükleme hatası: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Starts SSIS package execution for a data import
        /// </summary>
        [HttpPost("{id}/execute")]
        public async Task<ActionResult<SSISExecutionResponseDto>> ExecuteSSISPackage(int id, [FromBody] ExecuteSSISRequestDto request)
        {
            try
            {
                var dataImport = await _context.DataImports.FindAsync(id);
                
                if (dataImport == null)
                {
                    return NotFound("Veri import kaydı bulunamadı.");
                }
                
                if (dataImport.Status != "Pending")
                {
                    return BadRequest($"Bu import zaten işlemde veya tamamlanmış. Mevcut durum: {dataImport.Status}");
                }
                
                if (string.IsNullOrEmpty(dataImport.FilePath) || !System.IO.File.Exists(dataImport.FilePath))
                {
                    return BadRequest("Import dosyası bulunamadı.");
                }
                
                // Execute SSIS package asynchronously
                _ = Task.Run(async () =>
                {
                    await _ssisService.ExecutePackageAsync(request.PackageName, dataImport.FilePath, id);
                });
                
                return Ok(new SSISExecutionResponseDto
                {
                    DataImportId = id,
                    PackageName = request.PackageName,
                    Status = "Started",
                    Message = "SSIS paket çalıştırması başlatıldı."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error starting SSIS execution for DataImport ID: {id}");
                return StatusCode(500, $"SSIS çalıştırma hatası: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Gets the status of a data import and its SSIS execution
        /// </summary>
        [HttpGet("{id}/status")]
        public async Task<ActionResult<DataImportStatusDto>> GetImportStatus(int id)
        {
            try
            {
                var dataImport = await _context.DataImports.FindAsync(id);
                
                if (dataImport == null)
                {
                    return NotFound("Veri import kaydı bulunamadı.");
                }
                
                var ssisStatus = await _ssisService.GetExecutionStatusAsync(id);
                
                return Ok(new DataImportStatusDto
                {
                    Id = dataImport.Id,
                    FileName = dataImport.FileName,
                    Status = dataImport.Status,
                    Progress = ssisStatus.Progress,
                    CurrentStep = ssisStatus.CurrentStep,
                    RecordsProcessed = dataImport.RecordsProcessed,
                    RecordsTotal = dataImport.RecordsTotal,
                    StartedDate = dataImport.StartedDate,
                    CompletedDate = dataImport.CompletedDate,
                    ErrorMessage = dataImport.ErrorMessage,
                    SSISPackageName = dataImport.SSISPackageName,
                    ProcessingLog = dataImport.ProcessingLog
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting import status for ID: {id}");
                return StatusCode(500, $"Durum sorgulama hatası: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Gets all data imports
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<List<DataImportSummaryDto>>> GetDataImports([FromQuery] string? status = null)
        {
            try
            {
                var query = _context.DataImports.AsQueryable();
                
                if (!string.IsNullOrEmpty(status))
                {
                    query = query.Where(d => d.Status.ToLower() == status.ToLower());
                }
                
                var imports = await query
                    .OrderByDescending(d => d.CreatedDate)
                    .Select(d => new DataImportSummaryDto
                    {
                        Id = d.Id,
                        FileName = d.FileName,
                        Status = d.Status,
                        FileSize = d.FileSize,
                        RecordsProcessed = d.RecordsProcessed,
                        CreatedDate = d.CreatedDate,
                        CompletedDate = d.CompletedDate,
                        CreatedBy = d.CreatedBy,
                        SSISPackageName = d.SSISPackageName
                    })
                    .ToListAsync();
                
                return Ok(imports);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting data imports");
                return StatusCode(500, $"Veri import listesi alma hatası: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Gets available SSIS packages
        /// </summary>
        [HttpGet("ssis/packages")]
        public async Task<ActionResult<List<SSISPackageInfo>>> GetSSISPackages()
        {
            try
            {
                var packages = await _ssisService.GetAvailablePackagesAsync();
                return Ok(packages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting SSIS packages");
                return StatusCode(500, $"SSIS paket listesi alma hatası: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Cancels SSIS execution
        /// </summary>
        [HttpPost("{id}/cancel")]
        public async Task<ActionResult> CancelSSISExecution(int id)
        {
            try
            {
                var success = await _ssisService.CancelExecutionAsync(id);
                
                if (success)
                {
                    return Ok(new { message = "SSIS çalıştırması iptal edildi." });
                }
                else
                {
                    return BadRequest("İptal işlemi başarısız oldu.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error cancelling SSIS execution for ID: {id}");
                return StatusCode(500, $"İptal işlemi hatası: {ex.Message}");
            }
        }
    }
    
    // DTOs
    public class DataImportResponseDto
    {
        public int Id { get; set; }
        public string FileName { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string Status { get; set; } = string.Empty;
        public DateTime CreatedDate { get; set; }
        public string? CreatedBy { get; set; }
        public string Message { get; set; } = string.Empty;
    }
    
    public class ExecuteSSISRequestDto
    {
        public string PackageName { get; set; } = string.Empty;
    }
    
    public class SSISExecutionResponseDto
    {
        public int DataImportId { get; set; }
        public string PackageName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }
    
    public class DataImportStatusDto
    {
        public int Id { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public int Progress { get; set; }
        public string? CurrentStep { get; set; }
        public int RecordsProcessed { get; set; }
        public int RecordsTotal { get; set; }
        public DateTime? StartedDate { get; set; }
        public DateTime? CompletedDate { get; set; }
        public string? ErrorMessage { get; set; }
        public string? SSISPackageName { get; set; }
        public string? ProcessingLog { get; set; }
    }
    
    public class DataImportSummaryDto
    {
        public int Id { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public int RecordsProcessed { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime? CompletedDate { get; set; }
        public string? CreatedBy { get; set; }
        public string? SSISPackageName { get; set; }
    }
}