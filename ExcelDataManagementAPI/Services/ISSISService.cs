using ExcelDataManagementAPI.Models;

namespace ExcelDataManagementAPI.Services
{
    public interface ISSISService
    {
        /// <summary>
        /// Executes an SSIS package for data import
        /// </summary>
        /// <param name="packageName">Name of the SSIS package to execute</param>
        /// <param name="filePath">Path to the file to be processed</param>
        /// <param name="dataImportId">ID of the DataImport record to track progress</param>
        /// <returns>Execution result</returns>
        Task<SSISExecutionResult> ExecutePackageAsync(string packageName, string filePath, int dataImportId);
        
        /// <summary>
        /// Gets the status of a running SSIS package
        /// </summary>
        /// <param name="dataImportId">ID of the DataImport record</param>
        /// <returns>Current execution status</returns>
        Task<SSISExecutionStatus> GetExecutionStatusAsync(int dataImportId);
        
        /// <summary>
        /// Gets available SSIS packages
        /// </summary>
        /// <returns>List of available SSIS packages</returns>
        Task<List<SSISPackageInfo>> GetAvailablePackagesAsync();
        
        /// <summary>
        /// Cancels a running SSIS package execution
        /// </summary>
        /// <param name="dataImportId">ID of the DataImport record</param>
        /// <returns>Cancellation result</returns>
        Task<bool> CancelExecutionAsync(int dataImportId);
    }
    
    public class SSISExecutionResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public int RecordsProcessed { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public string? ExecutionLog { get; set; }
    }
    
    public class SSISExecutionStatus
    {
        public string Status { get; set; } = string.Empty;
        public int Progress { get; set; }
        public string? CurrentStep { get; set; }
        public string? ErrorMessage { get; set; }
    }
    
    public class SSISPackageInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }
    }
}