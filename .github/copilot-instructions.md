# Excel Data Management API
Excel Data Management API is a .NET 9.0 ASP.NET Core Web API application for managing and comparing Excel files. It provides endpoints for uploading, reading, updating, and comparing Excel files with data stored in either SQL Server (Windows) or SQLite (cross-platform).

Always reference these instructions first and fallback to search or bash commands only when you encounter unexpected information that does not match the info here.

## Working Effectively
- Bootstrap, build, and test the repository:
  - Install .NET 9.0 SDK: `curl -sSL https://dot.net/v1/dotnet-install.sh | bash /dev/stdin --channel 9.0 --install-dir /home/runner/.dotnet`
  - Update PATH: `export PATH="/home/runner/.dotnet:$PATH"`
  - Restore packages: `dotnet restore` -- takes 35 seconds. NEVER CANCEL. Set timeout to 60+ minutes.
  - Build: `dotnet build --configuration Release` -- takes 8.5 seconds with 2 warnings. NEVER CANCEL. Set timeout to 30+ minutes.
- Run the application:
  - ALWAYS run the bootstrapping steps first.
  - Development: `dotnet run --urls "http://localhost:5002"` -- starts in 2-3 seconds
  - Production: `dotnet run --configuration Release --urls "http://localhost:5002"`
- Database configuration:
  - Windows: Uses SQL Server LocalDB automatically
  - Linux/macOS: Uses SQLite (ExcelDataManagement.db) automatically
  - Database initializes on first run - check console for "✅ Veritabanı hazır!" success message

## Validation
- Always manually validate any new code by testing API endpoints after making changes.
- ALWAYS run through at least one complete end-to-end scenario after making changes.
- Test Swagger UI at: http://localhost:5002/swagger
- Test API endpoints at: http://localhost:5002/api/excel/files (returns JSON response)
- The application has 31 API endpoints available across Excel and Comparison controllers.
- Always build and test your changes before considering them complete.

## Database Support
- Cross-platform database support is built-in:
  - Windows: SQL Server LocalDB (automatic detection)
  - Linux/macOS: SQLite with file-based storage (automatic fallback)
- Connection strings in appsettings.json:
  - DefaultConnection: SQL Server LocalDB
  - SqliteConnection: SQLite file database
- Database auto-creates tables and indexes on first run

## Common tasks
The following are outputs from frequently run commands. Reference them instead of viewing, searching, or running bash commands to save time.

### Repository structure
```
.
├── .git/
├── .gitattributes
├── .gitignore
├── ExcelDataManagementAPI/
│   ├── Controllers/
│   │   ├── ComparisonController.cs
│   │   └── ExcelController.cs
│   ├── Data/
│   │   └── ExcelDataContext.cs
│   ├── Models/
│   │   ├── DTOs/
│   │   └── ExcelDataModels.cs
│   ├── Services/
│   ├── Properties/
│   │   └── launchSettings.json
│   ├── uploads/
│   ├── ExcelDataManagementAPI.csproj
│   ├── ExcelDataManagementAPI.http
│   ├── Program.cs
│   ├── appsettings.json
│   └── appsettings.Development.json
└── ExcelDataManagementSolution.sln
```

### Key project files
- **ExcelDataManagementAPI.csproj**: .NET 9.0 project with Entity Framework, EPPlus, Swagger
- **Program.cs**: Main application configuration with cross-platform database support
- **appsettings.json**: Configuration with both SQL Server and SQLite connection strings
- **Data/ExcelDataContext.cs**: Entity Framework context with SQL Server/SQLite compatibility

### Build timings (NEVER CANCEL these operations)
- `dotnet restore`: 35 seconds (timeout: 60+ minutes)
- `dotnet build`: 8.5 seconds (timeout: 30+ minutes) 
- `dotnet run`: 2-3 seconds to start (timeout: 30+ minutes)
- Application startup includes database initialization

### Package dependencies
- Microsoft.NET.Sdk.Web (9.0)
- EPPlus 7.5.0 (Excel processing)
- Microsoft.EntityFrameworkCore.SqlServer 9.0.7
- Microsoft.EntityFrameworkCore.Sqlite 9.0.8 (cross-platform support)
- Microsoft.EntityFrameworkCore.Tools 9.0.7
- Swashbuckle.AspNetCore 9.0.3 (Swagger)

### Application URLs
- Swagger UI: http://localhost:5002/swagger
- API Base: http://localhost:5002/api
- HTTPS Swagger: https://localhost:7002/swagger  
- HTTPS API: https://localhost:7002/api

### Common API endpoints to test
- GET /api/excel/files - List uploaded Excel files
- GET /api/excel/test - Simple test endpoint
- GET /api/comparison/test - Comparison service test
- POST /api/excel/upload - Upload Excel files
- GET /swagger/v1/swagger.json - API specification

### Known warnings (expected)
- CS8602: Dereference of a possibly null reference (ComparisonController.cs:716)
- CS8604: Possible null reference argument (ExcelController.cs:847)
These warnings do not prevent compilation or execution.

### Cross-platform compatibility notes
- SQL Server LocalDB only works on Windows
- SQLite automatically used as fallback on Linux/macOS
- No additional configuration needed - application detects platform automatically
- EPPlus license set to NonCommercial in Program.cs

### CORS configuration
- Frontend URLs pre-configured: localhost:5174, localhost:3000, localhost:5173
- Development policy allows any origin for testing
- Production policy restricts to configured origins

### File upload support
- Maximum request body size: 100MB
- Upload directory: uploads/ (created automatically)
- Supports Excel files (.xlsx, .xls) via EPPlus library

## Critical reminders
- **NEVER CANCEL**: Build operations may take several minutes - always wait for completion
- **Database**: Application handles cross-platform database differences automatically
- **Validation**: Always test API endpoints after changes to ensure functionality
- **Dependencies**: Ensure .NET 9.0 SDK is installed before building
- **Timeouts**: Use 60+ minute timeouts for restore operations, 30+ minutes for builds