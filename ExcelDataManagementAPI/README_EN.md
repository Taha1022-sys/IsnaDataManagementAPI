# Excel Data Management API

This project is a C# Web API designed to manage Excel files, compare data, and integrate with Angular frontend.

## Features

### Excel Operations
- ? Excel file upload (.xlsx, .xls)
- ? Read Excel data and save to database
- ? Data updates (single and bulk)
- ? Add new rows
- ? Delete data (soft delete)
- ? Excel export

### Data Comparison
- ? Compare two Excel files
- ? Compare different versions of the same file
- ? Change statistics
- ? Data change history

### API Features
- ? RESTful API
- ? Swagger UI documentation
- ? CORS support (for Angular)
- ? SQL Server database support
- ? Pagination support

## Technologies

- **Backend:** ASP.NET Core 9.0
- **Database:** SQL Server
- **ORM:** Entity Framework Core
- **Excel Processing:** EPPlus
- **API Documentation:** Swagger/OpenAPI
- **Frontend:** Angular (separate project)

## Installation

### Requirements
- .NET 9.0 SDK
- SQL Server (LocalDB or full version)
- Visual Studio 2022 or VS Code

### Steps

1. **Clone the project:**
```bash
git clone [repo-url]
cd ExcelDataManagementAPI
```

2. **Install packages:**
```bash
dotnet restore
```

3. **Configure database connection:**
Edit the SQL Server connection string in `appsettings.json`:
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=ExcelDataManagementDB;Trusted_Connection=true;TrustServerCertificate=true;"
  }
}
```

4. **Run the application:**
```bash
dotnet run
```

5. **Access Swagger UI:**
```
https://localhost:5001/swagger
```

## API Endpoints

### Excel Controller (/api/excel)
- `POST /upload` - Upload Excel file
- `POST /read/{fileName}` - Read Excel file
- `GET /data/{fileName}` - Get Excel data
- `PUT /data` - Update data
- `PUT /data/bulk` - Bulk data update
- `POST /data` - Add new row
- `DELETE /data/{id}` - Delete data
- `POST /export` - Excel export
- `GET /files` - File list
- `GET /sheets/{fileName}` - Sheet list
- `GET /statistics/{fileName}` - Data statistics

### Comparison Controller (/api/comparison)
- `POST /files` - Compare two files
- `POST /versions` - Compare versions
- `GET /changes/{fileName}` - Date-based changes
- `GET /history/{fileName}` - Change history
- `GET /row-history/{rowId}` - Row history

## Usage Examples

### Excel File Upload
```javascript
const formData = new FormData();
formData.append('file', fileInput.files[0]);
formData.append('uploadedBy', 'user@email.com');

fetch('/api/excel/upload', {
    method: 'POST',
    body: formData
});
```

### Data Update
```javascript
const updateData = {
    id: 1,
    data: {
        "FirstName": "New Name",
        "LastName": "New Surname",
        "Email": "new@email.com"
    },
    modifiedBy: "user@email.com"
};

fetch('/api/excel/data', {
    method: 'PUT',
    headers: {
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(updateData)
});
```

## Database Structure

### ExcelFiles Table
- Id, FileName, OriginalFileName, FilePath
- FileSize, UploadDate, UploadedBy, IsActive

### ExcelDataRows Table
- Id, FileName, SheetName, RowIndex
- RowData (JSON), CreatedDate, ModifiedDate
- IsDeleted, Version, ModifiedBy

## Frontend Integration

For Angular frontend:
```typescript
// Service example
@Injectable()
export class ExcelService {
    private apiUrl = 'https://localhost:5001/api';
    
    uploadFile(file: File): Observable<any> {
        const formData = new FormData();
        formData.append('file', file);
        return this.http.post(`${this.apiUrl}/excel/upload`, formData);
    }
    
    getData(fileName: string, page = 1): Observable<any> {
        return this.http.get(`${this.apiUrl}/excel/data/${fileName}?page=${page}`);
    }
    
    updateData(updateData: any): Observable<any> {
        return this.http.put(`${this.apiUrl}/excel/data`, updateData);
    }
    
    deleteData(id: number): Observable<any> {
        return this.http.delete(`${this.apiUrl}/excel/data/${id}`);
    }
    
    exportData(exportRequest: any): Observable<Blob> {
        return this.http.post(`${this.apiUrl}/excel/export`, exportRequest, {
            responseType: 'blob'
        });
    }
    
    compareFiles(file1: string, file2: string): Observable<any> {
        return this.http.post(`${this.apiUrl}/comparison/files`, {
            fileName1: file1,
            fileName2: file2
        });
    }
    
    getChangeHistory(fileName: string, page = 1): Observable<any> {
        return this.http.get(`${this.apiUrl}/comparison/history/${fileName}?page=${page}`);
    }
}
```

## API Response Formats

### Success Response
```json
{
    "success": true,
    "data": { ... },
    "message": "Operation completed successfully"
}
```

### Error Response
```json
{
    "success": false,
    "message": "Error description",
    "errors": [ ... ]
}
```

### Paginated Response
```json
{
    "success": true,
    "data": [ ... ],
    "pagination": {
        "page": 1,
        "pageSize": 50,
        "hasMore": true
    }
}
```

## Development Guidelines

### Project Structure
```
ExcelDataManagementAPI/
??? Controllers/          # API Controllers
??? Data/                # Database Context
??? Models/              # Data Models
?   ??? DTOs/           # Data Transfer Objects
?   ??? ExcelDataModels.cs
??? Services/           # Business Logic
??? Properties/         # Launch Settings
??? uploads/           # Uploaded Files (created at runtime)
```

### Adding New Features
1. Create DTOs in `Models/DTOs/`
2. Add business logic to appropriate service
3. Create controller endpoints
4. Update Swagger documentation
5. Add unit tests

### Error Handling
- All controllers use try-catch blocks
- Consistent error response format
- Detailed logging for debugging
- User-friendly error messages

## Configuration

### appsettings.json
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=ExcelDataManagementDB;Trusted_Connection=true;TrustServerCertificate=true;"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*",
  "Cors": {
    "AllowedOrigins": ["http://localhost:4200"]
  }
}
```

### Environment Variables
- `ASPNETCORE_ENVIRONMENT`: Development/Production
- `ConnectionStrings__DefaultConnection`: Database connection string

## Security Considerations

- File upload size limits
- Allowed file extensions validation
- SQL injection prevention through parameterized queries
- CORS configuration for trusted origins
- Input validation and sanitization

## Performance Optimization

- Pagination for large datasets
- Async/await patterns throughout
- Database indexing on frequently queried columns
- Memory-efficient Excel processing
- Response caching where appropriate

## Testing

### Unit Tests
```bash
dotnet test
```

### API Testing
Use the included `ExcelDataManagementAPI.http` file with VS Code REST Client extension.

## Deployment

### Development
```bash
dotnet run --environment Development
```

### Production
```bash
dotnet publish -c Release
```

## License

This project uses EPPlus Non-Commercial license.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## Support

For issues and questions, please create an issue in the repository.