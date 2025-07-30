# Excel Data Management API

Bu proje Excel dosyalar�n� y�netmek, verileri kar��la�t�rmak ve Angular frontend ile entegre olarak �al��mak i�in geli�tirilmi� bir C# Web API projesidir.

## �zellikler

### Excel ��lemleri
- ? Excel dosyas� y�kleme (.xlsx, .xls)
- ? Excel verilerini okuma ve veritaban�na kaydetme
- ? Veri g�ncelleme (tekil ve toplu)
- ? Yeni sat�r ekleme
- ? Veri silme (soft delete)
- ? Excel export

### Veri Kar��la�t�rma
- ? �ki Excel dosyas�n� kar��la�t�rma
- ? Ayn� dosyan�n farkl� versiyonlar�n� kar��la�t�rma
- ? De�i�iklik istatistikleri
- ? Veri de�i�iklik ge�mi�i

### API �zellikleri
- ? RESTful API
- ? Swagger UI dok�mantasyonu
- ? CORS deste�i (Angular i�in)
- ? SQL Server veritaban� deste�i
- ? Sayfalama deste�i

## Teknolojiler

- **Backend:** ASP.NET Core 9.0
- **Veritaban�:** SQL Server
- **ORM:** Entity Framework Core
- **Excel ��leme:** EPPlus
- **API Dok�mantasyonu:** Swagger/OpenAPI
- **Frontend:** Angular (ayr� proje)

## Kurulum

### Gereksinimler
- .NET 9.0 SDK
- SQL Server (LocalDB veya tam s�r�m)
- Visual Studio 2022 veya VS Code

### Ad�mlar

1. **Projeyi klonlay�n:**
```bash
git clone [repo-url]
cd ExcelDataManagementAPI
```

2. **Paketleri y�kleyin:**
```bash
dotnet restore
```

3. **Veritaban� ba�lant�s�n� ayarlay�n:**
`appsettings.json` dosyas�nda SQL Server ba�lant� string'ini d�zenleyin:
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=ExcelDataManagementDB;Trusted_Connection=true;TrustServerCertificate=true;"
  }
}
```

4. **Uygulamay� �al��t�r�n:**
```bash
dotnet run
```

5. **Swagger UI'a eri�in:**
```
http://localhost:5002
```

## API Endpoints

### Excel Controller (/api/excel)
- `GET /test` - API durumunu test et
- `GET /files` - Y�klenmi� dosya listesi
- `POST /upload` - Excel dosyas� y�kle
- `POST /read/{fileName}` - Excel dosyas�n� oku ve veritaban�na kaydet
- `GET /data/{fileName}` - Excel verilerini getir (sayfalama)
- `PUT /data` - Tekil veri g�ncelleme
- `PUT /data/bulk` - Toplu veri g�ncelleme
- `POST /data` - Yeni sat�r ekle
- `DELETE /data/{id}` - Veri sil (soft delete)
- `POST /export` - Excel export et
- `GET /sheets/{fileName}` - Dosyadaki sheet listesi
- `GET /statistics/{fileName}` - Dosya istatistikleri

### Comparison Controller (/api/comparison)
- `GET /test` - API durumunu test et
- `POST /files` - �ki dosyay� kar��la�t�r
- `POST /versions` - Ayn� dosyan�n farkl� versiyonlar�n� kar��la�t�r
- `GET /changes/{fileName}` - Tarih aral���ndaki de�i�iklikler
- `GET /history/{fileName}` - Dosya de�i�iklik ge�mi�i
- `GET /row-history/{rowId}` - Belirli sat�r�n ge�mi�i

## Kullan�m �rnekleri

### Excel Dosyas� Y�kleme
```javascript
const formData = new FormData();
formData.append('file', fileInput.files[0]);
formData.append('uploadedBy', 'kullanici@email.com');

fetch('/api/excel/upload', {
    method: 'POST',
    body: formData
});
```

### Excel Verilerini Okuma
```javascript
fetch('/api/excel/read/dosya_adi.xlsx?sheetName=Sheet1', {
    method: 'POST'
});
```

### Veri G�ncelleme
```javascript
const updateData = {
    id: 1,
    data: {
        "Ad": "Yeni Ad",
        "Soyad": "Yeni Soyad", 
        "Email": "yeni@email.com"
    },
    modifiedBy: "kullanici@email.com"
};

fetch('/api/excel/data', {
    method: 'PUT',
    headers: {
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(updateData)
});
```

### Yeni Sat�r Ekleme
```javascript
const newRowData = {
    fileName: "dosya_adi.xlsx",
    sheetName: "Sheet1",
    rowData: {
        "Ad": "Yeni Ki�i",
        "Soyad": "Yeni Soyad",
        "Email": "yeni@email.com"
    },
    addedBy: "kullanici@email.com"
};

fetch('/api/excel/data', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(newRowData)
});
```

### Dosya Kar��la�t�rma
```javascript
const compareRequest = {
    fileName1: "dosya1.xlsx",
    fileName2: "dosya2.xlsx", 
    sheetName: "Sheet1"
};

fetch('/api/comparison/files', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(compareRequest)
});
```

### Excel Export
```javascript
const exportRequest = {
    fileName: "dosya_adi.xlsx",
    sheetName: "Sheet1",
    rowIds: [1, 2, 3], // Opsiyonel - belirli sat�rlar� export et
    includeModificationHistory: false
};

fetch('/api/excel/export', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(exportRequest)
}).then(response => response.blob())
  .then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'export.xlsx';
      a.click();
  });
```

## Veritaban� Yap�s�

### ExcelFiles Tablosu
- Id, FileName, OriginalFileName, FilePath
- FileSize, UploadDate, UploadedBy, IsActive

### ExcelDataRows Tablosu
- Id, FileName, SheetName, RowIndex
- RowData (JSON), CreatedDate, ModifiedDate
- IsDeleted, Version, ModifiedBy

## Frontend Entegrasyonu

Angular frontend i�in:
```typescript
// Service �rne�i
@Injectable()
export class ExcelService {
    private apiUrl = 'http://localhost:5002/api';
    
    uploadFile(file: File, uploadedBy?: string): Observable<any> {
        const formData = new FormData();
        formData.append('file', file);
        if (uploadedBy) formData.append('uploadedBy', uploadedBy);
        return this.http.post(`${this.apiUrl}/excel/upload`, formData);
    }
    
    readExcelData(fileName: string, sheetName?: string): Observable<any> {
        const url = `${this.apiUrl}/excel/read/${fileName}`;
        const params = sheetName ? `?sheetName=${sheetName}` : '';
        return this.http.post(url + params, {});
    }
    
    getData(fileName: string, page = 1, pageSize = 50, sheetName?: string): Observable<any> {
        let params = `?page=${page}&pageSize=${pageSize}`;
        if (sheetName) params += `&sheetName=${sheetName}`;
        return this.http.get(`${this.apiUrl}/excel/data/${fileName}${params}`);
    }
    
    updateData(updateDto: any): Observable<any> {
        return this.http.put(`${this.apiUrl}/excel/data`, updateDto);
    }
    
    addRow(rowData: any): Observable<any> {
        return this.http.post(`${this.apiUrl}/excel/data`, rowData);
    }
    
    deleteRow(id: number, deletedBy?: string): Observable<any> {
        const params = deletedBy ? `?deletedBy=${deletedBy}` : '';
        return this.http.delete(`${this.apiUrl}/excel/data/${id}${params}`);
    }
    
    exportExcel(exportRequest: any): Observable<Blob> {
        return this.http.post(`${this.apiUrl}/excel/export`, exportRequest, 
            { responseType: 'blob' });
    }
    
    compareFiles(fileName1: string, fileName2: string, sheetName?: string): Observable<any> {
        return this.http.post(`${this.apiUrl}/comparison/files`, 
            { fileName1, fileName2, sheetName });
    }
}
```

## Test Endpoint'leri

API'nin �al���p �al��mad���n� test etmek i�in:
- `GET http://localhost:5002/api/excel/test`
- `GET http://localhost:5002/api/comparison/test`

## Lisans

Bu proje EPPlus Non-Commercial lisans� kullanmaktad�r.