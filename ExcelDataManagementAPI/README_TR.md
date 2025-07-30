# Excel Data Management API

Bu proje Excel dosyalarýný yönetmek, verileri karþýlaþtýrmak ve Angular frontend ile entegre olarak çalýþmak için geliþtirilmiþ bir C# Web API projesidir.

## Özellikler

### Excel Ýþlemleri
- ? Excel dosyasý yükleme (.xlsx, .xls)
- ? Excel verilerini okuma ve veritabanýna kaydetme
- ? Veri güncelleme (tekil ve toplu)
- ? Yeni satýr ekleme
- ? Veri silme (soft delete)
- ? Excel export

### Veri Karþýlaþtýrma
- ? Ýki Excel dosyasýný karþýlaþtýrma
- ? Ayný dosyanýn farklý versiyonlarýný karþýlaþtýrma
- ? Deðiþiklik istatistikleri
- ? Veri deðiþiklik geçmiþi

### API Özellikleri
- ? RESTful API
- ? Swagger UI dokümantasyonu
- ? CORS desteði (Angular için)
- ? SQL Server veritabaný desteði
- ? Sayfalama desteði

## Teknolojiler

- **Backend:** ASP.NET Core 9.0
- **Veritabaný:** SQL Server
- **ORM:** Entity Framework Core
- **Excel Ýþleme:** EPPlus
- **API Dokümantasyonu:** Swagger/OpenAPI
- **Frontend:** Angular (ayrý proje)

## Kurulum

### Gereksinimler
- .NET 9.0 SDK
- SQL Server (LocalDB veya tam sürüm)
- Visual Studio 2022 veya VS Code

### Adýmlar

1. **Projeyi klonlayýn:**
```bash
git clone [repo-url]
cd ExcelDataManagementAPI
```

2. **Paketleri yükleyin:**
```bash
dotnet restore
```

3. **Veritabaný baðlantýsýný ayarlayýn:**
`appsettings.json` dosyasýnda SQL Server baðlantý string'ini düzenleyin:
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=ExcelDataManagementDB;Trusted_Connection=true;TrustServerCertificate=true;"
  }
}
```

4. **Uygulamayý çalýþtýrýn:**
```bash
dotnet run
```

5. **Swagger UI'a eriþin:**
```
http://localhost:5002
```

## API Endpoints

### Excel Controller (/api/excel)
- `GET /test` - API durumunu test et
- `GET /files` - Yüklenmiþ dosya listesi
- `POST /upload` - Excel dosyasý yükle
- `POST /read/{fileName}` - Excel dosyasýný oku ve veritabanýna kaydet
- `GET /data/{fileName}` - Excel verilerini getir (sayfalama)
- `PUT /data` - Tekil veri güncelleme
- `PUT /data/bulk` - Toplu veri güncelleme
- `POST /data` - Yeni satýr ekle
- `DELETE /data/{id}` - Veri sil (soft delete)
- `POST /export` - Excel export et
- `GET /sheets/{fileName}` - Dosyadaki sheet listesi
- `GET /statistics/{fileName}` - Dosya istatistikleri

### Comparison Controller (/api/comparison)
- `GET /test` - API durumunu test et
- `POST /files` - Ýki dosyayý karþýlaþtýr
- `POST /versions` - Ayný dosyanýn farklý versiyonlarýný karþýlaþtýr
- `GET /changes/{fileName}` - Tarih aralýðýndaki deðiþiklikler
- `GET /history/{fileName}` - Dosya deðiþiklik geçmiþi
- `GET /row-history/{rowId}` - Belirli satýrýn geçmiþi

## Kullaným Örnekleri

### Excel Dosyasý Yükleme
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

### Veri Güncelleme
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

### Yeni Satýr Ekleme
```javascript
const newRowData = {
    fileName: "dosya_adi.xlsx",
    sheetName: "Sheet1",
    rowData: {
        "Ad": "Yeni Kiþi",
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

### Dosya Karþýlaþtýrma
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
    rowIds: [1, 2, 3], // Opsiyonel - belirli satýrlarý export et
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

## Veritabaný Yapýsý

### ExcelFiles Tablosu
- Id, FileName, OriginalFileName, FilePath
- FileSize, UploadDate, UploadedBy, IsActive

### ExcelDataRows Tablosu
- Id, FileName, SheetName, RowIndex
- RowData (JSON), CreatedDate, ModifiedDate
- IsDeleted, Version, ModifiedBy

## Frontend Entegrasyonu

Angular frontend için:
```typescript
// Service örneði
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

API'nin çalýþýp çalýþmadýðýný test etmek için:
- `GET http://localhost:5002/api/excel/test`
- `GET http://localhost:5002/api/comparison/test`

## Lisans

Bu proje EPPlus Non-Commercial lisansý kullanmaktadýr.