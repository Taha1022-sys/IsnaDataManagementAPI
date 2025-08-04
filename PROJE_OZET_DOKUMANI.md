# Excel Data Management Sistemi - Proje Özet Dokümaný

## ?? Genel Bakýþ

Bu proje, Excel dosyalarýný yönetmek, karþýlaþtýrmak ve version kontrolü yapmak için geliþtirilmiþ kapsamlý bir web uygulamasýdýr. Backend API (.NET 9.0) ve Frontend (HTML/Vite) olmak üzere iki ana bileþenden oluþur.

## ??? Proje Mimarisi

### Backend: ASP.NET Core Web API (.NET 9.0)
- **Port**: 5002 (HTTP), 7002 (HTTPS)
- **Veritabaný**: SQL Server (Entity Framework Core)
- **Excel Ýþlemleri**: EPPlus kütüphanesi
- **API Dokümantasyonu**: Swagger UI

### Frontend: Vite-based Web Application
- **Framework**: HTML5 + TypeScript/JavaScript
- **Build Tool**: Vite
- **Port**: Dinamik (genellikle 5173)

## ?? Dosya Yapýsý ve Açýklamalarý

### Backend Dosyalarý

#### ?? **Program.cs** - Ana Uygulama Konfigürasyonu
**Ne iþe yarar**: Uygulamanýn baþlangýç noktasý ve tüm servislerin konfigürasyonu
```
- Web API portlarýný ayarlar (5002, 7002)
- Veritabaný baðlantýsýný kurar
- CORS politikalarýný tanýmlar (Frontend baðlantýsý için)
- Swagger UI'yi aktif eder
- Dependency Injection yapýlandýrmasý
- Veritabanýný otomatik oluþturur
```

#### ?? **Models/ExcelDataModels.cs** - Veri Modelleri
**Ne iþe yarar**: Veritabaný tablolarýnýn C# sýnýf karþýlýklarý

**ExcelFile Modeli:**
- Excel dosyalarýnýn metadata bilgilerini tutar
- Dosya adý, yol, boyut, yükleme tarihi
- Dosya aktif/pasif durumu

**ExcelDataRow Modeli:**
- Excel satýrlarýnýn veri içeriðini JSON formatýnda saklar
- Version kontrolü için sürüm numarasý
- Satýr bazýnda deðiþiklik takibi
- Kim tarafýndan deðiþtirildiði bilgisi

#### ?? **Models/DTOs/ExcelDataDTOs.cs** - Veri Transfer Nesneleri
**Ne iþe yarar**: API ile frontend arasýnda veri alýþveriþi için kullanýlan özel sýnýflar

**Önemli DTO'lar:**
- `ExcelDataResponseDto`: Frontend'e gönderilen Excel verisi formatý
- `DataComparisonResultDto`: Dosya karþýlaþtýrma sonuçlarý
- `BulkUpdateDto`: Toplu veri güncelleme için
- `ExcelExportRequestDto`: Excel export iþlemleri için

#### ??? **Data/ExcelDataContext.cs** - Veritabaný Baðlamý
**Ne iþe yarar**: Entity Framework ile veritabaný iþlemleri
```
- ExcelFiles ve ExcelDataRows tablolarýný tanýmlar
- Tablo yapýsýný ve iliþkileri ayarlar
- Index'leri performans için optimize eder
- Veritabaný sorgularý için temel saðlar
```

#### ??? **Services/ExcelService.cs** - Ana Excel Ýþlem Servisi
**Ne iþe yarar**: Excel dosyalarýyla ilgili tüm business logic'i içerir

**Ana Fonksiyonlar:**
- `UploadExcelFileAsync()`: Dosya yükleme ve kaydetme
- `ReadExcelDataAsync()`: Excel içeriðini okuyup veritabanýna aktarma
- `GetExcelDataAsync()`: Sayfalama ile veri getirme
- `UpdateExcelDataAsync()`: Tekil veri güncelleme
- `BulkUpdateExcelDataAsync()`: Toplu veri güncelleme
- `ExportToExcelAsync()`: Veritabanýndan Excel dosyasý oluþturma
- `GetSheetsAsync()`: Excel'deki sheet'leri listeleme

**Teknik Detaylar:**
- EPPlus kütüphanesi kullanarak Excel okuma/yazma
- JSON formatýnda veri saklama (esnek yapý için)
- Version kontrolü (her deðiþiklikte versiyon artýþý)
- Soft delete (fiziksel silme yapmaz, sadece iþaretler)

#### ?? **Services/DataComparisonService.cs** - Veri Karþýlaþtýrma Servisi
**Ne iþe yarar**: Dosyalar ve versiyonlar arasý karþýlaþtýrma iþlemleri

**Ana Fonksiyonlar:**
- `CompareFilesAsync()`: Ýki farklý dosyayý karþýlaþtýrma
- `CompareVersionsAsync()`: Ayný dosyanýn farklý versiyonlarýný karþýlaþtýrma
- `GetChangesAsync()`: Belirli tarih aralýðýnda deðiþiklikleri getirme
- `GetChangeHistoryAsync()`: Dosya deðiþiklik geçmiþi
- `GetRowHistoryAsync()`: Satýr bazýnda deðiþiklik geçmiþi

**Karþýlaþtýrma Türleri:**
- `Modified`: Deðiþtirilmiþ veriler
- `Added`: Yeni eklenen satýrlar
- `Deleted`: Silinen satýrlar

#### ?? **Controllers/ExcelController.cs** - Ana API Controller
**Ne iþe yarar**: Excel iþlemleri için HTTP endpoint'leri saðlar

**API Endpoints:**
- `GET /api/excel/files`: Dosya listesi
- `POST /api/excel/upload`: Dosya yükleme
- `GET /api/excel/data/{fileName}`: Veri görüntüleme
- `PUT /api/excel/data`: Veri güncelleme
- `POST /api/excel/data`: Yeni satýr ekleme
- `DELETE /api/excel/data/{id}`: Veri silme
- `POST /api/excel/export`: Excel export
- `GET /api/excel/sheets/{fileName}`: Sheet listesi

#### ?? **Controllers/ComparisonController.cs** - Karþýlaþtýrma API Controller
**Ne iþe yarar**: Dosya karþýlaþtýrma iþlemleri için HTTP endpoint'leri

**API Endpoints:**
- `POST /api/comparison/files`: Ýki dosyayý karþýlaþtýr
- `POST /api/comparison/versions`: Versiyonlarý karþýlaþtýr
- `GET /api/comparison/changes/{fileName}`: Deðiþiklikleri getir
- `GET /api/comparison/history/{fileName}`: Dosya geçmiþi
- `GET /api/comparison/row-history/{rowId}`: Satýr geçmiþi

### Frontend Dosyalarý

#### ?? **index.html** - Ana HTML Dosyasý
**Ne iþe yarar**: Uygulamanýn giriþ noktasý
```html
- React/TypeScript uygulamasý için container
- Vite build sistemi entegrasyonu
- Meta taglar ve baþlýk ayarlarý
- Root div elementi (React bileþenleri burada render edilir)
```

## ?? Teknik Özellikler

### Veritabaný Yapýsý
**ExcelFiles Tablosu:**
- Dosya metadata'sý
- Yükleme bilgileri
- Aktif/pasif durumu

**ExcelDataRows Tablosu:**
- Satýr bazýnda veri saklama
- JSON formatýnda esnek içerik
- Version takibi
- Soft delete desteði
- Deðiþiklik audit trail'i

### Excel Ýþleme Sistemi
1. **Dosya Yükleme**: Multipart/form-data ile dosya alýmý
2. **Veri Okuma**: EPPlus ile Excel parsing
3. **Veri Saklama**: JSON formatýnda esnek saklama
4. **Version Kontrolü**: Her deðiþiklikte otomatik versiyon artýþý

### API Özellikleri
- **RESTful Design**: Standard HTTP methodlarý
- **CORS Desteði**: Frontend entegrasyonu için
- **Swagger UI**: Otomatik API dokümantasyonu
- **Error Handling**: Kapsamlý hata yönetimi
- **Logging**: Detaylý iþlem loglarý

### Güvenlik ve Performans
- **Validation**: Model validation ile veri doðrulama
- **Pagination**: Büyük veri setleri için sayfalama
- **Indexing**: Veritabaný performans optimizasyonu
- **Async/Await**: Non-blocking iþlemler

## ?? Çalýþma Akýþý

### 1. Dosya Yükleme Süreci
```
Frontend ? POST /api/excel/upload ? ExcelService.UploadExcelFileAsync()
? Dosya kaydetme ? Veritabanýna metadata kaydý ? Response
```

### 2. Veri Okuma Süreci
```
Frontend ? POST /api/excel/read/{fileName} ? ExcelService.ReadExcelDataAsync()
? EPPlus ile Excel okuma ? JSON'a çevirme ? Veritabanýna kaydetme
```

### 3. Veri Güncelleme Süreci
```
Frontend ? PUT /api/excel/data ? ExcelService.UpdateExcelDataAsync()
? Mevcut veriyi bulma ? Version artýrma ? Güncelleme ? Audit trail
```

### 4. Karþýlaþtýrma Süreci
```
Frontend ? POST /api/comparison/files ? DataComparisonService.CompareFilesAsync()
? Her iki dosyayý okuma ? Satýr satýr karþýlaþtýrma ? Farklarý tespit etme
```

## ?? Kullanýlan Teknolojiler

### Backend
- **.NET 9.0**: Modern C# framework
- **ASP.NET Core**: Web API framework
- **Entity Framework Core**: ORM
- **SQL Server**: Veritabaný
- **EPPlus**: Excel iþlemleri
- **Swagger**: API dokümantasyonu

### Frontend
- **Vite**: Modern build tool
- **TypeScript**: Type-safe JavaScript
- **HTML5**: Modern web standartlarý

### Paketler ve Kütüphaneler
- **EPPlus 7.5.0**: Excel okuma/yazma
- **Microsoft.EntityFrameworkCore.SqlServer**: SQL Server entegrasyonu
- **Swashbuckle.AspNetCore**: Swagger UI

## ?? Kullaným Senaryolarý

### 1. Excel Dosyasý Yönetimi
- Dosya yükleme ve saklama
- Çoklu sheet desteði
- Veri görüntüleme ve düzenleme

### 2. Version Kontrolü
- Deðiþiklik takibi
- Satýr bazýnda geçmiþ
- Kim ne zaman deðiþtirmiþ bilgisi

### 3. Veri Karþýlaþtýrma
- Ýki dosya arasýnda fark analizi
- Zaman bazlý versiyon karþýlaþtýrmasý
- Deðiþiklik raporlama

### 4. Veri Export
- Filtrelenmiþ veri export'u
- Belirli satýrlarý seçerek export
- Geçmiþ bilgileri dahil export

## ?? Geliþtirme Döngüsü

### 1. Backend API Geliþtirme
- Controller'da endpoint tanýmlama
- Service'de business logic yazma
- Model ve DTO tanýmlama
- Test etme (Swagger UI ile)

### 2. Frontend Entegrasyonu
- API çaðrýlarý yapma
- Kullanýcý arayüzü oluþturma
- Veri binding iþlemleri

### 3. Test ve Debug
- API endpoint'lerini test etme
- Hata durumlarýný kontrol etme
- Performance optimizasyonu

