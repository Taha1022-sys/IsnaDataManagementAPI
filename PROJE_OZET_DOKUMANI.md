# Excel Data Management Sistemi - Proje �zet Dok�man�

## ?? Genel Bak��

Bu proje, Excel dosyalar�n� y�netmek, kar��la�t�rmak ve version kontrol� yapmak i�in geli�tirilmi� kapsaml� bir web uygulamas�d�r. Backend API (.NET 9.0) ve Frontend (HTML/Vite) olmak �zere iki ana bile�enden olu�ur.

## ??? Proje Mimarisi

### Backend: ASP.NET Core Web API (.NET 9.0)
- **Port**: 5002 (HTTP), 7002 (HTTPS)
- **Veritaban�**: SQL Server (Entity Framework Core)
- **Excel ��lemleri**: EPPlus k�t�phanesi
- **API Dok�mantasyonu**: Swagger UI

### Frontend: Vite-based Web Application
- **Framework**: HTML5 + TypeScript/JavaScript
- **Build Tool**: Vite
- **Port**: Dinamik (genellikle 5173)

## ?? Dosya Yap�s� ve A��klamalar�

### Backend Dosyalar�

#### ?? **Program.cs** - Ana Uygulama Konfig�rasyonu
**Ne i�e yarar**: Uygulaman�n ba�lang�� noktas� ve t�m servislerin konfig�rasyonu
```
- Web API portlar�n� ayarlar (5002, 7002)
- Veritaban� ba�lant�s�n� kurar
- CORS politikalar�n� tan�mlar (Frontend ba�lant�s� i�in)
- Swagger UI'yi aktif eder
- Dependency Injection yap�land�rmas�
- Veritaban�n� otomatik olu�turur
```

#### ?? **Models/ExcelDataModels.cs** - Veri Modelleri
**Ne i�e yarar**: Veritaban� tablolar�n�n C# s�n�f kar��l�klar�

**ExcelFile Modeli:**
- Excel dosyalar�n�n metadata bilgilerini tutar
- Dosya ad�, yol, boyut, y�kleme tarihi
- Dosya aktif/pasif durumu

**ExcelDataRow Modeli:**
- Excel sat�rlar�n�n veri i�eri�ini JSON format�nda saklar
- Version kontrol� i�in s�r�m numaras�
- Sat�r baz�nda de�i�iklik takibi
- Kim taraf�ndan de�i�tirildi�i bilgisi

#### ?? **Models/DTOs/ExcelDataDTOs.cs** - Veri Transfer Nesneleri
**Ne i�e yarar**: API ile frontend aras�nda veri al��veri�i i�in kullan�lan �zel s�n�flar

**�nemli DTO'lar:**
- `ExcelDataResponseDto`: Frontend'e g�nderilen Excel verisi format�
- `DataComparisonResultDto`: Dosya kar��la�t�rma sonu�lar�
- `BulkUpdateDto`: Toplu veri g�ncelleme i�in
- `ExcelExportRequestDto`: Excel export i�lemleri i�in

#### ??? **Data/ExcelDataContext.cs** - Veritaban� Ba�lam�
**Ne i�e yarar**: Entity Framework ile veritaban� i�lemleri
```
- ExcelFiles ve ExcelDataRows tablolar�n� tan�mlar
- Tablo yap�s�n� ve ili�kileri ayarlar
- Index'leri performans i�in optimize eder
- Veritaban� sorgular� i�in temel sa�lar
```

#### ??? **Services/ExcelService.cs** - Ana Excel ��lem Servisi
**Ne i�e yarar**: Excel dosyalar�yla ilgili t�m business logic'i i�erir

**Ana Fonksiyonlar:**
- `UploadExcelFileAsync()`: Dosya y�kleme ve kaydetme
- `ReadExcelDataAsync()`: Excel i�eri�ini okuyup veritaban�na aktarma
- `GetExcelDataAsync()`: Sayfalama ile veri getirme
- `UpdateExcelDataAsync()`: Tekil veri g�ncelleme
- `BulkUpdateExcelDataAsync()`: Toplu veri g�ncelleme
- `ExportToExcelAsync()`: Veritaban�ndan Excel dosyas� olu�turma
- `GetSheetsAsync()`: Excel'deki sheet'leri listeleme

**Teknik Detaylar:**
- EPPlus k�t�phanesi kullanarak Excel okuma/yazma
- JSON format�nda veri saklama (esnek yap� i�in)
- Version kontrol� (her de�i�iklikte versiyon art���)
- Soft delete (fiziksel silme yapmaz, sadece i�aretler)

#### ?? **Services/DataComparisonService.cs** - Veri Kar��la�t�rma Servisi
**Ne i�e yarar**: Dosyalar ve versiyonlar aras� kar��la�t�rma i�lemleri

**Ana Fonksiyonlar:**
- `CompareFilesAsync()`: �ki farkl� dosyay� kar��la�t�rma
- `CompareVersionsAsync()`: Ayn� dosyan�n farkl� versiyonlar�n� kar��la�t�rma
- `GetChangesAsync()`: Belirli tarih aral���nda de�i�iklikleri getirme
- `GetChangeHistoryAsync()`: Dosya de�i�iklik ge�mi�i
- `GetRowHistoryAsync()`: Sat�r baz�nda de�i�iklik ge�mi�i

**Kar��la�t�rma T�rleri:**
- `Modified`: De�i�tirilmi� veriler
- `Added`: Yeni eklenen sat�rlar
- `Deleted`: Silinen sat�rlar

#### ?? **Controllers/ExcelController.cs** - Ana API Controller
**Ne i�e yarar**: Excel i�lemleri i�in HTTP endpoint'leri sa�lar

**API Endpoints:**
- `GET /api/excel/files`: Dosya listesi
- `POST /api/excel/upload`: Dosya y�kleme
- `GET /api/excel/data/{fileName}`: Veri g�r�nt�leme
- `PUT /api/excel/data`: Veri g�ncelleme
- `POST /api/excel/data`: Yeni sat�r ekleme
- `DELETE /api/excel/data/{id}`: Veri silme
- `POST /api/excel/export`: Excel export
- `GET /api/excel/sheets/{fileName}`: Sheet listesi

#### ?? **Controllers/ComparisonController.cs** - Kar��la�t�rma API Controller
**Ne i�e yarar**: Dosya kar��la�t�rma i�lemleri i�in HTTP endpoint'leri

**API Endpoints:**
- `POST /api/comparison/files`: �ki dosyay� kar��la�t�r
- `POST /api/comparison/versions`: Versiyonlar� kar��la�t�r
- `GET /api/comparison/changes/{fileName}`: De�i�iklikleri getir
- `GET /api/comparison/history/{fileName}`: Dosya ge�mi�i
- `GET /api/comparison/row-history/{rowId}`: Sat�r ge�mi�i

### Frontend Dosyalar�

#### ?? **index.html** - Ana HTML Dosyas�
**Ne i�e yarar**: Uygulaman�n giri� noktas�
```html
- React/TypeScript uygulamas� i�in container
- Vite build sistemi entegrasyonu
- Meta taglar ve ba�l�k ayarlar�
- Root div elementi (React bile�enleri burada render edilir)
```

## ?? Teknik �zellikler

### Veritaban� Yap�s�
**ExcelFiles Tablosu:**
- Dosya metadata's�
- Y�kleme bilgileri
- Aktif/pasif durumu

**ExcelDataRows Tablosu:**
- Sat�r baz�nda veri saklama
- JSON format�nda esnek i�erik
- Version takibi
- Soft delete deste�i
- De�i�iklik audit trail'i

### Excel ��leme Sistemi
1. **Dosya Y�kleme**: Multipart/form-data ile dosya al�m�
2. **Veri Okuma**: EPPlus ile Excel parsing
3. **Veri Saklama**: JSON format�nda esnek saklama
4. **Version Kontrol�**: Her de�i�iklikte otomatik versiyon art���

### API �zellikleri
- **RESTful Design**: Standard HTTP methodlar�
- **CORS Deste�i**: Frontend entegrasyonu i�in
- **Swagger UI**: Otomatik API dok�mantasyonu
- **Error Handling**: Kapsaml� hata y�netimi
- **Logging**: Detayl� i�lem loglar�

### G�venlik ve Performans
- **Validation**: Model validation ile veri do�rulama
- **Pagination**: B�y�k veri setleri i�in sayfalama
- **Indexing**: Veritaban� performans optimizasyonu
- **Async/Await**: Non-blocking i�lemler

## ?? �al��ma Ak���

### 1. Dosya Y�kleme S�reci
```
Frontend ? POST /api/excel/upload ? ExcelService.UploadExcelFileAsync()
? Dosya kaydetme ? Veritaban�na metadata kayd� ? Response
```

### 2. Veri Okuma S�reci
```
Frontend ? POST /api/excel/read/{fileName} ? ExcelService.ReadExcelDataAsync()
? EPPlus ile Excel okuma ? JSON'a �evirme ? Veritaban�na kaydetme
```

### 3. Veri G�ncelleme S�reci
```
Frontend ? PUT /api/excel/data ? ExcelService.UpdateExcelDataAsync()
? Mevcut veriyi bulma ? Version art�rma ? G�ncelleme ? Audit trail
```

### 4. Kar��la�t�rma S�reci
```
Frontend ? POST /api/comparison/files ? DataComparisonService.CompareFilesAsync()
? Her iki dosyay� okuma ? Sat�r sat�r kar��la�t�rma ? Farklar� tespit etme
```

## ?? Kullan�lan Teknolojiler

### Backend
- **.NET 9.0**: Modern C# framework
- **ASP.NET Core**: Web API framework
- **Entity Framework Core**: ORM
- **SQL Server**: Veritaban�
- **EPPlus**: Excel i�lemleri
- **Swagger**: API dok�mantasyonu

### Frontend
- **Vite**: Modern build tool
- **TypeScript**: Type-safe JavaScript
- **HTML5**: Modern web standartlar�

### Paketler ve K�t�phaneler
- **EPPlus 7.5.0**: Excel okuma/yazma
- **Microsoft.EntityFrameworkCore.SqlServer**: SQL Server entegrasyonu
- **Swashbuckle.AspNetCore**: Swagger UI

## ?? Kullan�m Senaryolar�

### 1. Excel Dosyas� Y�netimi
- Dosya y�kleme ve saklama
- �oklu sheet deste�i
- Veri g�r�nt�leme ve d�zenleme

### 2. Version Kontrol�
- De�i�iklik takibi
- Sat�r baz�nda ge�mi�
- Kim ne zaman de�i�tirmi� bilgisi

### 3. Veri Kar��la�t�rma
- �ki dosya aras�nda fark analizi
- Zaman bazl� versiyon kar��la�t�rmas�
- De�i�iklik raporlama

### 4. Veri Export
- Filtrelenmi� veri export'u
- Belirli sat�rlar� se�erek export
- Ge�mi� bilgileri dahil export

## ?? Geli�tirme D�ng�s�

### 1. Backend API Geli�tirme
- Controller'da endpoint tan�mlama
- Service'de business logic yazma
- Model ve DTO tan�mlama
- Test etme (Swagger UI ile)

### 2. Frontend Entegrasyonu
- API �a�r�lar� yapma
- Kullan�c� aray�z� olu�turma
- Veri binding i�lemleri

### 3. Test ve Debug
- API endpoint'lerini test etme
- Hata durumlar�n� kontrol etme
- Performance optimizasyonu

