# 📊 Excel Data Management API - Proje Özeti

## 🎯 Proje Tanımı
**Excel Data Management API**, Excel dosyalarını yönetmek, düzenlemek ve karşılaştırmak için geliştirilmiş kapsamlı bir .NET 9 Web API projesidir.

---

## 🛠️ Teknoloji Stack
- **.NET 9** - Ana framework
- **ASP.NET Core Web API** - API geliştirme
- **Entity Framework Core** - ORM ve veritabanı işlemleri
- **SQL Server** - Veritabanı
- **EPPlus** - Excel dosya işlemleri
- **Swagger/OpenAPI** - API dokümantasyonu
- **C# 13.0** - Programlama dili

---

## 📁 Proje Yapısı

### 🎮 Controllers (API Katmanı)
| Controller | Dosya Yolu | Görevler |
|------------|------------|----------|
| **ExcelController** | `Controllers/ExcelController.cs` | Excel dosya yönetimi, CRUD işlemleri |
| **ComparisonController** | `Controllers/ComparisonController.cs` | Dosya karşılaştırma işlemleri |

### ⚙️ Services (İş Mantığı Katmanı)
| Interface | Implementation | Görevler |
|-----------|----------------|----------|
| `IExcelService` | `ExcelService` | Excel dosya işlemleri |
| `IDataComparisonService` | `DataComparisonService` | Karşılaştırma algoritmaları |

### 📊 Models (Veri Modelleri)
| Model | Açıklama |
|-------|----------|
| `ExcelDataRow` | Veritabanındaki Excel satır verisi |
| `ExcelFile` | Yüklenen dosya bilgileri |
| `ExcelDataDTOs.cs` | Tüm DTO sınıfları |

### 🗄️ Data (Veritabanı Katmanı)
| Dosya | Görev |
|-------|-------|
| `ExcelDataContext.cs` | Entity Framework DbContext |

---

## 🌐 API Endpoints

### 📤 Excel Yönetimi (`/api/excel`)
| HTTP | Endpoint | Açıklama |
|------|----------|----------|
| `GET` | `/test` | API sağlık kontrolü |
| `GET` | `/files` | Yüklü dosyaların listesi |
| `POST` | `/upload` | Excel dosyası yükleme |
| `POST` | `/read-from-file` | Manuel dosya seçip okuma |
| `POST` | `/update-from-file` | Manuel dosya seçip güncelleme |
| `POST` | `/read/{fileName}` | Yüklü dosyadan veri okuma |
| `GET` | `/data/{fileName}` | Sayfalama ile veri getirme |
| `PUT` | `/data` | Tek satır güncelleme |
| `PUT` | `/data/bulk` | Toplu güncelleme |
| `POST` | `/data` | Yeni satır ekleme |
| `DELETE` | `/data/{id}` | Satır silme |
| `POST` | `/export` | Excel'e dışa aktarma |
| `GET` | `/sheets/{fileName}` | Dosyadaki sheet'leri listeleme |
| `GET` | `/statistics/{fileName}` | Veri istatistikleri |

### 🔍 Karşılaştırma (`/api/comparison`)
| HTTP | Endpoint | Açıklama |
|------|----------|----------|
| `GET` | `/test` | API sağlık kontrolü |
| `POST` | `/compare-from-files` | Manuel dosya karşılaştırma |
| `POST` | `/files` | Yüklü dosya karşılaştırma |
| `POST` | `/versions` | Versiyon karşılaştırma |
| `GET` | `/changes/{fileName}` | Değişiklikleri getirme |
| `GET` | `/history/{fileName}` | Değişiklik geçmişi |
| `GET` | `/row-history/{rowId}` | Satır geçmişi |
| `POST` | `/snapshot-compare` | Snapshot karşılaştırma |

---

## 🗃️ Veritabanı Yapısı

### ExcelFiles Tablosu
```sql
Id (int, PK)               - Benzersiz dosya ID'si
FileName (string, 255)     - Sistem dosya adı
OriginalFileName (string)  - Orijinal dosya adı
FilePath (string, 500)     - Dosya yolu
FileSize (long)            - Dosya boyutu (byte)
UploadDate (datetime)      - Yükleme tarihi
UploadedBy (string, 255)   - Yükleyen kişi
IsActive (bool)            - Aktif durumu
```

### ExcelDataRows Tablosu
```sql
Id (int, PK)               - Benzersiz satır ID'si
FileName (string, 255)     - Hangi dosyadan
SheetName (string, 255)    - Hangi sheet'ten
RowIndex (int)             - Satır numarası
RowData (nvarchar(max))    - JSON formatında satır verisi
CreatedDate (datetime)     - Oluşturulma tarihi
ModifiedDate (datetime?)   - Son değişiklik tarihi
IsDeleted (bool)           - Silinmiş mi?
Version (int)              - Versiyon numarası
ModifiedBy (string, 255)   - Son değiştiren kişi
```

---

## 🚀 Ana Özellikler

### 📁 Dosya İşlemleri
- ✅ Excel dosyası yükleme (.xlsx, .xls)
- ✅ Dosya validasyonu ve güvenlik kontrolleri
- ✅ 100MB'a kadar dosya desteği
- ✅ Çoklu sheet desteği

### 📊 Veri Yönetimi
- ✅ Excel verilerini veritabanına aktarma
- ✅ Sayfalama ile performanslı listeleme
- ✅ CRUD işlemleri (Create, Read, Update, Delete)
- ✅ Toplu güncelleme işlemleri
- ✅ Veri export işlemleri

### 🔍 Karşılaştırma ve Analiz
- ✅ İki farklı dosyayı karşılaştırma
- ✅ Aynı dosyanın farklı versiyonlarını karşılaştırma
- ✅ Detaylı değişiklik analizi
- ✅ Değişiklik geçmişi takibi
- ✅ Satır bazında geçmiş izleme

### 📈 İstatistik ve Raporlama
- ✅ Toplam satır sayısı
- ✅ Değiştirilmiş satır istatistikleri
- ✅ Eklenen/silinen satır sayıları
- ✅ Tarih bazlı filtreleme

---

## 🛡️ Güvenlik ve Performans

### Güvenlik
- ✅ Dosya uzantısı kontrolü
- ✅ Dosya boyutu sınırlaması
- ✅ SQL Injection korunması (EF Core)
- ✅ CORS politikaları

### Performans
- ✅ Sayfalama ile bellek optimizasyonu
- ✅ Veritabanı indexleri
- ✅ Asenkron işlemler
- ✅ Toplu işlemler için optimize edilmiş sorgular

---

## 🔧 Konfigürasyon

### Bağlantı Bilgileri
- **Swagger UI**: `http://localhost:5002/swagger`
- **HTTPS Swagger**: `https://localhost:7002/swagger`
- **API Base URL**: `http://localhost:5002/api`

### Önemli Ayarlar
- **Maksimum dosya boyutu**: 100MB
- **Desteklenen formatlar**: .xlsx, .xls
- **Varsayılan sayfa boyutu**: 50 kayıt
- **EPPlus lisansı**: NonCommercial

---

## 📱 Kullanım Senaryoları

### 1. Temel Excel İşlemleri
```
1. Dosya yükle → POST /api/excel/upload
2. Veriyi oku → POST /api/excel/read/{fileName}
3. Verileri görüntüle → GET /api/excel/data/{fileName}
4. Veri düzenle → PUT /api/excel/data
```

### 2. Karşılaştırma İşlemleri
```
1. İki dosya yükle → POST /api/excel/upload (2 kez)
2. Karşılaştır → POST /api/comparison/files
3. Sonuçları analiz et → GET /api/comparison/history/{fileName}
```

### 3. Manuel Dosya İşlemleri
```
1. Bilgisayardan dosya seç → POST /api/excel/read-from-file
2. Anında işle ve sonuçları al
3. Gerekirse karşılaştır → POST /api/comparison/compare-from-files
```

---

## 🚦 Hata Yönetimi
- ✅ Detaylı hata mesajları
- ✅ HTTP status kodları
- ✅ Loglama sistemi
- ✅ Try-catch yapıları
- ✅ Kullanıcı dostu hata mesajları

---

## 🎯 Sunum İçin Önemli Noktalar

### 💪 Güçlü Yanlar
1. **Kapsamlı API**: Hem temel hem gelişmiş Excel işlemleri
2. **Performans**: Sayfalama ve optimize edilmiş sorgular
3. **Güvenlik**: Dosya kontrolleri ve SQL korunması
4. **Flexibilite**: Manuel ve otomatik dosya işleme
5. **Traceability**: Tam değişiklik geçmişi

### 🎨 Teknik Kalite
1. **Clean Architecture**: Katmanlı yapı
2. **SOLID Principles**: Interface-based design
3. **Modern Technology**: .NET 9, EF Core
4. **API Documentation**: Swagger entegrasyonu
5. **Error Handling**: Kapsamlı hata yönetimi

### 📊 İş Değeri
1. **Verimlilik**: Excel işlemlerini otomatikleştirme
2. **Doğruluk**: Karşılaştırma ve analiz yetenekleri
3. **İzlenebilirlik**: Değişiklik geçmişi
4. **Ölçeklenebilirlik**: Büyük dosyalar için optimizasyon
5. **Entegrasyon**: REST API ile kolay entegrasyon

---

**Geliştirici**: Taha Teke  
**Teknoloji**: .NET 9, ASP.NET Core Web API  
**Tarih**: 2024  
**Lisans**: EPPlus NonCommercial