# 📊 Excel Data Management System - Comprehensive Documentation

## 🚀 Proje Genel Bakış

Excel Data Management System, Excel dosyalarını web üzerinden yönetmek, düzenlemek ve karşılaştırmak için geliştirilmiş full-stack bir uygulamadır. ASP.NET Core 9.0 backend ve React TypeScript frontend'i ile modern web teknolojilerini kullanır.

## 📁 Proje Yapısı

```
ISNA DATA MANAGEMENT PROJECT/
├── 🗂️ ExcelDataManagementAPI/          # Backend (ASP.NET Core 9.0)
│   ├── 🎮 Controllers/                   # API Controller'ları
│   │   ├── ExcelController.cs           # Excel işlemleri
│   │   └── ComparisonController.cs      # Karşılaştırma işlemleri
│   ├── 🔧 Services/                     # İş mantığı katmanı
│   │   ├── ExcelService.cs             # Excel işlem servisi
│   │   ├── IExcelService.cs            # Excel servis interface'i
│   │   ├── DataComparisonService.cs    # Karşılaştırma servisi
│   │   └── IDataComparisonService.cs   # Karşılaştırma interface'i
│   ├── 🏗️ Models/                       # Veri modelleri
│   │   ├── ExcelDataModels.cs          # Ana entity'ler
│   │   └── DTOs/                       # Data Transfer Objects
│   │       └── ExcelDataDTOs.cs        # DTO sınıfları
│   ├── 🗄️ Data/                         # Veritabanı katmanı
│   │   └── ExcelDataContext.cs         # Entity Framework DbContext
│   ├── ⚙️ appsettings.json              # Konfigürasyon
│   ├── 🚀 Program.cs                    # Uygulama başlangıç noktası
│   └── 📦 ExcelDataManagementAPI.csproj # Proje dosyası
├── 🌐 Frontend Files/                   # Frontend (React TypeScript)
│   ├── 🎣 hooks/                        # React Hook'ları
│   │   └── useExcelApi.ts              # Excel API hook'u
│   ├── 🔗 services/                     # API servis katmanı
│   │   ├── ExcelApiService.ts          # Ana API servisi
│   │   └── ExcelApiServiceEnhanced.ts  # Gelişmiş API servisi
│   ├── 🧩 components/                   # React bileşenleri
│   │   ├── ExcelComparison.tsx         # Karşılaştırma bileşeni
│   │   └── ExcelComparisonOptimized.tsx # Optimize edilmiş karşılaştırma
│   ├── 🏷️ types/                        # TypeScript tip tanımları
│   │   └── excel-types.ts              # Excel ile ilgili tipler
│   ├── ⚙️ config/                       # Konfigürasyon
│   │   └── environment.ts              # Ortam değişkenleri
│   ├── 📄 index.html                   # HTML demo sayfası
│   ├── 📦 package.json                 # NPM bağımlılıkları
│   ├── 🔧 vite.config.ts               # Vite build konfigürasyonu
│   └── 📘 tsconfig.json                # TypeScript konfigürasyonu
└── 📚 README.md                        # Bu dokümantasyon dosyası
```

## 💻 Teknoloji Stack'i

### 🔧 Backend Stack
- **ASP.NET Core 9.0** - Modern web API framework
- **Entity Framework Core 9.0** - ORM (Object-Relational Mapping)
- **SQL Server** - Veritabanı (LocalDB destekli)
- **EPPlus 7.5.0** - Excel dosya işleme kütüphanesi
- **Swagger/OpenAPI** - API dokümantasyonu
- **Microsoft.Extensions.Logging** - Loglama

### 🌐 Frontend Stack
- **React 18** - Modern UI kütüphanesi
- **TypeScript** - Type-safe JavaScript
- **Vite** - Hızlı build tool
- **Axios** - HTTP client
- **Custom Hooks** - Durum yönetimi

## 🗄️ Veritabanı Yapısı

### 📊 ExcelFile Entity
```csharp
public class ExcelFile
{
    public int Id { get; set; }                    // Primary Key
    public string FileName { get; set; }           // Sistem dosya adı
    public string OriginalFileName { get; set; }   // Orijinal dosya adı
    public string FilePath { get; set; }           // Dosya yolu
    public long FileSize { get; set; }             // Dosya boyutu (byte)
    public DateTime UploadDate { get; set; }       // Yükleme tarihi
    public string? UploadedBy { get; set; }        // Yükleyen kullanıcı
    public bool IsActive { get; set; }             // Aktif durumu
}
```

### 📈 ExcelDataRow Entity
```csharp
public class ExcelDataRow
{
    public int Id { get; set; }                    // Primary Key
    public string FileName { get; set; }           // Dosya adı referansı
    public string SheetName { get; set; }          // Sheet adı
    public int RowIndex { get; set; }              // Satır numarası
    public string RowData { get; set; }            // JSON formatında satır verisi
    public DateTime CreatedDate { get; set; }      // Oluşturulma tarihi
    public DateTime? ModifiedDate { get; set; }    // Değiştirilme tarihi
    public bool IsDeleted { get; set; }            // Silinme durumu (soft delete)
    public int Version { get; set; }               // Versiyon numarası
    public string? ModifiedBy { get; set; }        // Değiştiren kullanıcı
}
```

## 🔌 API Endpoints

### 📊 Excel Controller (`/api/excel`)

| Method | Endpoint | Açıklama | Parametreler |
|--------|----------|----------|--------------|
| `GET` | `/test` | API durumu kontrolü | - |
| `GET` | `/files` | Yüklenen dosya listesi | - |
| `POST` | `/upload` | Excel dosyası yükleme | FormFile, uploadedBy |
| `POST` | `/read/{fileName}` | Excel dosyasını okuma | fileName, sheetName? |
| `GET` | `/data/{fileName}` | Veri getirme (sayfalama) | fileName, sheetName?, page, pageSize |
| `PUT` | `/data` | Veri güncelleme | ExcelDataUpdateDto |
| `PUT` | `/data/bulk` | Toplu veri güncelleme | BulkUpdateDto |
| `POST` | `/data` | Yeni satır ekleme | AddRowRequestDto |
| `DELETE` | `/data/{id}` | Veri silme | id, deletedBy? |
| `POST` | `/export` | Excel export | ExcelExportRequestDto |
| `GET` | `/sheets/{fileName}` | Sheet listesi | fileName |
| `GET` | `/statistics/{fileName}` | Dosya istatistikleri | fileName, sheetName? |

### 🔄 Comparison Controller (`/api/comparison`)

| Method | Endpoint | Açıklama | Parametreler |
|--------|----------|----------|--------------|
| `GET` | `/test` | API durumu kontrolü | - |
| `POST` | `/files` | İki dosyayı karşılaştırma | CompareFilesRequestDto |
| `POST` | `/versions` | Versiyon karşılaştırma | CompareVersionsRequestDto |
| `GET` | `/changes/{fileName}` | Değişiklik listesi | fileName, fromDate?, toDate?, sheetName? |
| `GET` | `/history/{fileName}` | Değişiklik geçmişi | fileName, sheetName? |
| `GET` | `/row-history/{rowId}` | Satır geçmişi | rowId |

## 🎯 Ana Özellikler

### 📤 Excel Dosya Yükleme
- ✅ `.xlsx` ve `.xls` format desteği
- ✅ Dosya boyutu ve tip validasyonu
- ✅ Otomatik dosya okuma ve veritabanına kaydetme
- ✅ Kullanıcı tracking (kim yükledi)
- ✅ Dosya metadata yönetimi

### 📊 Veri Yönetimi
- ✅ **CRUD İşlemleri**: Create, Read, Update, Delete
- ✅ **Sayfalama**: Büyük veri setleri için performans
- ✅ **Soft Delete**: Veri güvenliği için geçici silme
- ✅ **Versiyon Kontrolü**: Her değişiklik için versiyon artırma
- ✅ **Audit Trail**: Kim, ne zaman, ne değiştirdi takibi

### 🔄 Dosya Karşılaştırma
- ✅ **İki Dosya Karşılaştırma**: Farklı dosyalar arası karşılaştırma
- ✅ **Versiyon Karşılaştırma**: Aynı dosyanın farklı tarihlerdeki hallerini karşılaştırma
- ✅ **Detaylı Fark Analizi**: Satır ve hücre düzeyinde farklar
- ✅ **İstatistiksel Özet**: Eklenen, silinen, değiştirilen satır sayıları
- ✅ **Değişiklik Geçmişi**: Tüm değişikliklerin kronolojik listesi

### 📥 Excel Export
- ✅ **Filtrelenmiş Export**: Belirli satırları seçerek export
- ✅ **Sheet Bazlı Export**: Sadece belirli sheet'i export etme
- ✅ **Değişiklik Geçmişi**: İsteğe bağlı olarak değişiklik bilgileri dahil etme
- ✅ **Format Korunması**: Orijinal Excel formatını koruma

### 📱 Frontend Özellikleri
- ✅ **React Hooks**: Modern React pattern'leri
- ✅ **TypeScript**: Type safety ve IntelliSense
- ✅ **Real-time Updates**: Değişikliklerin anlık yansıması
- ✅ **Responsive Design**: Mobil uyumlu arayüz
- ✅ **Error Handling**: Kullanıcı dostu hata mesajları

## 🚀 Kurulum ve Çalıştırma

### 📋 Sistem Gereksinimleri

#### Backend
- **Microsoft .NET 9.0 SDK** - [İndir](https://dotnet.microsoft.com/download)
- **SQL Server 2019+** veya **SQL Server LocalDB**
- **Visual Studio 2022** veya **VS Code**

#### Frontend
- **Node.js 18+** - [İndir](https://nodejs.org/)
- **npm 9+** (Node.js ile birlikte gelir)

### 🔧 Backend Kurulumu

1. **Proje klasörüne gidin:**
```bash
cd ExcelDataManagementAPI
```

2. **NuGet paketlerini yükleyin:**
```bash
dotnet restore
```

3. **Veritabanını oluşturun:**
```bash
dotnet ef database update
```
*Not: İlk çalıştırmada otomatik olarak oluşturulur*

4. **Uygulamayı çalıştırın:**
```bash
dotnet run
```

**Backend çalıştıktan sonra:**
- 🌐 **API**: http://localhost:5002/api
- 📖 **Swagger UI**: http://localhost:5002
- 📊 **Health Check**: http://localhost:5002/api/excel/test

### 🌐 Frontend Kurulumu

1. **NPM paketlerini yükleyin:**
```bash
npm install
```

2. **Development server'ı başlatın:**
```bash
npm run dev
```

**Frontend çalıştıktan sonra:**
- 🖥️ **React App**: http://localhost:3000
- 🔗 **API Integration**: Otomatik backend bağlantısı

### 🎯 Full Stack Test

1. ✅ Backend'i başlatın (Port 5002)
2. ✅ Frontend'i başlatın (Port 3000)
3. ✅ http://localhost:3000 adresine gidin
4. ✅ API bağlantısının "🟢 Başarılı" olduğunu kontrol edin
5. ✅ Excel dosyası yükleyip test edin

## ⚙️ Konfigürasyon

### 🔒 Backend Konfigürasyonu

**appsettings.json:**
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
  }
}
```

**appsettings.Development.json:**
```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Debug",
      "Microsoft.AspNetCore": "Information"
    }
  }
}
```

### 🌐 Frontend Konfigürasyonu

**Environment Variables (.env):**
```env
VITE_API_URL=http://localhost:5002
VITE_API_BASE_URL=http://localhost:5002/api
VITE_ENABLE_DEBUG=true
VITE_DEVELOPER_EMAIL=your-email@example.com
```

### 🔗 CORS Ayarları

Backend otomatik olarak şu portları destekler:
- **3000** - React/Vite default
- **4200** - Angular default  
- **5173** - Vite alternative
- **8080** - Vue.js default
- **127.0.0.1** - Local IP variants

## 📊 Kullanım Örnekleri

### 1. 📤 Excel Dosyası Yükleme

```typescript
// React Hook kullanımı
const { uploadFiles, loading, error, success } = useExcelApi();

const handleFileUpload = async (files: File[]) => {
  await uploadFiles(files, 'user@example.com');
};
```

### 2. 📊 Veri Görüntüleme

```typescript
// Sayfalama ile veri getirme
const { getExcelData } = useExcelApi();

const loadData = async () => {
  const response = await apiService.getExcelData(
    'filename.xlsx',  // dosya adı
    'Sheet1',         // sheet adı (opsiyonel)
    1,                // sayfa numarası
    50                // sayfa boyutu
  );
};
```

### 3. 🔄 Dosya Karşılaştırma

```typescript
// İki dosyayı karşılaştırma
const { compareFiles, comparisonResult } = useExcelApi();

await compareFiles('dosya1.xlsx', 'dosya2.xlsx', 'Sheet1');
console.log(comparisonResult);
```

### 4. ✏️ Veri Güncelleme

```typescript
// Tek satır güncelleme
const updateData = {
  id: 123,
  data: { "Ad": "Ahmet", "Soyad": "Yılmaz", "Yaş": 30 },
  modifiedBy: "user@example.com"
};

await apiService.updateData(updateData);
```

### 5. 📥 Excel Export

```typescript
// Excel dosyası indirme
const { exportExcel } = useExcelApi();

await exportExcel('filename.xlsx', 'Sheet1');
// Dosya otomatik olarak indirilir
```

## 🎮 Development Scripts

### 🔧 Backend Scripts

```bash
# Debug modunda çalıştırma
dotnet run --environment Development

# Watch mode (otomatik yeniden başlatma)
dotnet watch run

# Production build
dotnet build --configuration Release

# Database migration
dotnet ef migrations add MigrationName
dotnet ef database update

# Paket güncelleme
dotnet restore
dotnet list package --outdated
```

### 🌐 Frontend Scripts

```bash
# Development server
npm run dev              # http://localhost:3000

# Production build
npm run build           # dist/ klasörüne build eder

# Preview production build
npm run preview         # Build'i preview eder

# Type checking
npm run type-check      # TypeScript hatalarını kontrol eder

# Linting
npm run lint           # ESLint ile kod kontrolü
npm run lint:fix       # ESLint ile otomatik düzeltme

# Testing
npm run test           # Unit testleri çalıştır
npm run test:ui        # Test UI'ını aç
```

## 🐛 Troubleshooting

### 🔧 Backend Sorunları

**Port zaten kullanımda:**
```bash
# Port 5002'yi kullanan process'i bul
netstat -ano | findstr :5002
# Process'i sonlandır
taskkill /PID <PID> /F
```

**Veritabanı bağlantı hatası:**
```bash
# SQL Server durumunu kontrol et
sc query MSSQLSERVER
# LocalDB'yi başlat
sqllocaldb start mssqllocaldb
```

**NuGet paket hatası:**
```bash
# Cache'i temizle
dotnet nuget locals all --clear
# Yeniden yükle
dotnet restore --force
```

### 🌐 Frontend Sorunları

**Port çakışması:**
```bash
# Port 3000'i kullanılmıyorsa zorla sonlandır
npx kill-port 3000
# Alternatif port kullan
npm run dev -- --port 3001
```

**Node/NPM sürüm sorunu:**
```bash
# Node sürümünü kontrol et
node --version  # 18+ olmalı
npm --version   # 9+ olmalı

# NPM cache temizle
npm cache clean --force
```

**TypeScript hatalar:**
```bash
# Type definitions'ları yeniden yükle
rm -rf node_modules package-lock.json
npm install

# TypeScript compiler'ı güncelle
npm update typescript
```

### 🔗 API Bağlantı Sorunları

**CORS hatası:**
- Backend CORS policy'sini kontrol edin
- Frontend URL'ini backend'e ekleyin
- Browser cache'ini temizleyin

**404 API hatası:**
- Backend çalışıyor mu kontrol edin
- Endpoint URL'lerini kontrol edin
- Network tab'da isteği inceleyin

## 🔒 Güvenlik

### 🛡️ Backend Güvenlik
- ✅ **CORS Policy**: Cross-Origin istekleri kontrolü
- ✅ **File Validation**: Dosya türü ve boyutu kontrolü
- ✅ **SQL Injection**: Entity Framework ile korunma
- ✅ **Input Validation**: Model validation attributes
- ✅ **Error Handling**: Güvenli hata mesajları

### 🔐 Frontend Güvenlik
- ✅ **TypeScript**: Compile-time type safety
- ✅ **Environment Variables**: Güvenli konfigürasyon
- ✅ **XSS Protection**: React'ın built-in koruması
- ✅ **File Upload Validation**: Client-side validasyon

## 📈 Performans

### ⚡ Backend Optimizasyonları
- ✅ **Asenkron İşlemler**: Tüm DB işlemleri async
- ✅ **Sayfalama**: Büyük veri setleri için performans
- ✅ **Indexing**: Veritabanı indexleri
- ✅ **Connection Pooling**: EF Core connection yönetimi

### 🚀 Frontend Optimizasyonları  
- ✅ **Lazy Loading**: Component bazlı kod bölme
- ✅ **Memoization**: React.memo ve useMemo
- ✅ **Vite**: Hızlı build ve HMR
- ✅ **Bundle Optimization**: Tree shaking

## 📊 Monitoring ve Logging

### 📋 Backend Logging
```csharp
// Automatic logging levels:
_logger.LogInformation("Excel dosyası yüklendi: {FileName}", fileName);
_logger.LogWarning("Dosya bulunamadı: {FileName}", fileName);
_logger.LogError(ex, "Hata oluştu: {FileName}", fileName);
```

### 🔍 Frontend Debug
```typescript
// Development modunda otomatik debug logları
console.log('API Response:', response);
console.error('Error:', error);
```

## 🚀 Deployment

### 🏭 Production Deployment

**Backend Deployment:**
```bash
# Production build oluştur
dotnet publish -c Release -o ./publish

# IIS veya hosting servise deploy et
# Connection string'i production'a güncelle
```

**Frontend Deployment:**
```bash
# Production build
npm run build

# dist/ klasörünü web server'a deploy et
# Environment variables'ları production'a ayarla
```

### ☁️ Cloud Deployment Önerileri
- **Backend**: Azure App Service, AWS Elastic Beanstalk
- **Database**: Azure SQL Database, AWS RDS
- **Frontend**: Azure Static Web Apps, Netlify, Vercel

## 📚 Ek Kaynaklar

### 📖 Dokümantasyonlar
- [ASP.NET Core Documentation](https://docs.microsoft.com/en-us/aspnet/core/)
- [Entity Framework Core](https://docs.microsoft.com/en-us/ef/core/)
- [React Documentation](https://reactjs.org/docs/)
- [TypeScript Handbook](https://www.typescriptlang.org/docs/)
- [EPPlus Documentation](https://epplussoftware.com/docs)

### 🛠️ Development Tools
- **Visual Studio 2022** - Full IDE
- **VS Code** - Lightweight editor
- **SQL Server Management Studio** - Database management
- **Postman** - API testing
- **React Developer Tools** - Browser extension

## 🤝 Contributing

1. 🔧 Backend değişiklikleri için `ExcelDataManagementAPI/` klasöründe çalışın
2. 🌐 Frontend değişiklikleri için frontend dosyalarında çalışın
3. ⚙️ Environment variables'ları `.env.local` dosyasında override edin
4. ✅ Commit öncesi linting ve type-check yapın
5. 📝 Değişikliklerinizi dokümante edin

## 📄 License

**EPPlus Non-Commercial License** - Bu proje EPPlus kütüphanesini kullandığı için non-commercial lisans altındadır.

---

## 💫 Özellik Roadmap

### 🔮 Gelecek Sürümler
- [ ] **Real-time Collaboration**: Çoklu kullanıcı desteği
- [ ] **Advanced Filtering**: Karmaşık veri filtreleme
- [ ] **Charts & Visualization**: Grafik ve görselleştirme
- [ ] **API Authentication**: JWT token tabanlı güvenlik
- [ ] **Mobile App**: React Native mobile uygulama
- [ ] **Automated Testing**: Unit ve integration testleri
- [ ] **Docker Support**: Konteyner desteği
- [ ] **Microservices**: Mikroservis mimarisi

### 🎯 Kısa Vadeli Hedefler
- [ ] **Bulk Import**: Toplu dosya yükleme
- [ ] **Data Validation**: Gelişmiş veri validasyonu
- [ ] **Export Templates**: Export şablonları
- [ ] **User Management**: Kullanıcı yönetimi
- [ ] **Audit Dashboard**: Değişiklik raporu dashboard'u

---

> 📞 **Destek**: Herhangi bir sorun veya soru için lütfen [GitHub Issues](https://github.com/your-repo/issues) sayfasını kullanın.

> 📧 **İletişim**: Proje hakkında detaylı bilgi için proje geliştiricileri ile iletişime geçin.

**Son Güncelleme**: $(Get-Date -Format "dd/MM/yyyy HH:mm")
**Versiyon**: 1.0.0
**Durum**: ✅ Aktif Geliştirme