import { ExcelApiService } from './excel-api-service';

// Excel API Service'ini baþlatma
const excelService = new ExcelApiService('http://localhost:5002/api');

// Örnek kullaným fonksiyonlarý
export class ExcelApp {
  private service: ExcelApiService;

  constructor() {
    this.service = new ExcelApiService();
  }

  // API baðlantýsýný test etme
  async testConnection(): Promise<boolean> {
    try {
      const response = await this.service.testApi();
      console.log('API Test Sonucu:', response);
      return response.success;
    } catch (error) {
      console.error('API baðlantý hatasý:', error);
      return false;
    }
  }

  // Dosya yükleme örneði
  async uploadFile(fileInput: HTMLInputElement, userEmail?: string): Promise<void> {
    try {
      if (!fileInput.files || fileInput.files.length === 0) {
        alert('Lütfen bir dosya seçin');
        return;
      }

      const file = fileInput.files[0];
      
      // Dosya formatýný kontrol et
      if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        alert('Sadece Excel dosyalarý (.xlsx, .xls) desteklenmektedir');
        return;
      }

      console.log('Dosya yükleniyor...', file.name);
      
      const result = await this.service.uploadExcelFile(file, userEmail);
      
      if (result.success) {
        console.log('Dosya baþarýyla yüklendi:', result.data);
        alert(`Dosya yüklendi: ${result.data?.originalFileName}`);
        
        // Dosyayý otomatik olarak oku
        await this.readExcelFile(result.data?.fileName);
      }
    } catch (error) {
      console.error('Dosya yükleme hatasý:', error);
      alert('Dosya yükleme baþarýsýz');
    }
  }

  // Excel dosyasýný okuma
  async readExcelFile(fileName: string, sheetName?: string): Promise<void> {
    try {
      console.log(`Excel dosyasý okunuyor: ${fileName}`);
      
      const result = await this.service.readExcelData(fileName, sheetName);
      
      if (result.success) {
        console.log('Excel verileri okundu:', result.data);
        alert(`${result.data?.length} satýr veri baþarýyla okundu`);
        
        // Verileri göster
        await this.displayData(fileName, sheetName);
      }
    } catch (error) {
      console.error('Excel okuma hatasý:', error);
      alert('Excel dosyasý okunamadý');
    }
  }

  // Verileri sayfalama ile getirme ve görüntüleme
  async displayData(fileName: string, sheetName?: string, page: number = 1): Promise<void> {
    try {
      const result = await this.service.getExcelData(fileName, sheetName, page, 20);
      
      if (result.success && result.data) {
        console.log(`Sayfa ${page} verileri:`, result.data);
        
        // HTML tablosu oluþtur
        this.createDataTable(result.data, `data-table-${fileName}`);
      }
    } catch (error) {
      console.error('Veri görüntüleme hatasý:', error);
    }
  }

  // HTML tablosu oluþturma
  private createDataTable(data: any[], containerId: string): void {
    const container = document.getElementById(containerId) || document.body;
    
    if (data.length === 0) {
      container.innerHTML = '<p>Veri bulunamadý</p>';
      return;
    }

    // Sütun baþlýklarýný al
    const columns = Object.keys(data[0].data || {});
    
    let html = `
      <table border="1" style="border-collapse: collapse; width: 100%;">
        <thead>
          <tr>
            <th>ID</th>
            <th>Satýr No</th>
            ${columns.map(col => `<th>${col}</th>`).join('')}
            <th>Güncelleme Tarihi</th>
            <th>Ýþlemler</th>
          </tr>
        </thead>
        <tbody>
    `;

    data.forEach(row => {
      html += `
        <tr>
          <td>${row.id}</td>
          <td>${row.rowIndex}</td>
          ${columns.map(col => `<td>${row.data[col] || ''}</td>`).join('')}
          <td>${row.modifiedDate ? new Date(row.modifiedDate).toLocaleString() : 'Hiç'}</td>
          <td>
            <button onclick="window.excelApp.editRow(${row.id})">Düzenle</button>
            <button onclick="window.excelApp.deleteRow(${row.id})">Sil</button>
          </td>
        </tr>
      `;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  }

  // Satýr düzenleme örneði
  async editRow(id: number): Promise<void> {
    try {
      // Basit prompt ile veri güncelleme (gerçek uygulamada modal kullanýn)
      const newValue = prompt('Yeni deðeri girin (JSON formatýnda):');
      if (!newValue) return;

      const data = JSON.parse(newValue);
      const userEmail = prompt('Email adresiniz:') || undefined;

      const result = await this.service.updateData(id, data, userEmail);
      
      if (result.success) {
        alert('Veri güncellendi');
        // Tabloyu yenile
        location.reload();
      }
    } catch (error) {
      console.error('Güncelleme hatasý:', error);
      alert('Güncelleme baþarýsýz');
    }
  }

  // Satýr silme örneði
  async deleteRow(id: number): Promise<void> {
    try {
      if (!confirm('Bu satýrý silmek istediðinizden emin misiniz?')) {
        return;
      }

      const userEmail = prompt('Email adresiniz:') || undefined;
      const result = await this.service.deleteData(id, userEmail);
      
      if (result.success) {
        alert('Veri silindi');
        // Tabloyu yenile
        location.reload();
      }
    } catch (error) {
      console.error('Silme hatasý:', error);
      alert('Silme baþarýsýz');
    }
  }

  // Dosya listesini alma
  async loadFileList(): Promise<void> {
    try {
      const result = await this.service.getFiles();
      
      if (result.success && result.data) {
        console.log('Yüklenmiþ dosyalar:', result.data);
        this.createFileList(result.data);
      }
    } catch (error) {
      console.error('Dosya listesi hatasý:', error);
    }
  }

  // Dosya listesi HTML'i oluþturma
  private createFileList(files: any[]): void {
    const container = document.getElementById('file-list') || document.body;
    
    let html = '<h3>Yüklenmiþ Dosyalar</h3><ul>';
    
    files.forEach(file => {
      html += `
        <li>
          <strong>${file.originalFileName}</strong> 
          (${(file.fileSize / 1024).toFixed(2)} KB)
          <br>
          Yüklenme: ${new Date(file.uploadDate).toLocaleString()}
          <br>
          <button onclick="window.excelApp.readExcelFile('${file.fileName}')">Oku</button>
          <button onclick="window.excelApp.displayData('${file.fileName}')">Görüntüle</button>
        </li>
      `;
    });
    
    html += '</ul>';
    container.innerHTML = html;
  }

  // Ýki dosyayý karþýlaþtýrma
  async compareFiles(fileName1: string, fileName2: string, sheetName?: string): Promise<void> {
    try {
      console.log(`Dosyalar karþýlaþtýrýlýyor: ${fileName1} vs ${fileName2}`);
      
      const result = await this.service.compareExcelFiles(fileName1, fileName2, sheetName);
      
      if (result.success && result.data) {
        console.log('Karþýlaþtýrma sonucu:', result.data);
        this.displayComparisonResult(result.data);
      }
    } catch (error) {
      console.error('Karþýlaþtýrma hatasý:', error);
      alert('Karþýlaþtýrma baþarýsýz');
    }
  }

  // Karþýlaþtýrma sonucunu görüntüleme
  private displayComparisonResult(result: any): void {
    const container = document.getElementById('comparison-result') || document.body;
    
    let html = `
      <h3>Karþýlaþtýrma Sonucu</h3>
      <p><strong>Dosya 1:</strong> ${result.file1Name}</p>
      <p><strong>Dosya 2:</strong> ${result.file2Name}</p>
      <p><strong>Karþýlaþtýrma Tarihi:</strong> ${new Date(result.comparisonDate).toLocaleString()}</p>
      
      <h4>Özet</h4>
      <ul>
        <li>Toplam Satýr: ${result.summary.totalRows}</li>
        <li>Deðiþtirilen: ${result.summary.modifiedRows}</li>
        <li>Eklenen: ${result.summary.addedRows}</li>
        <li>Silinen: ${result.summary.deletedRows}</li>
        <li>Deðiþmeyen: ${result.summary.unchangedRows}</li>
      </ul>
    `;

    if (result.differences && result.differences.length > 0) {
      html += '<h4>Farklýlýklar</h4><table border="1"><tr><th>Satýr</th><th>Sütun</th><th>Eski Deðer</th><th>Yeni Deðer</th><th>Tip</th></tr>';
      
      result.differences.forEach((diff: any) => {
        html += `
          <tr>
            <td>${diff.rowIndex}</td>
            <td>${diff.columnName}</td>
            <td>${diff.oldValue}</td>
            <td>${diff.newValue}</td>
            <td>${diff.type}</td>
          </tr>
        `;
      });
      
      html += '</table>';
    }

    container.innerHTML = html;
  }

  // Excel export örneði
  async exportData(fileName: string, sheetName?: string, includeHistory: boolean = false): Promise<void> {
    try {
      console.log('Excel export ediliyor...');
      
      const blob = await this.service.exportToExcel(fileName, sheetName, undefined, includeHistory);
      
      // Dosyayý indirme
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${fileName}_export_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
      
      alert('Dosya indirildi');
    } catch (error) {
      console.error('Export hatasý:', error);
      alert('Export baþarýsýz');
    }
  }
}

// Global eriþim için
declare global {
  interface Window {
    excelApp: ExcelApp;
  }
}

// Sayfa yüklendiðinde baþlat
document.addEventListener('DOMContentLoaded', () => {
  window.excelApp = new ExcelApp();
  
  // Baðlantýyý test et
  window.excelApp.testConnection().then(connected => {
    if (connected) {
      console.log('? API baðlantýsý baþarýlý');
      // Dosya listesini yükle
      window.excelApp.loadFileList();
    } else {
      console.error('? API baðlantýsý baþarýsýz');
    }
  });
});

export default ExcelApp;