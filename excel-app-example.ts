import { ExcelApiService } from './excel-api-service';

// Excel API Service'ini ba�latma
const excelService = new ExcelApiService('http://localhost:5002/api');

// �rnek kullan�m fonksiyonlar�
export class ExcelApp {
  private service: ExcelApiService;

  constructor() {
    this.service = new ExcelApiService();
  }

  // API ba�lant�s�n� test etme
  async testConnection(): Promise<boolean> {
    try {
      const response = await this.service.testApi();
      console.log('API Test Sonucu:', response);
      return response.success;
    } catch (error) {
      console.error('API ba�lant� hatas�:', error);
      return false;
    }
  }

  // Dosya y�kleme �rne�i
  async uploadFile(fileInput: HTMLInputElement, userEmail?: string): Promise<void> {
    try {
      if (!fileInput.files || fileInput.files.length === 0) {
        alert('L�tfen bir dosya se�in');
        return;
      }

      const file = fileInput.files[0];
      
      // Dosya format�n� kontrol et
      if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        alert('Sadece Excel dosyalar� (.xlsx, .xls) desteklenmektedir');
        return;
      }

      console.log('Dosya y�kleniyor...', file.name);
      
      const result = await this.service.uploadExcelFile(file, userEmail);
      
      if (result.success) {
        console.log('Dosya ba�ar�yla y�klendi:', result.data);
        alert(`Dosya y�klendi: ${result.data?.originalFileName}`);
        
        // Dosyay� otomatik olarak oku
        await this.readExcelFile(result.data?.fileName);
      }
    } catch (error) {
      console.error('Dosya y�kleme hatas�:', error);
      alert('Dosya y�kleme ba�ar�s�z');
    }
  }

  // Excel dosyas�n� okuma
  async readExcelFile(fileName: string, sheetName?: string): Promise<void> {
    try {
      console.log(`Excel dosyas� okunuyor: ${fileName}`);
      
      const result = await this.service.readExcelData(fileName, sheetName);
      
      if (result.success) {
        console.log('Excel verileri okundu:', result.data);
        alert(`${result.data?.length} sat�r veri ba�ar�yla okundu`);
        
        // Verileri g�ster
        await this.displayData(fileName, sheetName);
      }
    } catch (error) {
      console.error('Excel okuma hatas�:', error);
      alert('Excel dosyas� okunamad�');
    }
  }

  // Verileri sayfalama ile getirme ve g�r�nt�leme
  async displayData(fileName: string, sheetName?: string, page: number = 1): Promise<void> {
    try {
      const result = await this.service.getExcelData(fileName, sheetName, page, 20);
      
      if (result.success && result.data) {
        console.log(`Sayfa ${page} verileri:`, result.data);
        
        // HTML tablosu olu�tur
        this.createDataTable(result.data, `data-table-${fileName}`);
      }
    } catch (error) {
      console.error('Veri g�r�nt�leme hatas�:', error);
    }
  }

  // HTML tablosu olu�turma
  private createDataTable(data: any[], containerId: string): void {
    const container = document.getElementById(containerId) || document.body;
    
    if (data.length === 0) {
      container.innerHTML = '<p>Veri bulunamad�</p>';
      return;
    }

    // S�tun ba�l�klar�n� al
    const columns = Object.keys(data[0].data || {});
    
    let html = `
      <table border="1" style="border-collapse: collapse; width: 100%;">
        <thead>
          <tr>
            <th>ID</th>
            <th>Sat�r No</th>
            ${columns.map(col => `<th>${col}</th>`).join('')}
            <th>G�ncelleme Tarihi</th>
            <th>��lemler</th>
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
          <td>${row.modifiedDate ? new Date(row.modifiedDate).toLocaleString() : 'Hi�'}</td>
          <td>
            <button onclick="window.excelApp.editRow(${row.id})">D�zenle</button>
            <button onclick="window.excelApp.deleteRow(${row.id})">Sil</button>
          </td>
        </tr>
      `;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  }

  // Sat�r d�zenleme �rne�i
  async editRow(id: number): Promise<void> {
    try {
      // Basit prompt ile veri g�ncelleme (ger�ek uygulamada modal kullan�n)
      const newValue = prompt('Yeni de�eri girin (JSON format�nda):');
      if (!newValue) return;

      const data = JSON.parse(newValue);
      const userEmail = prompt('Email adresiniz:') || undefined;

      const result = await this.service.updateData(id, data, userEmail);
      
      if (result.success) {
        alert('Veri g�ncellendi');
        // Tabloyu yenile
        location.reload();
      }
    } catch (error) {
      console.error('G�ncelleme hatas�:', error);
      alert('G�ncelleme ba�ar�s�z');
    }
  }

  // Sat�r silme �rne�i
  async deleteRow(id: number): Promise<void> {
    try {
      if (!confirm('Bu sat�r� silmek istedi�inizden emin misiniz?')) {
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
      console.error('Silme hatas�:', error);
      alert('Silme ba�ar�s�z');
    }
  }

  // Dosya listesini alma
  async loadFileList(): Promise<void> {
    try {
      const result = await this.service.getFiles();
      
      if (result.success && result.data) {
        console.log('Y�klenmi� dosyalar:', result.data);
        this.createFileList(result.data);
      }
    } catch (error) {
      console.error('Dosya listesi hatas�:', error);
    }
  }

  // Dosya listesi HTML'i olu�turma
  private createFileList(files: any[]): void {
    const container = document.getElementById('file-list') || document.body;
    
    let html = '<h3>Y�klenmi� Dosyalar</h3><ul>';
    
    files.forEach(file => {
      html += `
        <li>
          <strong>${file.originalFileName}</strong> 
          (${(file.fileSize / 1024).toFixed(2)} KB)
          <br>
          Y�klenme: ${new Date(file.uploadDate).toLocaleString()}
          <br>
          <button onclick="window.excelApp.readExcelFile('${file.fileName}')">Oku</button>
          <button onclick="window.excelApp.displayData('${file.fileName}')">G�r�nt�le</button>
        </li>
      `;
    });
    
    html += '</ul>';
    container.innerHTML = html;
  }

  // �ki dosyay� kar��la�t�rma
  async compareFiles(fileName1: string, fileName2: string, sheetName?: string): Promise<void> {
    try {
      console.log(`Dosyalar kar��la�t�r�l�yor: ${fileName1} vs ${fileName2}`);
      
      const result = await this.service.compareExcelFiles(fileName1, fileName2, sheetName);
      
      if (result.success && result.data) {
        console.log('Kar��la�t�rma sonucu:', result.data);
        this.displayComparisonResult(result.data);
      }
    } catch (error) {
      console.error('Kar��la�t�rma hatas�:', error);
      alert('Kar��la�t�rma ba�ar�s�z');
    }
  }

  // Kar��la�t�rma sonucunu g�r�nt�leme
  private displayComparisonResult(result: any): void {
    const container = document.getElementById('comparison-result') || document.body;
    
    let html = `
      <h3>Kar��la�t�rma Sonucu</h3>
      <p><strong>Dosya 1:</strong> ${result.file1Name}</p>
      <p><strong>Dosya 2:</strong> ${result.file2Name}</p>
      <p><strong>Kar��la�t�rma Tarihi:</strong> ${new Date(result.comparisonDate).toLocaleString()}</p>
      
      <h4>�zet</h4>
      <ul>
        <li>Toplam Sat�r: ${result.summary.totalRows}</li>
        <li>De�i�tirilen: ${result.summary.modifiedRows}</li>
        <li>Eklenen: ${result.summary.addedRows}</li>
        <li>Silinen: ${result.summary.deletedRows}</li>
        <li>De�i�meyen: ${result.summary.unchangedRows}</li>
      </ul>
    `;

    if (result.differences && result.differences.length > 0) {
      html += '<h4>Farkl�l�klar</h4><table border="1"><tr><th>Sat�r</th><th>S�tun</th><th>Eski De�er</th><th>Yeni De�er</th><th>Tip</th></tr>';
      
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

  // Excel export �rne�i
  async exportData(fileName: string, sheetName?: string, includeHistory: boolean = false): Promise<void> {
    try {
      console.log('Excel export ediliyor...');
      
      const blob = await this.service.exportToExcel(fileName, sheetName, undefined, includeHistory);
      
      // Dosyay� indirme
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
      console.error('Export hatas�:', error);
      alert('Export ba�ar�s�z');
    }
  }
}

// Global eri�im i�in
declare global {
  interface Window {
    excelApp: ExcelApp;
  }
}

// Sayfa y�klendi�inde ba�lat
document.addEventListener('DOMContentLoaded', () => {
  window.excelApp = new ExcelApp();
  
  // Ba�lant�y� test et
  window.excelApp.testConnection().then(connected => {
    if (connected) {
      console.log('? API ba�lant�s� ba�ar�l�');
      // Dosya listesini y�kle
      window.excelApp.loadFileList();
    } else {
      console.error('? API ba�lant�s� ba�ar�s�z');
    }
  });
});

export default ExcelApp;