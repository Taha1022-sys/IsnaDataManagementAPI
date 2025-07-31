// services/ExcelApiService.ts - React için optimize edilmiþ versiyon
import { 
  ApiResponse, 
  ExcelFile, 
  ExcelDataRow, 
  ComparisonResult,
  UploadedFileInfo 
} from '../types/excel-types';

export class ExcelApiService {
  private baseUrl: string;

  constructor(baseUrl: string = 'http://localhost:5002/api') {
    this.baseUrl = baseUrl;
  }

  // API durumunu test etme
  async testConnection(): Promise<ApiResponse> {
    const response = await fetch(`${this.baseUrl}/excel/test`);
    if (!response.ok) {
      throw new Error(`Test failed: ${response.statusText}`);
    }
    return await response.json();
  }

  // Yüklenmiþ dosya listesini alma
  async getFiles(): Promise<ApiResponse<ExcelFile[]>> {
    const response = await fetch(`${this.baseUrl}/excel/files`);
    if (!response.ok) {
      throw new Error(`Failed to get files: ${response.statusText}`);
    }
    return await response.json();
  }

  // Excel dosyasý yükleme - React için optimize edildi
  async uploadExcelFile(file: File, uploadedBy?: string): Promise<ApiResponse<UploadedFileInfo>> {
    const formData = new FormData();
    formData.append('file', file);
    if (uploadedBy) {
      formData.append('uploadedBy', uploadedBy);
    }

    const response = await fetch(`${this.baseUrl}/excel/upload`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Upload failed: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Excel dosyasýný okuma ve veritabanýna kaydetme
  async readExcelData(fileName: string, sheetName?: string): Promise<ApiResponse<ExcelDataRow[]>> {
    const url = sheetName 
      ? `${this.baseUrl}/excel/read/${fileName}?sheetName=${encodeURIComponent(sheetName)}`
      : `${this.baseUrl}/excel/read/${fileName}`;

    const response = await fetch(url, {
      method: 'POST',
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Read failed: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Excel verilerini getirme (sayfalama ile)
  async getExcelData(
    fileName: string, 
    sheetName?: string, 
    page: number = 1, 
    pageSize: number = 50
  ): Promise<ApiResponse<ExcelDataRow[]>> {
    const params = new URLSearchParams({
      page: page.toString(),
      pageSize: pageSize.toString()
    });
    
    if (sheetName) {
      params.append('sheetName', sheetName);
    }

    const response = await fetch(
      `${this.baseUrl}/excel/data/${encodeURIComponent(fileName)}?${params}`,
      { method: 'GET' }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to get data: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Ýki Excel dosyasýný karþýlaþtýrma - mevcut API'ye uygun
  async compareExcelFiles(fileName1: string, fileName2: string, sheetName?: string): Promise<ApiResponse<ComparisonResult>> {
    const response = await fetch(`${this.baseUrl}/comparison/files`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        fileName1,
        fileName2,
        sheetName
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Comparison failed: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Veri güncelleme
  async updateData(id: number, data: Record<string, any>, modifiedBy?: string): Promise<ApiResponse<ExcelDataRow>> {
    const response = await fetch(`${this.baseUrl}/excel/data`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        id,
        data,
        modifiedBy
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Update failed: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Yeni satýr ekleme
  async addRow(
    fileName: string, 
    sheetName: string, 
    rowData: Record<string, any>, 
    addedBy?: string
  ): Promise<ApiResponse<ExcelDataRow>> {
    const response = await fetch(`${this.baseUrl}/excel/data`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        fileName,
        sheetName,
        rowData,
        addedBy
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Add row failed: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Veri silme (soft delete)
  async deleteData(id: number, deletedBy?: string): Promise<ApiResponse> {
    const url = deletedBy 
      ? `${this.baseUrl}/excel/data/${id}?deletedBy=${encodeURIComponent(deletedBy)}`
      : `${this.baseUrl}/excel/data/${id}`;

    const response = await fetch(url, {
      method: 'DELETE',
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Delete failed: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Excel export
  async exportToExcel(
    fileName: string, 
    sheetName?: string, 
    rowIds?: number[], 
    includeModificationHistory: boolean = false
  ): Promise<Blob> {
    const response = await fetch(`${this.baseUrl}/excel/export`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        fileName,
        sheetName,
        rowIds,
        includeModificationHistory
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Export failed: ${response.statusText} - ${errorText}`);
    }

    return await response.blob();
  }

  // Dosyadaki sheet listesini alma
  async getSheets(fileName: string): Promise<ApiResponse<string[]>> {
    const response = await fetch(`${this.baseUrl}/excel/sheets/${encodeURIComponent(fileName)}`);
    
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to get sheets: ${response.statusText} - ${errorText}`);
    }

    return await response.json();
  }

  // Comparison API durumunu test etme
  async testComparisonApi(): Promise<ApiResponse> {
    const response = await fetch(`${this.baseUrl}/comparison/test`);
    if (!response.ok) {
      throw new Error(`Comparison test failed: ${response.statusText}`);
    }
    return await response.json();
  }
}