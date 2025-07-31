// services/ExcelApiService.ts - Environment config kullanan versiyon
import { env, debugLog, buildApiUrl } from '../config/environment';
import { 
  ApiResponse, 
  ExcelFile, 
  ExcelDataRow, 
  ComparisonResult,
  UploadedFileInfo 
} from '../types/excel-types';

export class ExcelApiService {
  private baseUrl: string;
  private timeout: number;
  private uploadTimeout: number;

  constructor(baseUrl?: string) {
    this.baseUrl = baseUrl || env.apiBaseUrl;
    this.timeout = env.apiTimeout;
    this.uploadTimeout = env.uploadTimeout;
    
    debugLog('ExcelApiService initialized:', {
      baseUrl: this.baseUrl,
      timeout: this.timeout,
      uploadTimeout: this.uploadTimeout
    });
  }

  // Helper method for fetch with timeout
  private async fetchWithTimeout(url: string, options: RequestInit = {}, timeoutMs?: number): Promise<Response> {
    const controller = new AbortController();
    const timeout = timeoutMs || this.timeout;
    
    const timeoutId = setTimeout(() => controller.abort(), timeout);
    
    try {
      const response = await fetch(url, {
        ...options,
        signal: controller.signal,
        headers: {
          ...options.headers,
        }
      });
      
      clearTimeout(timeoutId);
      return response;
    } catch (error) {
      clearTimeout(timeoutId);
      if (error instanceof Error && error.name === 'AbortError') {
        throw new Error(`Request timeout after ${timeout}ms`);
      }
      throw error;
    }
  }

  // API durumunu test etme
  async testConnection(): Promise<ApiResponse> {
    debugLog('Testing API connection...');
    
    const response = await this.fetchWithTimeout(buildApiUrl('excel/test'));
    if (!response.ok) {
      throw new Error(`Test failed: ${response.statusText}`);
    }
    
    const result = await response.json();
    debugLog('API test result:', result);
    return result;
  }

  // Yüklenmiþ dosya listesini alma
  async getFiles(): Promise<ApiResponse<ExcelFile[]>> {
    debugLog('Getting files list...');
    
    const response = await this.fetchWithTimeout(buildApiUrl('excel/files'));
    if (!response.ok) {
      throw new Error(`Failed to get files: ${response.statusText}`);
    }
    
    const result = await response.json();
    debugLog('Files retrieved:', result.data?.length || 0);
    return result;
  }

  // Excel dosyasý yükleme - Environment config ile validation
  async uploadExcelFile(file: File, uploadedBy?: string): Promise<ApiResponse<UploadedFileInfo>> {
    debugLog('Uploading file:', file.name, file.size);
    
    // File validation using environment config
    if (!env.allowedFileTypes.some(type => file.name.toLowerCase().endsWith(type))) {
      throw new Error(`Desteklenmeyen dosya türü. Ýzin verilen türler: ${env.allowedFileTypes.join(', ')}`);
    }
    
    if (file.size > env.maxFileSize) {
      throw new Error(`Dosya boyutu çok büyük. Maksimum boyut: ${Math.round(env.maxFileSize / 1024 / 1024)}MB`);
    }

    const formData = new FormData();
    formData.append('file', file);
    if (uploadedBy) {
      formData.append('uploadedBy', uploadedBy);
    }

    const response = await this.fetchWithTimeout(
      buildApiUrl('excel/upload'), 
      {
        method: 'POST',
        body: formData,
      },
      this.uploadTimeout // Use upload timeout for file uploads
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Upload failed: ${response.statusText} - ${errorText}`);
    }

    const result = await response.json();
    debugLog('File uploaded successfully:', result.data);
    return result;
  }

  // Excel dosyasýný okuma ve veritabanýna kaydetme
  async readExcelData(fileName: string, sheetName?: string): Promise<ApiResponse<ExcelDataRow[]>> {
    debugLog('Reading Excel data:', fileName, sheetName);
    
    const url = sheetName 
      ? buildApiUrl(`excel/read/${fileName}?sheetName=${encodeURIComponent(sheetName)}`)
      : buildApiUrl(`excel/read/${fileName}`);

    const response = await this.fetchWithTimeout(url, {
      method: 'POST',
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Read failed: ${response.statusText} - ${errorText}`);
    }

    const result = await response.json();
    debugLog('Excel data read:', result.data?.length || 0, 'rows');
    return result;
  }

  // Excel verilerini getirme (sayfalama ile)
  async getExcelData(
    fileName: string, 
    sheetName?: string, 
    page: number = 1, 
    pageSize: number = env.pageSize
  ): Promise<ApiResponse<ExcelDataRow[]>> {
    debugLog('Getting Excel data:', fileName, { page, pageSize, sheetName });
    
    const params = new URLSearchParams({
      page: page.toString(),
      pageSize: pageSize.toString()
    });
    
    if (sheetName) {
      params.append('sheetName', sheetName);
    }

    const response = await this.fetchWithTimeout(
      buildApiUrl(`excel/data/${encodeURIComponent(fileName)}?${params}`),
      { method: 'GET' }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to get data: ${response.statusText} - ${errorText}`);
    }

    const result = await response.json();
    debugLog('Excel data retrieved:', result.data?.length || 0, 'rows for page', page);
    return result;
  }

  // Ýki Excel dosyasýný karþýlaþtýrma
  async compareExcelFiles(fileName1: string, fileName2: string, sheetName?: string): Promise<ApiResponse<ComparisonResult>> {
    debugLog('Comparing files:', fileName1, 'vs', fileName2, sheetName);
    
    const response = await this.fetchWithTimeout(buildApiUrl('comparison/files'), {
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

    const result = await response.json();
    debugLog('Comparison completed:', result.data?.differences?.length || 0, 'differences found');
    return result;
  }

  // Veri güncelleme
  async updateData(id: number, data: Record<string, any>, modifiedBy?: string): Promise<ApiResponse<ExcelDataRow>> {
    debugLog('Updating data:', id, data);
    
    const response = await this.fetchWithTimeout(buildApiUrl('excel/data'), {
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

    const result = await response.json();
    debugLog('Data updated successfully:', result.data);
    return result;
  }

  // Yeni satýr ekleme
  async addRow(
    fileName: string, 
    sheetName: string, 
    rowData: Record<string, any>, 
    addedBy?: string
  ): Promise<ApiResponse<ExcelDataRow>> {
    debugLog('Adding new row:', fileName, sheetName, rowData);
    
    const response = await this.fetchWithTimeout(buildApiUrl('excel/data'), {
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

    const result = await response.json();
    debugLog('Row added successfully:', result.data);
    return result;
  }

  // Veri silme (soft delete)
  async deleteData(id: number, deletedBy?: string): Promise<ApiResponse> {
    debugLog('Deleting data:', id);
    
    const url = deletedBy 
      ? buildApiUrl(`excel/data/${id}?deletedBy=${encodeURIComponent(deletedBy)}`)
      : buildApiUrl(`excel/data/${id}`);

    const response = await this.fetchWithTimeout(url, {
      method: 'DELETE',
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Delete failed: ${response.statusText} - ${errorText}`);
    }

    const result = await response.json();
    debugLog('Data deleted successfully:', result);
    return result;
  }

  // Excel export
  async exportToExcel(
    fileName: string, 
    sheetName?: string, 
    rowIds?: number[], 
    includeModificationHistory: boolean = false
  ): Promise<Blob> {
    debugLog('Exporting Excel:', fileName, { sheetName, rowIds, includeModificationHistory });
    
    const response = await this.fetchWithTimeout(buildApiUrl('excel/export'), {
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
    }, this.uploadTimeout); // Use upload timeout for exports

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Export failed: ${response.statusText} - ${errorText}`);
    }

    const blob = await response.blob();
    debugLog('Excel export completed, blob size:', blob.size);
    return blob;
  }

  // Dosyadaki sheet listesini alma
  async getSheets(fileName: string): Promise<ApiResponse<string[]>> {
    debugLog('Getting sheets for file:', fileName);
    
    const response = await this.fetchWithTimeout(buildApiUrl(`excel/sheets/${encodeURIComponent(fileName)}`));
    
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to get sheets: ${response.statusText} - ${errorText}`);
    }

    const result = await response.json();
    debugLog('Sheets retrieved:', result.data);
    return result;
  }

  // Comparison API durumunu test etme
  async testComparisonApi(): Promise<ApiResponse> {
    debugLog('Testing Comparison API...');
    
    const response = await this.fetchWithTimeout(buildApiUrl('comparison/test'));
    if (!response.ok) {
      throw new Error(`Comparison test failed: ${response.statusText}`);
    }
    
    const result = await response.json();
    debugLog('Comparison API test result:', result);
    return result;
  }
}