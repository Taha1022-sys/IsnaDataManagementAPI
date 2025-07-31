// Types for API responses and requests
interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  message?: string;
  page?: number;
  pageSize?: number;
  count?: number;
}

interface ExcelFile {
  id: number;
  fileName: string;
  originalFileName: string;
  filePath: string;
  fileSize: number;
  uploadDate: string;
  uploadedBy?: string;
  isActive: boolean;
}

interface ExcelDataRow {
  id: number;
  fileName: string;
  sheetName: string;
  rowIndex: number;
  data: Record<string, any>;
  createdDate: string;
  modifiedDate?: string;
  version: number;
  modifiedBy?: string;
}

interface ComparisonResult {
  comparisonId: string;
  file1Name: string;
  file2Name: string;
  comparisonDate: string;
  differences: DataDifference[];
  summary: ComparisonSummary;
}

interface DataDifference {
  rowIndex: number;
  columnName: string;
  oldValue: any;
  newValue: any;
  type: 'Modified' | 'Added' | 'Deleted';
}

interface ComparisonSummary {
  totalRows: number;
  modifiedRows: number;
  addedRows: number;
  deletedRows: number;
  unchangedRows: number;
}

// Main Excel API Service Class
export class ExcelApiService {
  private baseUrl: string;

  constructor(baseUrl: string = 'http://localhost:5002/api') {
    this.baseUrl = baseUrl;
  }

  // API durumunu test etme
  async testApi(): Promise<ApiResponse> {
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

  // Excel dosyasý yükleme
  async uploadExcelFile(file: File, uploadedBy?: string): Promise<ApiResponse<ExcelFile>> {
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
      throw new Error(`Upload failed: ${response.statusText}`);
    }

    return await response.json();
  }

  // Excel dosyasýný okuma ve veritabanýna kaydetme
  async readExcelData(fileName: string, sheetName?: string): Promise<ApiResponse<ExcelDataRow[]>> {
    const url = sheetName 
      ? `${this.baseUrl}/excel/read/${fileName}?sheetName=${sheetName}`
      : `${this.baseUrl}/excel/read/${fileName}`;

    const response = await fetch(url, {
      method: 'POST',
    });

    if (!response.ok) {
      throw new Error(`Read failed: ${response.statusText}`);
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

    const response = await fetch(`${this.baseUrl}/excel/data/${fileName}?${params}`, {
      method: 'GET',
    });

    if (!response.ok) {
      throw new Error(`Failed to get data: ${response.statusText}`);
    }

    return await response.json();
  }

  // Tekil veri güncelleme
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
      throw new Error(`Update failed: ${response.statusText}`);
    }

    return await response.json();
  }

  // Toplu veri güncelleme
  async bulkUpdateData(updates: Array<{id: number, data: Record<string, any>}>, modifiedBy?: string): Promise<ApiResponse<ExcelDataRow[]>> {
    const response = await fetch(`${this.baseUrl}/excel/data/bulk`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        updates: updates.map(update => ({
          id: update.id,
          data: update.data,
          modifiedBy
        })),
        modifiedBy
      }),
    });

    if (!response.ok) {
      throw new Error(`Bulk update failed: ${response.statusText}`);
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
      throw new Error(`Add row failed: ${response.statusText}`);
    }

    return await response.json();
  }

  // Veri silme (soft delete)
  async deleteData(id: number, deletedBy?: string): Promise<ApiResponse> {
    const url = deletedBy 
      ? `${this.baseUrl}/excel/data/${id}?deletedBy=${deletedBy}`
      : `${this.baseUrl}/excel/data/${id}`;

    const response = await fetch(url, {
      method: 'DELETE',
    });

    if (!response.ok) {
      throw new Error(`Delete failed: ${response.statusText}`);
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
      throw new Error(`Export failed: ${response.statusText}`);
    }

    return await response.blob();
  }

  // Dosyadaki sheet listesini alma
  async getSheets(fileName: string): Promise<ApiResponse<string[]>> {
    const response = await fetch(`${this.baseUrl}/excel/sheets/${fileName}`);
    
    if (!response.ok) {
      throw new Error(`Failed to get sheets: ${response.statusText}`);
    }

    return await response.json();
  }

  // Dosya istatistikleri alma
  async getStatistics(fileName: string, sheetName?: string): Promise<ApiResponse> {
    const url = sheetName 
      ? `${this.baseUrl}/excel/statistics/${fileName}?sheetName=${sheetName}`
      : `${this.baseUrl}/excel/statistics/${fileName}`;

    const response = await fetch(url);
    
    if (!response.ok) {
      throw new Error(`Failed to get statistics: ${response.statusText}`);
    }

    return await response.json();
  }

  // Ýki Excel dosyasýný karþýlaþtýrma
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
      throw new Error(`Comparison failed: ${response.statusText}`);
    }

    return await response.json();
  }

  // Ayný dosyanýn farklý versiyonlarýný karþýlaþtýrma
  async compareVersions(
    fileName: string, 
    version1Date: Date, 
    version2Date: Date, 
    sheetName?: string
  ): Promise<ApiResponse<ComparisonResult>> {
    const response = await fetch(`${this.baseUrl}/comparison/versions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        fileName,
        version1Date: version1Date.toISOString(),
        version2Date: version2Date.toISOString(),
        sheetName
      }),
    });

    if (!response.ok) {
      throw new Error(`Version comparison failed: ${response.statusText}`);
    }

    return await response.json();
  }

  // Tarih aralýðýndaki deðiþiklikleri alma
  async getChanges(
    fileName: string, 
    fromDate?: Date, 
    toDate?: Date, 
    sheetName?: string
  ): Promise<ApiResponse> {
    const params = new URLSearchParams();
    if (fromDate) params.append('fromDate', fromDate.toISOString());
    if (toDate) params.append('toDate', toDate.toISOString());
    if (sheetName) params.append('sheetName', sheetName);

    const url = params.toString() 
      ? `${this.baseUrl}/comparison/changes/${fileName}?${params}`
      : `${this.baseUrl}/comparison/changes/${fileName}`;

    const response = await fetch(url);
    
    if (!response.ok) {
      throw new Error(`Failed to get changes: ${response.statusText}`);
    }

    return await response.json();
  }

  // Dosya deðiþiklik geçmiþini alma
  async getChangeHistory(fileName: string, sheetName?: string): Promise<ApiResponse> {
    const url = sheetName 
      ? `${this.baseUrl}/comparison/history/${fileName}?sheetName=${sheetName}`
      : `${this.baseUrl}/comparison/history/${fileName}`;

    const response = await fetch(url);
    
    if (!response.ok) {
      throw new Error(`Failed to get change history: ${response.statusText}`);
    }

    return await response.json();
  }

  // Belirli satýrýn geçmiþini alma
  async getRowHistory(rowId: number): Promise<ApiResponse> {
    const response = await fetch(`${this.baseUrl}/comparison/row-history/${rowId}`);
    
    if (!response.ok) {
      throw new Error(`Failed to get row history: ${response.statusText}`);
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

// Export types for use in other files
export type {
  ApiResponse,
  ExcelFile,
  ExcelDataRow,
  ComparisonResult,
  DataDifference,
  ComparisonSummary
};