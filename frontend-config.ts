// API endpoint'leri için types
interface ApiEndpoints {
  // Excel Controller Endpoints
  testExcel: string;
  getFiles: string;
  uploadExcel: string;
  readExcel: string;
  getExcelData: string;
  updateData: string;
  bulkUpdateData: string;
  addRow: string;
  deleteData: string;
  exportExcel: string;
  getSheets: string;
  getStatistics: string;
  
  // Comparison Controller Endpoints
  testComparison: string;
  compareFiles: string;
  compareVersions: string;
  getChanges: string;
  getChangeHistory: string;
  getRowHistory: string;
}

// API base URL - Mevcut backend'inizin adresi
const API_BASE_URL = 'http://localhost:5002/api';

const endpoints: ApiEndpoints = {
  // Excel Controller Endpoints (/api/excel)
  testExcel: `${API_BASE_URL}/excel/test`,
  getFiles: `${API_BASE_URL}/excel/files`,
  uploadExcel: `${API_BASE_URL}/excel/upload`,
  readExcel: `${API_BASE_URL}/excel/read`, // + fileName parameter
  getExcelData: `${API_BASE_URL}/excel/data`, // + fileName parameter
  updateData: `${API_BASE_URL}/excel/data`,
  bulkUpdateData: `${API_BASE_URL}/excel/data/bulk`,
  addRow: `${API_BASE_URL}/excel/data`,
  deleteData: `${API_BASE_URL}/excel/data`, // + id parameter
  exportExcel: `${API_BASE_URL}/excel/export`,
  getSheets: `${API_BASE_URL}/excel/sheets`, // + fileName parameter
  getStatistics: `${API_BASE_URL}/excel/statistics`, // + fileName parameter
  
  // Comparison Controller Endpoints (/api/comparison)
  testComparison: `${API_BASE_URL}/comparison/test`,
  compareFiles: `${API_BASE_URL}/comparison/files`,
  compareVersions: `${API_BASE_URL}/comparison/versions`,
  getChanges: `${API_BASE_URL}/comparison/changes`, // + fileName parameter
  getChangeHistory: `${API_BASE_URL}/comparison/history`, // + fileName parameter
  getRowHistory: `${API_BASE_URL}/comparison/row-history` // + rowId parameter
};

// Export endpoints for use in components
export { endpoints, API_BASE_URL };
export type { ApiEndpoints };

// Helper functions for building URLs with parameters
export const buildUrl = {
  readExcel: (fileName: string, sheetName?: string) => {
    const url = `${endpoints.readExcel}/${fileName}`;
    return sheetName ? `${url}?sheetName=${sheetName}` : url;
  },
  
  getExcelData: (fileName: string, sheetName?: string, page?: number, pageSize?: number) => {
    const url = `${endpoints.getExcelData}/${fileName}`;
    const params = new URLSearchParams();
    if (sheetName) params.append('sheetName', sheetName);
    if (page) params.append('page', page.toString());
    if (pageSize) params.append('pageSize', pageSize.toString());
    return params.toString() ? `${url}?${params.toString()}` : url;
  },
  
  deleteData: (id: number, deletedBy?: string) => {
    const url = `${endpoints.deleteData}/${id}`;
    return deletedBy ? `${url}?deletedBy=${deletedBy}` : url;
  },
  
  getSheets: (fileName: string) => `${endpoints.getSheets}/${fileName}`,
  
  getStatistics: (fileName: string, sheetName?: string) => {
    const url = `${endpoints.getStatistics}/${fileName}`;
    return sheetName ? `${url}?sheetName=${sheetName}` : url;
  },
  
  getChanges: (fileName: string, fromDate?: Date, toDate?: Date, sheetName?: string) => {
    const url = `${endpoints.getChanges}/${fileName}`;
    const params = new URLSearchParams();
    if (fromDate) params.append('fromDate', fromDate.toISOString());
    if (toDate) params.append('toDate', toDate.toISOString());
    if (sheetName) params.append('sheetName', sheetName);
    return params.toString() ? `${url}?${params.toString()}` : url;
  },
  
  getChangeHistory: (fileName: string, sheetName?: string) => {
    const url = `${endpoints.getChangeHistory}/${fileName}`;
    return sheetName ? `${url}?sheetName=${sheetName}` : url;
  },
  
  getRowHistory: (rowId: number) => `${endpoints.getRowHistory}/${rowId}`
};