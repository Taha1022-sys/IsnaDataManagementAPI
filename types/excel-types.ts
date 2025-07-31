// types/excel-types.ts
export interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  message?: string;
  page?: number;
  pageSize?: number;
  count?: number;
}

export interface ExcelFile {
  id: number;
  fileName: string;
  originalFileName: string;
  filePath: string;
  fileSize: number;
  uploadDate: string;
  uploadedBy?: string;
  isActive: boolean;
}

export interface ExcelDataRow {
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

export interface ComparisonResult {
  comparisonId: string;
  file1Name: string;
  file2Name: string;
  comparisonDate: string;
  differences: DataDifference[];
  summary: ComparisonSummary;
}

export interface DataDifference {
  rowIndex: number;
  columnName: string;
  oldValue: any;
  newValue: any;
  type: 'Modified' | 'Added' | 'Deleted';
}

export interface ComparisonSummary {
  totalRows: number;
  modifiedRows: number;
  addedRows: number;
  deletedRows: number;
  unchangedRows: number;
}

export interface UploadedFileInfo {
  fileName: string;
  originalFileName: string;
  fileSize: number;
}