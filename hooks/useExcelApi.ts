// hooks/useExcelApi.ts - React Hook for Excel API operations
import { useState, useEffect, useCallback } from 'react';
import { ExcelApiService } from '../services/ExcelApiService';
import { ExcelFile, ComparisonResult } from '../types/excel-types';

interface UseExcelApiReturn {
  // State
  loading: boolean;
  error: string | null;
  success: string | null;
  apiConnected: boolean | null;
  uploadedFiles: ExcelFile[];
  comparisonResult: ComparisonResult | null;
  
  // Actions
  testConnection: () => Promise<void>;
  uploadFiles: (files: File[], userEmail?: string) => Promise<void>;
  loadFiles: () => Promise<void>;
  compareFiles: (fileName1: string, fileName2: string, sheetName?: string) => Promise<void>;
  exportExcel: (fileName: string, sheetName?: string) => Promise<void>;
  clearMessages: () => void;
  clearComparisonResult: () => void;
}

export const useExcelApi = (baseUrl?: string): UseExcelApiReturn => {
  // State
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [apiConnected, setApiConnected] = useState<boolean | null>(null);
  const [uploadedFiles, setUploadedFiles] = useState<ExcelFile[]>([]);
  const [comparisonResult, setComparisonResult] = useState<ComparisonResult | null>(null);
  
  // API Service instance
  const apiService = new ExcelApiService(baseUrl);
  
  // Helper function for error handling
  const handleError = useCallback((err: unknown) => {
    const errorMessage = err instanceof Error ? err.message : 'Bilinmeyen hata oluþtu';
    setError(errorMessage);
    setSuccess(null);
    console.error('API Error:', err);
  }, []);
  
  // Helper function for success messages
  const handleSuccess = useCallback((message: string) => {
    setSuccess(message);
    setError(null);
  }, []);
  
  // Clear messages
  const clearMessages = useCallback(() => {
    setError(null);
    setSuccess(null);
  }, []);
  
  // Clear comparison result
  const clearComparisonResult = useCallback(() => {
    setComparisonResult(null);
  }, []);
  
  // Test API connection
  const testConnection = useCallback(async () => {
    try {
      const response = await apiService.testConnection();
      setApiConnected(response.success);
      if (response.success) {
        handleSuccess('API baðlantýsý baþarýlý');
      }
    } catch (err) {
      setApiConnected(false);
      handleError(err);
    }
  }, [apiService, handleError, handleSuccess]);
  
  // Load uploaded files
  const loadFiles = useCallback(async () => {
    try {
      setLoading(true);
      const response = await apiService.getFiles();
      if (response.success && response.data) {
        setUploadedFiles(response.data);
      }
    } catch (err) {
      handleError(err);
    } finally {
      setLoading(false);
    }
  }, [apiService, handleError]);
  
  // Upload multiple files
  const uploadFiles = useCallback(async (files: File[], userEmail?: string) => {
    if (files.length === 0) {
      setError('Lütfen en az bir dosya seçin');
      return;
    }
    
    // Validate file types
    const validFiles = files.filter(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );
    
    if (validFiles.length !== files.length) {
      setError('Sadece Excel dosyalarý (.xlsx, .xls) desteklenmektedir');
      return;
    }
    
    try {
      setLoading(true);
      clearMessages();
      
      // Upload files
      const uploadPromises = validFiles.map(file => 
        apiService.uploadExcelFile(file, userEmail)
      );
      
      const uploadResults = await Promise.all(uploadPromises);
      
      // Check for failed uploads
      const failedUploads = uploadResults.filter(result => !result.success);
      if (failedUploads.length > 0) {
        throw new Error(`${failedUploads.length} dosya yüklenemedi`);
      }
      
      // Read files automatically
      const readPromises = uploadResults.map(result => 
        result.data ? apiService.readExcelData(result.data.fileName) : Promise.resolve(null)
      );
      
      await Promise.all(readPromises);
      
      handleSuccess(`${validFiles.length} dosya baþarýyla yüklendi ve iþlendi`);
      
      // Refresh file list
      await loadFiles();
      
    } catch (err) {
      handleError(err);
    } finally {
      setLoading(false);
    }
  }, [apiService, handleError, handleSuccess, loadFiles]);
  
  // Compare two files
  const compareFiles = useCallback(async (fileName1: string, fileName2: string, sheetName?: string) => {
    if (!fileName1 || !fileName2) {
      setError('Lütfen karþýlaþtýrýlacak iki dosyayý seçin');
      return;
    }
    
    if (fileName1 === fileName2) {
      setError('Farklý dosyalar seçmelisiniz');
      return;
    }
    
    try {
      setLoading(true);
      clearMessages();
      
      const response = await apiService.compareExcelFiles(fileName1, fileName2, sheetName);
      
      if (response.success && response.data) {
        setComparisonResult(response.data);
        handleSuccess('Dosyalar baþarýyla karþýlaþtýrýldý');
      } else {
        throw new Error('Karþýlaþtýrma baþarýsýz');
      }
    } catch (err) {
      handleError(err);
    } finally {
      setLoading(false);
    }
  }, [apiService, handleError, handleSuccess]);
  
  // Export Excel file
  const exportExcel = useCallback(async (fileName: string, sheetName?: string) => {
    try {
      setLoading(true);
      clearMessages();
      
      const blob = await apiService.exportToExcel(fileName, sheetName);
      
      // Download file
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${fileName}_export_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      
      handleSuccess('Excel dosyasý baþarýyla indirildi');
    } catch (err) {
      handleError(err);
    } finally {
      setLoading(false);
    }
  }, [apiService, handleError, handleSuccess]);
  
  // Initialize on mount
  useEffect(() => {
    testConnection();
    loadFiles();
  }, [testConnection, loadFiles]);
  
  return {
    // State
    loading,
    error,
    success,
    apiConnected,
    uploadedFiles,
    comparisonResult,
    
    // Actions
    testConnection,
    uploadFiles,
    loadFiles,
    compareFiles,
    exportExcel,
    clearMessages,
    clearComparisonResult
  };
};