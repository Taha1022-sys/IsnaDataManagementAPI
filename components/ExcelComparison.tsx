import React, { useState, useEffect } from 'react';
import { ExcelApiService } from '../services/ExcelApiService';
import { 
  ExcelFile, 
  ComparisonResult, 
  ExcelDataRow,
  UploadedFileInfo 
} from '../types/excel-types';

interface ExcelComparisonProps {
  userEmail?: string;
  onError?: (error: string) => void;
  onSuccess?: (message: string) => void;
}

const ExcelComparison: React.FC<ExcelComparisonProps> = ({ 
  userEmail = '', 
  onError, 
  onSuccess 
}) => {
  // State tan�mlamalar�
  const [files, setFiles] = useState<File[]>([]);
  const [uploadedFiles, setUploadedFiles] = useState<ExcelFile[]>([]);
  const [comparisonResult, setComparisonResult] = useState<ComparisonResult | null>(null);
  const [selectedFiles, setSelectedFiles] = useState<{file1: string, file2: string}>({file1: '', file2: ''});
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [apiConnected, setApiConnected] = useState<boolean | null>(null);
  const [currentUserEmail, setCurrentUserEmail] = useState(userEmail);
  
  const apiService = new ExcelApiService();

  // Component mount edildi�inde API ba�lant�s�n� test et ve dosyalar� y�kle
  useEffect(() => {
    testApiConnection();
    loadUploadedFiles();
  }, []);

  // API ba�lant�s�n� test etme
  const testApiConnection = async () => {
    try {
      const response = await apiService.testConnection();
      setApiConnected(response.success);
      if (response.success && onSuccess) {
        onSuccess('API ba�lant�s� ba�ar�l�');
      }
    } catch (error) {
      setApiConnected(false);
      const errorMessage = error instanceof Error ? error.message : 'API ba�lant� hatas�';
      if (onError) onError(errorMessage);
      console.error('API ba�lant� hatas�:', error);
    }
  };

  // Y�klenmi� dosyalar� getirme
  const loadUploadedFiles = async () => {
    try {
      const response = await apiService.getFiles();
      if (response.success && response.data) {
        setUploadedFiles(response.data);
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Dosya listesi y�klenemedi';
      if (onError) onError(errorMessage);
      console.error('Dosya listesi hatas�:', error);
    }
  };

  // Dosya se�me i�lemi
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files || []);
    
    // Excel dosyas� kontrol�
    const validFiles = selectedFiles.filter(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );
    
    if (validFiles.length !== selectedFiles.length) {
      if (onError) onError('Sadece Excel dosyalar� (.xlsx, .xls) desteklenmektedir');
      return;
    }
    
    setFiles(validFiles);
  };

  // Dosyalar� y�kleme
  const uploadFiles = async () => {
    if (files.length === 0) {
      if (onError) onError('L�tfen en az bir dosya se�in');
      return;
    }

    setLoading(true);
    try {
      const uploadPromises = files.map(file => 
        apiService.uploadExcelFile(file, currentUserEmail || undefined)
      );
      
      const uploadResults = await Promise.all(uploadPromises);
      
      // Upload ba�ar�l� m� kontrol et
      const failedUploads = uploadResults.filter(result => !result.success);
      if (failedUploads.length > 0) {
        throw new Error(`${failedUploads.length} dosya y�klenemedi`);
      }

      // Dosyalar� otomatik olarak oku ve veritaban�na kaydet
      const readPromises = uploadResults.map(result => 
        result.data ? apiService.readExcelData(result.data.fileName) : Promise.resolve(null)
      );
      
      await Promise.all(readPromises);
      
      if (onSuccess) onSuccess(`${files.length} dosya ba�ar�yla y�klendi ve i�lendi`);
      
      // Dosya listesini yenile
      await loadUploadedFiles();
      
      // Form temizle
      setFiles([]);
      const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      if (fileInput) fileInput.value = '';
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Dosya y�kleme ba�ar�s�z';
      if (onError) onError(errorMessage);
      console.error('Dosya y�kleme hatas�:', error);
    } finally {
      setLoading(false);
    }
  };

  // Sheet listesini alma
  const loadSheets = async (fileName: string) => {
    try {
      const response = await apiService.getSheets(fileName);
      if (response.success && response.data) {
        setAvailableSheets(response.data);
      }
    } catch (error) {
      console.error('Sheet listesi al�namad�:', error);
      setAvailableSheets([]);
    }
  };

  // �lk dosya se�ildi�inde sheet listesini y�kle
  useEffect(() => {
    if (selectedFiles.file1) {
      loadSheets(selectedFiles.file1);
    }
  }, [selectedFiles.file1]);

  // Dosyalar� kar��la�t�rma
  const handleCompareFiles = async () => {
    if (!selectedFiles.file1 || !selectedFiles.file2) {
      if (onError) onError('L�tfen kar��la�t�r�lacak iki dosyay� se�in');
      return;
    }

    if (selectedFiles.file1 === selectedFiles.file2) {
      if (onError) onError('Farkl� dosyalar se�melisiniz');
      return;
    }

    setLoading(true);
    try {
      const response = await apiService.compareExcelFiles(
        selectedFiles.file1,
        selectedFiles.file2,
        selectedSheet || undefined
      );
      
      if (response.success && response.data) {
        setComparisonResult(response.data);
        if (onSuccess) onSuccess('Dosyalar ba�ar�yla kar��la�t�r�ld�');
      } else {
        throw new Error('Kar��la�t�rma ba�ar�s�z');
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Dosya kar��la�t�rma ba�ar�s�z';
      if (onError) onError(errorMessage);
      console.error('Kar��la�t�rma hatas�:', error);
    } finally {
      setLoading(false);
    }
  };

  // Kar��la�t�rma sonucunu kaydetme (�rnek olarak export)
  const handleSaveComparisonResult = async () => {
    if (!comparisonResult) {
      if (onError) onError('Kaydedilecek kar��la�t�rma sonucu yok');
      return;
    }

    try {
      // Kar��la�t�rma sonucunu JSON olarak indir
      const dataStr = JSON.stringify(comparisonResult, null, 2);
      const dataBlob = new Blob([dataStr], { type: 'application/json' });
      
      const url = URL.createObjectURL(dataBlob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `comparison_${comparisonResult.file1Name}_vs_${comparisonResult.file2Name}_${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      
      if (onSuccess) onSuccess('Kar��la�t�rma sonucu ba�ar�yla kaydedildi');
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Kaydetme i�lemi ba�ar�s�z';
      if (onError) onError(errorMessage);
      console.error('Kaydetme hatas�:', error);
    }
  };

  // Excel olarak export etme
  const handleExportExcel = async (fileName: string) => {
    try {
      setLoading(true);
      const blob = await apiService.exportToExcel(fileName, selectedSheet || undefined);
      
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${fileName}_export_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      
      if (onSuccess) onSuccess('Excel dosyas� ba�ar�yla indirildi');
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Export i�lemi ba�ar�s�z';
      if (onError) onError(errorMessage);
      console.error('Export hatas�:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="excel-comparison" style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h2>?? Excel Dosya Kar��la�t�rma</h2>
      
      {/* API Ba�lant� Durumu */}
      <div style={{ 
        padding: '10px', 
        marginBottom: '20px', 
        borderRadius: '5px',
        backgroundColor: apiConnected === true ? '#d4edda' : apiConnected === false ? '#f8d7da' : '#d1ecf1',
        color: apiConnected === true ? '#155724' : apiConnected === false ? '#721c24' : '#0c5460'
      }}>
        {apiConnected === true && '? API ba�lant�s� ba�ar�l�'}
        {apiConnected === false && '? API ba�lant�s� ba�ar�s�z'}
        {apiConnected === null && '?? API ba�lant�s� kontrol ediliyor...'}
      </div>

      {/* Kullan�c� Email */}
      <div style={{ marginBottom: '20px' }}>
        <label htmlFor="userEmail">Email Adresiniz:</label>
        <input
          id="userEmail"
          type="email"
          value={currentUserEmail}
          onChange={(e) => setCurrentUserEmail(e.target.value)}
          placeholder="ornek@email.com"
          style={{ marginLeft: '10px', padding: '8px' }}
        />
      </div>

      {/* Dosya Y�kleme B�l�m� */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <h3>?? Yeni Dosya Y�kleme</h3>
        <div style={{ marginBottom: '10px' }}>
          <input
            type="file"
            multiple
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            disabled={loading}
          />
        </div>
        
        {files.length > 0 && (
          <div style={{ marginBottom: '10px' }}>
            <p><strong>Se�ilen dosyalar:</strong></p>
            <ul>
              {files.map((file, index) => (
                <li key={index}>{file.name} ({(file.size / 1024).toFixed(2)} KB)</li>
              ))}
            </ul>
          </div>
        )}
        
        <button 
          onClick={uploadFiles} 
          disabled={loading || files.length === 0}
          style={{ 
            padding: '10px 20px', 
            backgroundColor: '#007bff', 
            color: 'white', 
            border: 'none', 
            borderRadius: '3px',
            cursor: loading ? 'not-allowed' : 'pointer'
          }}
        >
          {loading ? '?? Y�kleniyor...' : '?? Dosyalar� Y�kle ve ��le'}
        </button>
      </div>

      {/* Dosya Kar��la�t�rma B�l�m� */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <h3>?? Dosya Kar��la�t�rma</h3>
        
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '15px' }}>
          <div>
            <label htmlFor="file1">1. Dosya:</label>
            <select
              id="file1"
              value={selectedFiles.file1}
              onChange={(e) => setSelectedFiles(prev => ({ ...prev, file1: e.target.value }))}
              disabled={loading}
              style={{ width: '100%', padding: '8px', marginTop: '5px' }}
            >
              <option value="">Dosya se�in...</option>
              {uploadedFiles.map(file => (
                <option key={file.id} value={file.fileName}>
                  {file.originalFileName}
                </option>
              ))}
            </select>
          </div>
          
          <div>
            <label htmlFor="file2">2. Dosya:</label>
            <select
              id="file2"
              value={selectedFiles.file2}
              onChange={(e) => setSelectedFiles(prev => ({ ...prev, file2: e.target.value }))}
              disabled={loading}
              style={{ width: '100%', padding: '8px', marginTop: '5px' }}
            >
              <option value="">Dosya se�in...</option>
              {uploadedFiles.map(file => (
                <option key={file.id} value={file.fileName}>
                  {file.originalFileName}
                </option>
              ))}
            </select>
          </div>
        </div>

        {availableSheets.length > 0 && (
          <div style={{ marginBottom: '15px' }}>
            <label htmlFor="sheetSelect">Sheet Se�in (Opsiyonel):</label>
            <select
              id="sheetSelect"
              value={selectedSheet}
              onChange={(e) => setSelectedSheet(e.target.value)}
              disabled={loading}
              style={{ width: '100%', padding: '8px', marginTop: '5px' }}
            >
              <option value="">T�m sheet'ler</option>
              {availableSheets.map(sheet => (
                <option key={sheet} value={sheet}>{sheet}</option>
              ))}
            </select>
          </div>
        )}
        
        <button 
          onClick={handleCompareFiles} 
          disabled={loading || !selectedFiles.file1 || !selectedFiles.file2}
          style={{ 
            padding: '10px 20px', 
            backgroundColor: '#28a745', 
            color: 'white', 
            border: 'none', 
            borderRadius: '3px',
            cursor: loading || !selectedFiles.file1 || !selectedFiles.file2 ? 'not-allowed' : 'pointer'
          }}
        >
          {loading ? '?? Kar��la�t�r�l�yor...' : '?? Dosyalar� Kar��la�t�r'}
        </button>
      </div>

      {/* Y�klenmi� Dosyalar Listesi */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <h3>?? Y�klenmi� Dosyalar ({uploadedFiles.length} adet)</h3>
        
        <button 
          onClick={loadUploadedFiles}
          style={{ 
            padding: '8px 16px', 
            backgroundColor: '#6c757d', 
            color: 'white', 
            border: 'none', 
            borderRadius: '3px',
            marginBottom: '15px'
          }}
        >
          ?? Listeyi Yenile
        </button>

        {uploadedFiles.length > 0 ? (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ backgroundColor: '#f8f9fa' }}>
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>Dosya Ad�</th>
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>Boyut</th>
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>Y�klenme Tarihi</th>
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>��lemler</th>
                </tr>
              </thead>
              <tbody>
                {uploadedFiles.map(file => (
                  <tr key={file.id}>
                    <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                      {file.originalFileName}
                    </td>
                    <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                      {(file.fileSize / 1024).toFixed(2)} KB
                    </td>
                    <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                      {new Date(file.uploadDate).toLocaleString()}
                    </td>
                    <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                      <button
                        onClick={() => handleExportExcel(file.fileName)}
                        disabled={loading}
                        style={{ 
                          padding: '5px 10px', 
                          backgroundColor: '#17a2b8', 
                          color: 'white', 
                          border: 'none', 
                          borderRadius: '3px',
                          marginRight: '5px'
                        }}
                      >
                        ?? �ndir
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <p>Hen�z y�klenmi� dosya yok.</p>
        )}
      </div>

      {/* Kar��la�t�rma Sonucu */}
      {comparisonResult && (
        <div style={{ 
          border: '1px solid #ddd', 
          borderRadius: '5px', 
          padding: '20px',
          backgroundColor: '#f8f9fa' 
        }}>
          <h3>?? Kar��la�t�rma Sonucu</h3>
          
          <div style={{ marginBottom: '15px' }}>
            <p><strong>Dosya 1:</strong> {comparisonResult.file1Name}</p>
            <p><strong>Dosya 2:</strong> {comparisonResult.file2Name}</p>
            <p><strong>Kar��la�t�rma Tarihi:</strong> {new Date(comparisonResult.comparisonDate).toLocaleString()}</p>
          </div>

          <div style={{ marginBottom: '15px' }}>
            <h4>?? �zet Bilgiler</h4>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '10px' }}>
              <div>?? Toplam Sat�r: <strong>{comparisonResult.summary.totalRows}</strong></div>
              <div>?? De�i�tirilen: <strong>{comparisonResult.summary.modifiedRows}</strong></div>
              <div>? Eklenen: <strong>{comparisonResult.summary.addedRows}</strong></div>
              <div>? Silinen: <strong>{comparisonResult.summary.deletedRows}</strong></div>
              <div>? De�i�meyen: <strong>{comparisonResult.summary.unchangedRows}</strong></div>
            </div>
          </div>

          {comparisonResult.differences && comparisonResult.differences.length > 0 && (
            <div style={{ marginBottom: '15px' }}>
              <h4>?? Farkl�l�klar ({comparisonResult.differences.length} adet)</h4>
              <div style={{ maxHeight: '400px', overflowY: 'auto', border: '1px solid #ddd' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#e9ecef', position: 'sticky', top: 0 }}>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Sat�r</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>S�tun</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Eski De�er</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Yeni De�er</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Tip</th>
                    </tr>
                  </thead>
                  <tbody>
                    {comparisonResult.differences.slice(0, 100).map((diff, index) => (
                      <tr key={index}>
                        <td style={{ border: '1px solid #ddd', padding: '8px' }}>{diff.rowIndex}</td>
                        <td style={{ border: '1px solid #ddd', padding: '8px' }}>{diff.columnName}</td>
                        <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                          {diff.oldValue?.toString() || ''}
                        </td>
                        <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                          {diff.newValue?.toString() || ''}
                        </td>
                        <td style={{ 
                          border: '1px solid #ddd', 
                          padding: '8px',
                          backgroundColor: 
                            diff.type === 'Added' ? '#d4edda' :
                            diff.type === 'Deleted' ? '#f8d7da' : 
                            '#fff3cd'
                        }}>
                          {diff.type === 'Added' ? '? Eklendi' :
                           diff.type === 'Deleted' ? '? Silindi' : 
                           '?? De�i�ti'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              
              {comparisonResult.differences.length > 100 && (
                <p style={{ marginTop: '10px', fontStyle: 'italic' }}>
                  ... ve {comparisonResult.differences.length - 100} farkl�l�k daha (ilk 100 g�steriliyor)
                </p>
              )}
            </div>
          )}

          <button 
            onClick={handleSaveComparisonResult}
            style={{ 
              padding: '10px 20px', 
              backgroundColor: '#007bff', 
              color: 'white', 
              border: 'none', 
              borderRadius: '3px'
            }}
          >
            ?? Sonucu JSON Olarak Kaydet
          </button>
        </div>
      )}
    </div>
  );
};

export default ExcelComparison;