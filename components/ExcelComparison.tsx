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
  // State tanýmlamalarý
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

  // Component mount edildiðinde API baðlantýsýný test et ve dosyalarý yükle
  useEffect(() => {
    testApiConnection();
    loadUploadedFiles();
  }, []);

  // API baðlantýsýný test etme
  const testApiConnection = async () => {
    try {
      const response = await apiService.testConnection();
      setApiConnected(response.success);
      if (response.success && onSuccess) {
        onSuccess('API baðlantýsý baþarýlý');
      }
    } catch (error) {
      setApiConnected(false);
      const errorMessage = error instanceof Error ? error.message : 'API baðlantý hatasý';
      if (onError) onError(errorMessage);
      console.error('API baðlantý hatasý:', error);
    }
  };

  // Yüklenmiþ dosyalarý getirme
  const loadUploadedFiles = async () => {
    try {
      const response = await apiService.getFiles();
      if (response.success && response.data) {
        setUploadedFiles(response.data);
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Dosya listesi yüklenemedi';
      if (onError) onError(errorMessage);
      console.error('Dosya listesi hatasý:', error);
    }
  };

  // Dosya seçme iþlemi
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files || []);
    
    // Excel dosyasý kontrolü
    const validFiles = selectedFiles.filter(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );
    
    if (validFiles.length !== selectedFiles.length) {
      if (onError) onError('Sadece Excel dosyalarý (.xlsx, .xls) desteklenmektedir');
      return;
    }
    
    setFiles(validFiles);
  };

  // Dosyalarý yükleme
  const uploadFiles = async () => {
    if (files.length === 0) {
      if (onError) onError('Lütfen en az bir dosya seçin');
      return;
    }

    setLoading(true);
    try {
      const uploadPromises = files.map(file => 
        apiService.uploadExcelFile(file, currentUserEmail || undefined)
      );
      
      const uploadResults = await Promise.all(uploadPromises);
      
      // Upload baþarýlý mý kontrol et
      const failedUploads = uploadResults.filter(result => !result.success);
      if (failedUploads.length > 0) {
        throw new Error(`${failedUploads.length} dosya yüklenemedi`);
      }

      // Dosyalarý otomatik olarak oku ve veritabanýna kaydet
      const readPromises = uploadResults.map(result => 
        result.data ? apiService.readExcelData(result.data.fileName) : Promise.resolve(null)
      );
      
      await Promise.all(readPromises);
      
      if (onSuccess) onSuccess(`${files.length} dosya baþarýyla yüklendi ve iþlendi`);
      
      // Dosya listesini yenile
      await loadUploadedFiles();
      
      // Form temizle
      setFiles([]);
      const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      if (fileInput) fileInput.value = '';
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Dosya yükleme baþarýsýz';
      if (onError) onError(errorMessage);
      console.error('Dosya yükleme hatasý:', error);
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
      console.error('Sheet listesi alýnamadý:', error);
      setAvailableSheets([]);
    }
  };

  // Ýlk dosya seçildiðinde sheet listesini yükle
  useEffect(() => {
    if (selectedFiles.file1) {
      loadSheets(selectedFiles.file1);
    }
  }, [selectedFiles.file1]);

  // Dosyalarý karþýlaþtýrma
  const handleCompareFiles = async () => {
    if (!selectedFiles.file1 || !selectedFiles.file2) {
      if (onError) onError('Lütfen karþýlaþtýrýlacak iki dosyayý seçin');
      return;
    }

    if (selectedFiles.file1 === selectedFiles.file2) {
      if (onError) onError('Farklý dosyalar seçmelisiniz');
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
        if (onSuccess) onSuccess('Dosyalar baþarýyla karþýlaþtýrýldý');
      } else {
        throw new Error('Karþýlaþtýrma baþarýsýz');
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Dosya karþýlaþtýrma baþarýsýz';
      if (onError) onError(errorMessage);
      console.error('Karþýlaþtýrma hatasý:', error);
    } finally {
      setLoading(false);
    }
  };

  // Karþýlaþtýrma sonucunu kaydetme (örnek olarak export)
  const handleSaveComparisonResult = async () => {
    if (!comparisonResult) {
      if (onError) onError('Kaydedilecek karþýlaþtýrma sonucu yok');
      return;
    }

    try {
      // Karþýlaþtýrma sonucunu JSON olarak indir
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
      
      if (onSuccess) onSuccess('Karþýlaþtýrma sonucu baþarýyla kaydedildi');
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Kaydetme iþlemi baþarýsýz';
      if (onError) onError(errorMessage);
      console.error('Kaydetme hatasý:', error);
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
      
      if (onSuccess) onSuccess('Excel dosyasý baþarýyla indirildi');
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Export iþlemi baþarýsýz';
      if (onError) onError(errorMessage);
      console.error('Export hatasý:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="excel-comparison" style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h2>?? Excel Dosya Karþýlaþtýrma</h2>
      
      {/* API Baðlantý Durumu */}
      <div style={{ 
        padding: '10px', 
        marginBottom: '20px', 
        borderRadius: '5px',
        backgroundColor: apiConnected === true ? '#d4edda' : apiConnected === false ? '#f8d7da' : '#d1ecf1',
        color: apiConnected === true ? '#155724' : apiConnected === false ? '#721c24' : '#0c5460'
      }}>
        {apiConnected === true && '? API baðlantýsý baþarýlý'}
        {apiConnected === false && '? API baðlantýsý baþarýsýz'}
        {apiConnected === null && '?? API baðlantýsý kontrol ediliyor...'}
      </div>

      {/* Kullanýcý Email */}
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

      {/* Dosya Yükleme Bölümü */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <h3>?? Yeni Dosya Yükleme</h3>
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
            <p><strong>Seçilen dosyalar:</strong></p>
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
          {loading ? '?? Yükleniyor...' : '?? Dosyalarý Yükle ve Ýþle'}
        </button>
      </div>

      {/* Dosya Karþýlaþtýrma Bölümü */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <h3>?? Dosya Karþýlaþtýrma</h3>
        
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
              <option value="">Dosya seçin...</option>
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
              <option value="">Dosya seçin...</option>
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
            <label htmlFor="sheetSelect">Sheet Seçin (Opsiyonel):</label>
            <select
              id="sheetSelect"
              value={selectedSheet}
              onChange={(e) => setSelectedSheet(e.target.value)}
              disabled={loading}
              style={{ width: '100%', padding: '8px', marginTop: '5px' }}
            >
              <option value="">Tüm sheet'ler</option>
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
          {loading ? '?? Karþýlaþtýrýlýyor...' : '?? Dosyalarý Karþýlaþtýr'}
        </button>
      </div>

      {/* Yüklenmiþ Dosyalar Listesi */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <h3>?? Yüklenmiþ Dosyalar ({uploadedFiles.length} adet)</h3>
        
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
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>Dosya Adý</th>
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>Boyut</th>
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>Yüklenme Tarihi</th>
                  <th style={{ border: '1px solid #ddd', padding: '8px' }}>Ýþlemler</th>
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
                        ?? Ýndir
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <p>Henüz yüklenmiþ dosya yok.</p>
        )}
      </div>

      {/* Karþýlaþtýrma Sonucu */}
      {comparisonResult && (
        <div style={{ 
          border: '1px solid #ddd', 
          borderRadius: '5px', 
          padding: '20px',
          backgroundColor: '#f8f9fa' 
        }}>
          <h3>?? Karþýlaþtýrma Sonucu</h3>
          
          <div style={{ marginBottom: '15px' }}>
            <p><strong>Dosya 1:</strong> {comparisonResult.file1Name}</p>
            <p><strong>Dosya 2:</strong> {comparisonResult.file2Name}</p>
            <p><strong>Karþýlaþtýrma Tarihi:</strong> {new Date(comparisonResult.comparisonDate).toLocaleString()}</p>
          </div>

          <div style={{ marginBottom: '15px' }}>
            <h4>?? Özet Bilgiler</h4>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '10px' }}>
              <div>?? Toplam Satýr: <strong>{comparisonResult.summary.totalRows}</strong></div>
              <div>?? Deðiþtirilen: <strong>{comparisonResult.summary.modifiedRows}</strong></div>
              <div>? Eklenen: <strong>{comparisonResult.summary.addedRows}</strong></div>
              <div>? Silinen: <strong>{comparisonResult.summary.deletedRows}</strong></div>
              <div>? Deðiþmeyen: <strong>{comparisonResult.summary.unchangedRows}</strong></div>
            </div>
          </div>

          {comparisonResult.differences && comparisonResult.differences.length > 0 && (
            <div style={{ marginBottom: '15px' }}>
              <h4>?? Farklýlýklar ({comparisonResult.differences.length} adet)</h4>
              <div style={{ maxHeight: '400px', overflowY: 'auto', border: '1px solid #ddd' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#e9ecef', position: 'sticky', top: 0 }}>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Satýr</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Sütun</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Eski Deðer</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px' }}>Yeni Deðer</th>
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
                           '?? Deðiþti'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              
              {comparisonResult.differences.length > 100 && (
                <p style={{ marginTop: '10px', fontStyle: 'italic' }}>
                  ... ve {comparisonResult.differences.length - 100} farklýlýk daha (ilk 100 gösteriliyor)
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