import React, { useState, useEffect } from 'react';
import { useExcelApi } from '../hooks/useExcelApi';
import { ExcelApiService } from '../services/ExcelApiService';

interface ExcelComparisonOptimizedProps {
  userEmail?: string;
  baseUrl?: string;
  onError?: (error: string) => void;
  onSuccess?: (message: string) => void;
}

const ExcelComparisonOptimized: React.FC<ExcelComparisonOptimizedProps> = ({ 
  userEmail = '', 
  baseUrl,
  onError, 
  onSuccess 
}) => {
  // Local state
  const [selectedFiles, setSelectedFiles] = useState<{file1: string, file2: string}>({file1: '', file2: ''});
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [currentUserEmail, setCurrentUserEmail] = useState(userEmail);
  const [filesToUpload, setFilesToUpload] = useState<File[]>([]);
  
  // Excel API Hook
  const {
    loading,
    error,
    success,
    apiConnected,
    uploadedFiles,
    comparisonResult,
    testConnection,
    uploadFiles,
    loadFiles,
    compareFiles,
    exportExcel,
    clearMessages,
    clearComparisonResult
  } = useExcelApi(baseUrl);
  
  // API Service for additional operations
  const apiService = new ExcelApiService(baseUrl);
  
  // Effect for external callbacks
  useEffect(() => {
    if (error && onError) {
      onError(error);
    }
  }, [error, onError]);
  
  useEffect(() => {
    if (success && onSuccess) {
      onSuccess(success);
    }
  }, [success, onSuccess]);
  
  // Handle file selection
  const handleFileSelection = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files || []);
    
    // Validate Excel files
    const validFiles = selectedFiles.filter(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );
    
    if (validFiles.length !== selectedFiles.length) {
      onError?.('Sadece Excel dosyalar� (.xlsx, .xls) desteklenmektedir');
      return;
    }
    
    setFilesToUpload(validFiles);
  };
  
  // Handle file upload
  const handleUpload = async () => {
    await uploadFiles(filesToUpload, currentUserEmail || undefined);
    setFilesToUpload([]);
    
    // Clear file input
    const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
    if (fileInput) fileInput.value = '';
  };
  
  // Load sheets for selected file
  const loadSheets = async (fileName: string) => {
    try {
      const response = await apiService.getSheets(fileName);
      if (response.success && response.data) {
        setAvailableSheets(response.data);
        setSelectedSheet(''); // Reset sheet selection
      }
    } catch (error) {
      console.error('Sheet listesi al�namad�:', error);
      setAvailableSheets([]);
    }
  };
  
  // Effect to load sheets when first file is selected
  useEffect(() => {
    if (selectedFiles.file1) {
      loadSheets(selectedFiles.file1);
    } else {
      setAvailableSheets([]);
      setSelectedSheet('');
    }
  }, [selectedFiles.file1]);
  
  // Handle comparison
  const handleComparison = async () => {
    await compareFiles(selectedFiles.file1, selectedFiles.file2, selectedSheet || undefined);
  };
  
  // Handle comparison result save
  const handleSaveComparison = () => {
    if (!comparisonResult) return;
    
    try {
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
      
      onSuccess?.('Kar��la�t�rma sonucu ba�ar�yla kaydedildi');
    } catch (err) {
      onError?.('Kaydetme i�lemi ba�ar�s�z');
    }
  };

  return (
    <div className="excel-comparison-optimized" style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h2>?? Excel Dosya Kar��la�t�rma (Optimized)</h2>
      
      {/* Message Display */}
      {(error || success) && (
        <div style={{ 
          padding: '10px', 
          marginBottom: '20px', 
          borderRadius: '5px',
          backgroundColor: error ? '#f8d7da' : '#d4edda',
          color: error ? '#721c24' : '#155724',
          border: `1px solid ${error ? '#f5c6cb' : '#c3e6cb'}`,
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center'
        }}>
          <span>{error || success}</span>
          <button 
            onClick={clearMessages}
            style={{ 
              background: 'none', 
              border: 'none', 
              fontSize: '18px', 
              cursor: 'pointer',
              color: 'inherit'
            }}
          >
            �
          </button>
        </div>
      )}
      
      {/* API Connection Status */}
      <div style={{ 
        padding: '10px', 
        marginBottom: '20px', 
        borderRadius: '5px',
        backgroundColor: apiConnected === true ? '#d4edda' : apiConnected === false ? '#f8d7da' : '#d1ecf1',
        color: apiConnected === true ? '#155724' : apiConnected === false ? '#721c24' : '#0c5460',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center'
      }}>
        <span>
          {apiConnected === true && '? API ba�lant�s� ba�ar�l�'}
          {apiConnected === false && '? API ba�lant�s� ba�ar�s�z'}
          {apiConnected === null && '?? API ba�lant�s� kontrol ediliyor...'}
        </span>
        <button 
          onClick={testConnection}
          disabled={loading}
          style={{ 
            padding: '5px 10px', 
            backgroundColor: 'transparent', 
            border: '1px solid currentColor', 
            borderRadius: '3px',
            color: 'inherit',
            cursor: 'pointer'
          }}
        >
          ?? Yeniden Test Et
        </button>
      </div>

      {/* User Email */}
      <div style={{ marginBottom: '20px' }}>
        <label htmlFor="userEmail" style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
          Email Adresiniz:
        </label>
        <input
          id="userEmail"
          type="email"
          value={currentUserEmail}
          onChange={(e) => setCurrentUserEmail(e.target.value)}
          placeholder="ornek@email.com"
          style={{ 
            width: '100%', 
            maxWidth: '400px', 
            padding: '8px', 
            border: '1px solid #ddd', 
            borderRadius: '3px' 
          }}
        />
      </div>

      {/* File Upload Section */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <h3>?? Yeni Dosya Y�kleme</h3>
        
        <div style={{ marginBottom: '15px' }}>
          <input
            type="file"
            multiple
            accept=".xlsx,.xls"
            onChange={handleFileSelection}
            disabled={loading}
            style={{ marginBottom: '10px' }}
          />
        </div>
        
        {filesToUpload.length > 0 && (
          <div style={{ marginBottom: '15px' }}>
            <p><strong>Se�ilen dosyalar:</strong></p>
            <ul style={{ margin: '10px 0', paddingLeft: '20px' }}>
              {filesToUpload.map((file, index) => (
                <li key={index}>
                  {file.name} <span style={{ color: '#666' }}>({(file.size / 1024).toFixed(2)} KB)</span>
                </li>
              ))}
            </ul>
          </div>
        )}
        
        <button 
          onClick={handleUpload} 
          disabled={loading || filesToUpload.length === 0}
          style={{ 
            padding: '10px 20px', 
            backgroundColor: filesToUpload.length === 0 ? '#6c757d' : '#007bff', 
            color: 'white', 
            border: 'none', 
            borderRadius: '3px',
            cursor: loading || filesToUpload.length === 0 ? 'not-allowed' : 'pointer'
          }}
        >
          {loading ? '?? Y�kleniyor...' : `?? ${filesToUpload.length} Dosyay� Y�kle ve ��le`}
        </button>
      </div>

      {/* File Comparison Section */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
          <h3 style={{ margin: 0 }}>?? Dosya Kar��la�t�rma</h3>
          {comparisonResult && (
            <button
              onClick={clearComparisonResult}
              style={{ 
                padding: '5px 10px', 
                backgroundColor: '#dc3545', 
                color: 'white', 
                border: 'none', 
                borderRadius: '3px'
              }}
            >
              ??? Sonucu Temizle
            </button>
          )}
        </div>
        
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '15px' }}>
          <div>
            <label htmlFor="file1" style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
              1. Dosya:
            </label>
            <select
              id="file1"
              value={selectedFiles.file1}
              onChange={(e) => setSelectedFiles(prev => ({ ...prev, file1: e.target.value }))}
              disabled={loading}
              style={{ width: '100%', padding: '8px', border: '1px solid #ddd', borderRadius: '3px' }}
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
            <label htmlFor="file2" style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
              2. Dosya:
            </label>
            <select
              id="file2"
              value={selectedFiles.file2}
              onChange={(e) => setSelectedFiles(prev => ({ ...prev, file2: e.target.value }))}
              disabled={loading}
              style={{ width: '100%', padding: '8px', border: '1px solid #ddd', borderRadius: '3px' }}
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
            <label htmlFor="sheetSelect" style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
              Sheet Se�in (Opsiyonel):
            </label>
            <select
              id="sheetSelect"
              value={selectedSheet}
              onChange={(e) => setSelectedSheet(e.target.value)}
              disabled={loading}
              style={{ width: '100%', maxWidth: '300px', padding: '8px', border: '1px solid #ddd', borderRadius: '3px' }}
            >
              <option value="">T�m sheet'ler</option>
              {availableSheets.map(sheet => (
                <option key={sheet} value={sheet}>{sheet}</option>
              ))}
            </select>
          </div>
        )}
        
        <button 
          onClick={handleComparison} 
          disabled={loading || !selectedFiles.file1 || !selectedFiles.file2}
          style={{ 
            padding: '10px 20px', 
            backgroundColor: !selectedFiles.file1 || !selectedFiles.file2 ? '#6c757d' : '#28a745', 
            color: 'white', 
            border: 'none', 
            borderRadius: '3px',
            cursor: loading || !selectedFiles.file1 || !selectedFiles.file2 ? 'not-allowed' : 'pointer'
          }}
        >
          {loading ? '?? Kar��la�t�r�l�yor...' : '?? Dosyalar� Kar��la�t�r'}
        </button>
      </div>

      {/* Uploaded Files List */}
      <div style={{ 
        border: '1px solid #ddd', 
        borderRadius: '5px', 
        padding: '20px', 
        marginBottom: '20px' 
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
          <h3 style={{ margin: 0 }}>?? Y�klenmi� Dosyalar ({uploadedFiles.length} adet)</h3>
          <button 
            onClick={loadFiles}
            disabled={loading}
            style={{ 
              padding: '8px 16px', 
              backgroundColor: '#6c757d', 
              color: 'white', 
              border: 'none', 
              borderRadius: '3px'
            }}
          >
            ?? Yenile
          </button>
        </div>

        {uploadedFiles.length > 0 ? (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px' }}>
              <thead>
                <tr style={{ backgroundColor: '#f8f9fa' }}>
                  <th style={{ border: '1px solid #ddd', padding: '10px', textAlign: 'left' }}>Dosya Ad�</th>
                  <th style={{ border: '1px solid #ddd', padding: '10px', textAlign: 'left' }}>Boyut</th>
                  <th style={{ border: '1px solid #ddd', padding: '10px', textAlign: 'left' }}>Y�klenme Tarihi</th>
                  <th style={{ border: '1px solid #ddd', padding: '10px', textAlign: 'center' }}>��lemler</th>
                </tr>
              </thead>
              <tbody>
                {uploadedFiles.map(file => (
                  <tr key={file.id} style={{ ':hover': { backgroundColor: '#f8f9fa' } }}>
                    <td style={{ border: '1px solid #ddd', padding: '10px' }}>
                      <strong>{file.originalFileName}</strong>
                      <br />
                      <small style={{ color: '#666' }}>{file.fileName}</small>
                    </td>
                    <td style={{ border: '1px solid #ddd', padding: '10px' }}>
                      {(file.fileSize / 1024).toFixed(2)} KB
                    </td>
                    <td style={{ border: '1px solid #ddd', padding: '10px' }}>
                      {new Date(file.uploadDate).toLocaleString()}
                    </td>
                    <td style={{ border: '1px solid #ddd', padding: '10px', textAlign: 'center' }}>
                      <button
                        onClick={() => exportExcel(file.fileName, selectedSheet || undefined)}
                        disabled={loading}
                        style={{ 
                          padding: '5px 10px', 
                          backgroundColor: '#17a2b8', 
                          color: 'white', 
                          border: 'none', 
                          borderRadius: '3px',
                          fontSize: '12px'
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
          <p style={{ textAlign: 'center', color: '#666', fontStyle: 'italic' }}>
            Hen�z y�klenmi� dosya yok.
          </p>
        )}
      </div>

      {/* Comparison Result */}
      {comparisonResult && (
        <div style={{ 
          border: '1px solid #ddd', 
          borderRadius: '5px', 
          padding: '20px',
          backgroundColor: '#f8f9fa' 
        }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <h3 style={{ margin: 0 }}>?? Kar��la�t�rma Sonucu</h3>
            <button
              onClick={handleSaveComparison}
              style={{ 
                padding: '8px 16px', 
                backgroundColor: '#007bff', 
                color: 'white', 
                border: 'none', 
                borderRadius: '3px'
              }}
            >
              ?? JSON Olarak Kaydet
            </button>
          </div>
          
          <div style={{ marginBottom: '20px', padding: '15px', backgroundColor: 'white', borderRadius: '5px' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '15px', marginBottom: '15px' }}>
              <div>
                <strong>?? Dosya 1:</strong> {comparisonResult.file1Name}
              </div>
              <div>
                <strong>?? Dosya 2:</strong> {comparisonResult.file2Name}
              </div>
            </div>
            <div>
              <strong>?? Kar��la�t�rma Tarihi:</strong> {new Date(comparisonResult.comparisonDate).toLocaleString()}
            </div>
          </div>

          <div style={{ marginBottom: '20px' }}>
            <h4>?? �zet Bilgiler</h4>
            <div style={{ 
              display: 'grid', 
              gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', 
              gap: '15px',
              padding: '15px',
              backgroundColor: 'white',
              borderRadius: '5px'
            }}>
              <div style={{ textAlign: 'center', padding: '10px' }}>
                <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#495057' }}>
                  {comparisonResult.summary.totalRows}
                </div>
                <div style={{ fontSize: '12px', color: '#6c757d' }}>?? Toplam Sat�r</div>
              </div>
              <div style={{ textAlign: 'center', padding: '10px' }}>
                <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#ffc107' }}>
                  {comparisonResult.summary.modifiedRows}
                </div>
                <div style={{ fontSize: '12px', color: '#6c757d' }}>?? De�i�tirilen</div>
              </div>
              <div style={{ textAlign: 'center', padding: '10px' }}>
                <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#28a745' }}>
                  {comparisonResult.summary.addedRows}
                </div>
                <div style={{ fontSize: '12px', color: '#6c757d' }}>? Eklenen</div>
              </div>
              <div style={{ textAlign: 'center', padding: '10px' }}>
                <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#dc3545' }}>
                  {comparisonResult.summary.deletedRows}
                </div>
                <div style={{ fontSize: '12px', color: '#6c757d' }}>? Silinen</div>
              </div>
              <div style={{ textAlign: 'center', padding: '10px' }}>
                <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#6c757d' }}>
                  {comparisonResult.summary.unchangedRows}
                </div>
                <div style={{ fontSize: '12px', color: '#6c757d' }}>? De�i�meyen</div>
              </div>
            </div>
          </div>

          {comparisonResult.differences && comparisonResult.differences.length > 0 && (
            <div>
              <h4>?? Farkl�l�klar ({comparisonResult.differences.length} adet)</h4>
              <div style={{ 
                maxHeight: '500px', 
                overflowY: 'auto', 
                border: '1px solid #ddd', 
                borderRadius: '5px',
                backgroundColor: 'white'
              }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#e9ecef', position: 'sticky', top: 0 }}>
                      <th style={{ border: '1px solid #ddd', padding: '8px', minWidth: '60px' }}>Sat�r</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', minWidth: '100px' }}>S�tun</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', minWidth: '150px' }}>Eski De�er</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', minWidth: '150px' }}>Yeni De�er</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', minWidth: '80px' }}>Tip</th>
                    </tr>
                  </thead>
                  <tbody>
                    {comparisonResult.differences.slice(0, 200).map((diff, index) => (
                      <tr key={index}>
                        <td style={{ border: '1px solid #ddd', padding: '8px', textAlign: 'center' }}>
                          {diff.rowIndex}
                        </td>
                        <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                          {diff.columnName}
                        </td>
                        <td style={{ 
                          border: '1px solid #ddd', 
                          padding: '8px',
                          maxWidth: '200px',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis'
                        }}>
                          {diff.oldValue?.toString() || ''}
                        </td>
                        <td style={{ 
                          border: '1px solid #ddd', 
                          padding: '8px',
                          maxWidth: '200px',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis'
                        }}>
                          {diff.newValue?.toString() || ''}
                        </td>
                        <td style={{ 
                          border: '1px solid #ddd', 
                          padding: '8px',
                          textAlign: 'center',
                          fontWeight: 'bold',
                          backgroundColor: 
                            diff.type === 'Added' ? '#d4edda' :
                            diff.type === 'Deleted' ? '#f8d7da' : 
                            '#fff3cd'
                        }}>
                          {diff.type === 'Added' ? '?' :
                           diff.type === 'Deleted' ? '?' : 
                           '??'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              
              {comparisonResult.differences.length > 200 && (
                <p style={{ 
                  marginTop: '15px', 
                  padding: '10px', 
                  backgroundColor: '#fff3cd', 
                  borderRadius: '5px',
                  fontSize: '14px'
                }}>
                  ?? <strong>Not:</strong> Performans i�in ilk 200 farkl�l�k g�steriliyor. 
                  Toplam {comparisonResult.differences.length} farkl�l�k var.
                </p>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default ExcelComparisonOptimized;