// config/environment.ts - Environment variables'larý type-safe kullanmak için
interface EnvironmentConfig {
  // API Settings
  apiUrl: string;
  apiBaseUrl: string;
  apiTimeout: number;
  uploadTimeout: number;
  
  // Development Settings
  nodeEnv: string;
  port: number;
  enableDebug: boolean;
  enableMockData: boolean;
  
  // File Upload Settings
  maxFileSize: number;
  allowedFileTypes: string[];
  
  // UI Settings
  pageSize: number;
  maxComparisonResults: number;
  
  // Optional Developer Settings
  developerEmail?: string;
  testUserEmail?: string;
}

const getEnvironmentConfig = (): EnvironmentConfig => {
  // Vite environment variables (VITE_ prefix ile baþlayanlar)
  const config: EnvironmentConfig = {
    // API Settings
    apiUrl: import.meta.env.VITE_API_URL || 'http://localhost:5002',
    apiBaseUrl: import.meta.env.VITE_API_BASE_URL || 'http://localhost:5002/api',
    apiTimeout: parseInt(import.meta.env.VITE_API_TIMEOUT || '30000'),
    uploadTimeout: parseInt(import.meta.env.VITE_UPLOAD_TIMEOUT || '60000'),
    
    // Development Settings
    nodeEnv: import.meta.env.VITE_NODE_ENV || 'development',
    port: parseInt(import.meta.env.VITE_PORT || '3000'),
    enableDebug: import.meta.env.VITE_ENABLE_DEBUG === 'true',
    enableMockData: import.meta.env.VITE_ENABLE_MOCK_DATA === 'true',
    
    // File Upload Settings
    maxFileSize: parseInt(import.meta.env.VITE_MAX_FILE_SIZE || '10485760'), // 10MB
    allowedFileTypes: (import.meta.env.VITE_ALLOWED_FILE_TYPES || '.xlsx,.xls').split(','),
    
    // UI Settings
    pageSize: parseInt(import.meta.env.VITE_PAGE_SIZE || '50'),
    maxComparisonResults: parseInt(import.meta.env.VITE_MAX_COMPARISON_RESULTS || '200'),
    
    // Optional Settings
    developerEmail: import.meta.env.VITE_DEVELOPER_EMAIL,
    testUserEmail: import.meta.env.VITE_TEST_USER_EMAIL,
  };
  
  // Validation
  if (!config.apiUrl || !config.apiBaseUrl) {
    throw new Error('API URL configuration is missing');
  }
  
  return config;
};

// Global config instance
export const env = getEnvironmentConfig();

// Development helpers
export const isDevelopment = env.nodeEnv === 'development';
export const isProduction = env.nodeEnv === 'production';

// Debug logger
export const debugLog = (...args: any[]) => {
  if (env.enableDebug) {
    console.log('[DEBUG]', ...args);
  }
};

// API URL builders
export const buildApiUrl = (endpoint: string) => {
  const cleanEndpoint = endpoint.startsWith('/') ? endpoint.slice(1) : endpoint;
  return `${env.apiBaseUrl}/${cleanEndpoint}`;
};

// File validation helpers
export const isValidFileType = (fileName: string): boolean => {
  const extension = '.' + fileName.split('.').pop()?.toLowerCase();
  return env.allowedFileTypes.includes(extension);
};

export const isValidFileSize = (fileSize: number): boolean => {
  return fileSize <= env.maxFileSize;
};

// Format file size for display
export const formatFileSize = (bytes: number): string => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

export default env;