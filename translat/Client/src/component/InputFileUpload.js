import * as React from 'react';
import { useState, useEffect } from 'react';
import Button from '@mui/material/Button';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import TranslateIcon from '@mui/icons-material/Translate';
import DownloadIcon from '@mui/icons-material/Download';
import RefreshIcon from '@mui/icons-material/Refresh';
import CircularProgress from '@mui/material/CircularProgress';
import Alert from '@mui/material/Alert';
import Table from '@mui/material/Table';
import TableBody from '@mui/material/TableBody';
import TableCell from '@mui/material/TableCell';
import TableContainer from '@mui/material/TableContainer';
import TableHead from '@mui/material/TableHead';
import TableRow from '@mui/material/TableRow';
import Paper from '@mui/material/Paper';
import Typography from '@mui/material/Typography';
import '../componentCss/InputFileUploadCss.css';

export default function InputFileUpload() {
  const [uploading, setUploading] = useState(false);
  const [translating, setTranslating] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [uploadedFilename, setUploadedFilename] = useState('');
  const [excelFile, setExcelFile] = useState('');
  const [downloadUrl, setDownloadUrl] = useState('');
  const [existingFiles, setExistingFiles] = useState([]);
  const [loadingFiles, setLoadingFiles] = useState(false);

  // טעינת קבצים קיימים בעת טעינת הקומפוננטה
  useEffect(() => {
    loadExistingFiles();
  }, []);

  const loadExistingFiles = async () => {
    setLoadingFiles(true);
    try {
      const response = await fetch('http://localhost:5000/files');
      if (response.ok) {
        const data = await response.json();
        setExistingFiles(data.files);
      }
    } catch (err) {
      console.error('Error loading files:', err);
    } finally {
      setLoadingFiles(false);
    }
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    // בדיקה שהקובץ הוא PDF
    if (file.type !== 'application/pdf') {
      setError('אנא בחר קובץ PDF בלבד');
      return;
    }

    setUploading(true);
    setError('');
    setMessage('');
    setExcelFile('');
    setDownloadUrl('');

    const formData = new FormData();
    formData.append('pdf', file);

    try {
      const response = await fetch('http://localhost:5000/upload', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('שגיאה בהעלאת הקובץ');
      }

      const result = await response.json();
      setMessage(`קובץ ${result.filename} הועלה בהצלחה!`);
      setUploadedFilename(result.filename);
    } catch (err) {
      setError(err.message);
    } finally {
      setUploading(false);
    }
  };

  const handleTranslate = async () => {
    if (!uploadedFilename) return;

    setTranslating(true);
    setError('');
    setMessage('');

    try {
      const response = await fetch('http://localhost:5000/translate', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ filename: uploadedFilename }),
      });

      if (!response.ok) {
        throw new Error('שגיאה בתרגום הקובץ');
      }

      const result = await response.json();
      setMessage(`${result.message} - קובץ Excel נוצר בהצלחה!`);
      setExcelFile(result.excel_file);
      setDownloadUrl(result.download_url);
      
      // רענון רשימת הקבצים
      loadExistingFiles();
    } catch (err) {
      setError(err.message);
    } finally {
      setTranslating(false);
    }
  };

  const handleDownload = (downloadUrl) => {
    window.open(`http://localhost:5000${downloadUrl}`, '_blank');
  };

  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div>
      <Button
        component="label"
        role={undefined}
        variant="contained"
        tabIndex={-1}
        startIcon={uploading ? <CircularProgress size={20} color="inherit" /> : <CloudUploadIcon />}
        disabled={uploading}
      >
        {uploading ? 'מעלה קובץ...' : 'העלה קובץ PDF'}
        <input
          type="file"
          className="visually-hidden-input"
          onChange={handleFileUpload}
          accept=".pdf"
        />
      </Button>
      
      {uploadedFilename && !excelFile && (
        <Button
          variant="contained"
          color="secondary"
          startIcon={translating ? <CircularProgress size={20} color="inherit" /> : <TranslateIcon />}
          disabled={translating}
          onClick={handleTranslate}
          style={{ marginTop: '10px', marginLeft: '10px' }}
        >
          {translating ? 'מתרגם...' : 'תרגם קובץ'}
        </Button>
      )}

      {excelFile && (
        <Button
          variant="contained"
          color="success"
          startIcon={<DownloadIcon />}
          onClick={() => handleDownload(downloadUrl)}
          style={{ marginTop: '10px', marginLeft: '10px' }}
        >
          הורד קובץ Excel
        </Button>
      )}

      <Button
        variant="outlined"
        startIcon={loadingFiles ? <CircularProgress size={20} /> : <RefreshIcon />}
        onClick={loadExistingFiles}
        disabled={loadingFiles}
        style={{ marginTop: '10px', marginLeft: '10px' }}
      >
        רענן רשימה
      </Button>
      
      {message && (
        <Alert severity="success" style={{ marginTop: '10px' }}>
          {message}
        </Alert>
      )}
      
      {error && (
        <Alert severity="error" style={{ marginTop: '10px' }}>
          {error}
        </Alert>
      )}

      {/* רשימת קבצי Excel קיימים */}
      <div style={{ marginTop: '20px' }}>
        <Typography variant="h6" gutterBottom>
          קבצי Excel קיימים ({existingFiles.length})
        </Typography>
        
        {existingFiles.length > 0 ? (
          <TableContainer component={Paper}>
            <Table>
              <TableHead>
                <TableRow>
                  <TableCell>שם הקובץ</TableCell>
                  <TableCell>תאריך יצירה</TableCell>
                  <TableCell>גודל</TableCell>
                  <TableCell>פעולות</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {existingFiles.map((file, index) => (
                  <TableRow key={index}>
                    <TableCell>{file.name}</TableCell>
                    <TableCell>{file.created}</TableCell>
                    <TableCell>{formatFileSize(file.size)}</TableCell>
                    <TableCell>
                      <Button
                        variant="contained"
                        size="small"
                        startIcon={<DownloadIcon />}
                        onClick={() => handleDownload(file.download_url)}
                      >
                        הורד
                      </Button>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
        ) : (
          <Typography variant="body2" color="textSecondary">
            אין קבצי Excel זמינים
          </Typography>
        )}
      </div>
    </div>
  );
}
