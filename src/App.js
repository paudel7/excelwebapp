import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import PivotTableUI from 'react-pivottable/PivotTableUI';
import 'react-pivottable/pivottable.css';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import styles from './App.css';

const ExcelAnalysisApp = () => {
  const [file, setFile] = useState(null);
  const [columns, setColumns] = useState([]);
  const [dataSummary, setDataSummary] = useState(null);
  const [dataPreview, setDataPreview] = useState([]);
  const [selectedOperation, setSelectedOperation] = useState('');
  const [showFullPreview, setShowFullPreview] = useState(false);
  const [selectedColumns, setSelectedColumns] = useState([]);
  const [uniqueList, setUniqueList] = useState([]);
  const [fullData, setFullData] = useState([]);
  const [pivotConfig, setPivotConfig] = useState({
    rows: [],
    cols: [],
    vals: [],
    aggregatorName: "Count",
    rendererName: "Table"
  });
  const [pivotData, setPivotData] = useState([]);
  const [chartData, setChartData] = useState([]);
  const [activeTab, setActiveTab] = useState('data');

  const handleFileUpload = (event) => {
    const uploadedFile = event.target.files[0];
    setFile(uploadedFile);
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      const headers = jsonData[0];
      const rows = jsonData.slice(1);

      const columnsInfo = headers.map(header => ({
        name: header,
        type: inferColumnType(rows.map(row => row[headers.indexOf(header)]))
      }));

      const processedData = rows.map(row => 
        Object.fromEntries(headers.map((header, i) => [header, row[i]]))
      );

      setColumns(columnsInfo);
      setFullData(processedData);
      setDataPreview(processedData.slice(0, 5));

      setDataSummary({
        rowCount: rows.length,
        columnCount: headers.length,
        numericColumns: columnsInfo.filter(col => col.type === 'number').map(col => col.name),
        categoricalColumns: columnsInfo.filter(col => col.type === 'string').map(col => col.name),
        dateColumns: columnsInfo.filter(col => col.type === 'date').map(col => col.name)
      });
    };
    reader.readAsArrayBuffer(uploadedFile);
  };

  const inferColumnType = (values) => {
    const nonNullValues = values.filter(v => v != null);
    if (nonNullValues.every(v => !isNaN(v))) return 'number';
    if (nonNullValues.every(v => !isNaN(Date.parse(v)))) return 'date';
    return 'string';
  };

  const handleOperationSelect = (operation) => {
    setSelectedOperation(operation);
  };

  const handleColumnSelect = (columnName) => {
    setSelectedColumns(prev => 
      prev.includes(columnName) 
        ? prev.filter(col => col !== columnName)
        : [...prev, columnName]
    );
  };

  const createUniqueList = () => {
    if (selectedColumns.length === 0) return;

    const uniqueValues = new Set();
    fullData.forEach(row => {
      const combinedValue = selectedColumns.map(col => row[col]).join(' - ');
      uniqueValues.add(combinedValue);
    });

    setUniqueList(Array.from(uniqueValues));
  };
  
  const handlePivotConfigChange = (config) => {
    setPivotConfig(config);
    updatePivotAndChartData(config);
  };
  
  const updatePivotAndChartData = (config) => {
    if (fullData.length === 0) return;
  
    // Create pivot data
    const pivotResult = fullData.reduce((acc, row) => {
      const rowKey = config.rows.map(r => row[r]).join('-');
      const colKey = config.cols.map(c => row[c]).join('-');
      const value = config.vals.length > 0 ? Number(row[config.vals[0]]) || 0 : 1;
  
      if (!acc[rowKey]) acc[rowKey] = {};
      if (!acc[rowKey][colKey]) acc[rowKey][colKey] = 0;
      acc[rowKey][colKey] += value;
  
      return acc;
    }, {});
  
    setPivotData(pivotResult);
  
    // Create chart data
    const chartData = Object.entries(pivotResult).map(([rowKey, colValues]) => ({
      name: rowKey,
      ...colValues
    }));
    setChartData(chartData);
  };
  

  useEffect(() => {
    if (fullData.length > 0 && (pivotConfig.rows.length > 0 || pivotConfig.cols.length > 0)) {
      updatePivotAndChartData(pivotConfig);
    }
  }, [fullData, pivotConfig]);
  
  const operations = [
    { name: 'Create Pivot Table', value: 'pivot' },
    { name: 'Create Chart', value: 'chart' },
    { name: 'Generate VBA Code', value: 'vba' },
    { name: 'Create Macro', value: 'macro' },
    { name: 'Summarize Data', value: 'summary' },
    { name: 'Map Data', value: 'map' }
  ];

  return (
    <div style={{ padding: '1rem' }}>
      <h1 style={{ fontSize: '1.5rem', fontWeight: 'bold', marginBottom: '1rem' }}>Excel Analysis Web App</h1>
      
      <div style={{ marginBottom: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
        <h2>File Upload</h2>
        <input type="file" onChange={handleFileUpload} accept=".xlsx, .xls" style={{ marginBottom: '0.5rem' }} />
        {file && <p>File uploaded: {file.name}</p>}
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1rem' }}>
        <div style={{ border: '1px solid #ccc', padding: '1rem' }}>
          <h2>Column Information</h2>
          {columns.length > 0 ? (
            <table style={{ width: '100%' }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left' }}>Name</th>
                  <th style={{ textAlign: 'left' }}>Type</th>
                </tr>
              </thead>
              <tbody>
                {columns.map((col, index) => (
                  <tr key={index}>
                    <td>{col.name}</td>
                    <td>{col.type}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <p>Upload a file to see column information</p>
          )}
        </div>

        <div style={{ border: '1px solid #ccc', padding: '1rem' }}>
          <h2>Data Summary</h2>
          {dataSummary ? (
            <ul>
              <li>Total Rows: {dataSummary.rowCount}</li>
              <li>Total Columns: {dataSummary.columnCount}</li>
              <li>Numeric Columns: {dataSummary.numericColumns.join(', ')}</li>
              <li>Categorical Columns: {dataSummary.categoricalColumns.join(', ')}</li>
              <li>Date Columns: {dataSummary.dateColumns.join(', ')}</li>
            </ul>
          ) : (
            <p>Upload a file to see data summary</p>
          )}
        </div>
      </div>

      <div style={{ marginBottom: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
        <h2 style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span>Data Preview</span>
          {dataPreview.length > 0 && (
            <button onClick={() => setShowFullPreview(!showFullPreview)}>
              {showFullPreview ? 'Show Less' : 'Show More'}
            </button>
          )}
        </h2>
        {dataPreview.length > 0 ? (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%' }}>
              <thead>
                <tr>
                  {columns.map((col) => (
                    <th key={col.name} style={{ textAlign: 'left', padding: '0.5rem', border: '1px solid #ccc' }}>{col.name}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(showFullPreview ? fullData : dataPreview).map((row, index) => (
                  <tr key={index}>
                    {columns.map((col) => (
                      <td key={col.name} style={{ padding: '0.5rem', border: '1px solid #ccc' }}>{row[col.name]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <p>Upload a file to see data preview</p>
        )}
      </div>

      <div style={{ marginBottom: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
        <h2>Create Unique List</h2>
        <div style={{ marginBottom: '1rem' }}>
          {columns.map((col) => (
            <label key={col.name} style={{ display: 'flex', alignItems: 'center', marginBottom: '0.5rem' }}>
              <input 
                type="checkbox"
                checked={selectedColumns.includes(col.name)}
                onChange={() => handleColumnSelect(col.name)}
              />
              <span style={{ marginLeft: '0.5rem' }}>{col.name}</span>
            </label>
          ))}
        </div>
        <button onClick={createUniqueList} disabled={selectedColumns.length === 0}>
          Create Unique List
        </button>
        {uniqueList.length > 0 && (
          <div style={{ marginTop: '1rem' }}>
            <h3 style={{ fontWeight: 'bold', marginBottom: '0.5rem' }}>Unique Values:</h3>
            <ul>
              {uniqueList.map((value, index) => (
                <li key={index}>{value}</li>
              ))}
            </ul>
          </div>
        )}
      </div>
  
          


  
    <div style={{ padding: '1rem' }}>
      <h1 style={{ fontSize: '1.5rem', fontWeight: 'bold', marginBottom: '1rem' }}>Excel Analysis Web App</h1>
      
      {/* ... (keep all existing sections: File Upload, Column Information, Data Summary, Data Preview, Create Unique List, Select Operation) */}

      <div style={{ marginBottom: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
        <h2>Pivot Table and Chart Configuration</h2>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '1rem' }}>
          <div>
            <h3>Rows</h3>
            <select 
              multiple 
              value={pivotConfig.rows} 
              onChange={(e) => setPivotConfig({...pivotConfig, rows: Array.from(e.target.selectedOptions, option => option.value)})}
              style={{ width: '100%', height: '100px' }}
            >
              {columns.map(col => <option key={col.name} value={col.name}>{col.name}</option>)}
            </select>
          </div>
          <div>
            <h3>Columns</h3>
            <select 
              multiple 
              value={pivotConfig.cols} 
              onChange={(e) => setPivotConfig({...pivotConfig, cols: Array.from(e.target.selectedOptions, option => option.value)})}
              style={{ width: '100%', height: '100px' }}
            >
              {columns.map(col => <option key={col.name} value={col.name}>{col.name}</option>)}
            </select>
          </div>
          <div>
            <h3>Values</h3>
            <select 
              value={pivotConfig.vals[0] || ''} 
              onChange={(e) => setPivotConfig({...pivotConfig, vals: [e.target.value]})}
              style={{ width: '100%' }}
            >
              <option value="">Select a value</option>
              {columns.filter(col => col.type === 'number').map(col => <option key={col.name} value={col.name}>{col.name}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div style={{ marginBottom: '1rem' }}>
        <button onClick={() => setActiveTab('data')} style={{ marginRight: '0.5rem' }}>Data</button>
        <button onClick={() => setActiveTab('pivot')} style={{ marginRight: '0.5rem' }}>Pivot Table</button>
        <button onClick={() => setActiveTab('chart')}>Chart</button>
      </div>

      {activeTab === 'data' && (
        <div style={{ marginBottom: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
          <h2>Data Preview</h2>
          {/* ... (keep the existing data preview table) */}
        </div>
      )}

      {activeTab === 'pivot' && (
        <div style={{ marginBottom: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
          <h2>Pivot Table</h2>
          {fullData.length > 0 ? (
            <PivotTableUI
              data={fullData}
              onChange={handlePivotConfigChange}
              {...pivotConfig}
            />
          ) : (
            <p>Upload a file and configure the pivot table to see results</p>
          )}
        </div>
      )}

      {activeTab === 'chart' && (
        <div style={{ marginBottom: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
          <h2>Chart</h2>
          {chartData.length > 0 ? (
            <ResponsiveContainer width="100%" height={400}>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" />
                <YAxis />
                <Tooltip />
                <Legend />
                {Object.keys(chartData[0]).filter(key => key !== 'name').map((key, index) => (
                  <Bar key={key} dataKey={key} fill={`#${Math.floor(Math.random()*16777215).toString(16)}`} />
                ))}
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <p>Configure the pivot table to see the chart</p>
          )}
        </div>
      )}

  <div style={{ border: '1px solid #ccc', padding: '1rem' }}>
    <h2>Select Operation</h2>
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '0.5rem' }}>
      {operations.map((op) => (
        <button
          key={op.value}
          onClick={() => handleOperationSelect(op.value)}
          style={{
            padding: '0.5rem',
            backgroundColor: selectedOperation === op.value ? '#3b82f6' : '#e5e7eb',
            color: selectedOperation === op.value ? 'white' : 'black',
            border: 'none',
            borderRadius: '0.25rem',
            cursor: 'pointer'
          }}
          disabled={!file}
        >
          {op.name}
        </button>
      ))}
    </div>
  </div>
      
  {selectedOperation && (
        <div style={{ marginTop: '1rem', border: '1px solid #ccc', padding: '1rem' }}>
          <h2>Operation Result</h2>
          <p>Result for {selectedOperation} operation would be displayed here</p>
        </div>
      )}
  </div>
</div>
  );

};

export default ExcelAnalysisApp;