import React, { useState, useCallback, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import _ from 'lodash';

const EmployeeDataVisualization = () => {
  const [file, setFile] = useState(null);
  const [fileStructure, setFileStructure] = useState(null);
  const [data, setData] = useState([]);
  const [stats, setStats] = useState({});
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState('upload');
  const [error, setError] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const [showSampleData, setShowSampleData] = useState(false);

  // Handle file drop
  const handleDrag = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  }, []);

  // Process the uploaded file
  const processFile = useCallback(async (file) => {
    try {
      setLoading(true);
      setError(null);
      setFile(file);
      
      // Read file as ArrayBuffer
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, {
            cellStyles: true,
            cellFormulas: true,
            cellDates: true,
            cellNF: true,
            sheetStubs: true,
            type: 'array'
          });
          
          // Analyze file structure
          const firstSheetName = workbook.SheetNames[0];
          const firstSheet = workbook.Sheets[firstSheetName];
          const range = XLSX.utils.decode_range(firstSheet['!ref'] || 'A1');
          
          // Extract headers
          const headers = [];
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell = firstSheet[XLSX.utils.encode_cell({r: 0, c: C})];
            headers.push(cell ? cell.v : null);
          }
          
          // Get column types from sample data
          const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
          const columnTypes = {};
          headers.forEach((header, idx) => {
            if (header) {
              const types = new Set();
              jsonData.slice(1, 10).forEach(row => {
                if (row[idx] !== undefined) {
                  types.add(typeof row[idx]);
                }
              });
              columnTypes[header] = Array.from(types).join('/');
            }
          });
          
          // Save file structure information
          setFileStructure({
            fileName: file.name,
            fileSize: (file.size / 1024).toFixed(2) + ' KB',
            sheetNames: workbook.SheetNames,
            activeSheet: firstSheetName,
            rowCount: range.e.r + 1,
            columnCount: range.e.c + 1,
            headers,
            columnTypes,
            sampleRows: jsonData.slice(1, Math.min(5, jsonData.length))
          });
          
          // Process data for visualization
          const rows = jsonData.slice(1);
          
          // Find column indices
          const leidinggevendeIndex = headers.findIndex(h => h === 'Leidinggevende');
          const parttimeIndex = headers.findIndex(h => h === 'Parttime (%)');
          const aanwezigIndex = headers.findIndex(h => 
            h === 'Aanwezig' || h === 'Present' || h === 'Participation'
          );
          
          if (leidinggevendeIndex === -1 || parttimeIndex === -1) {
            setError('Required columns missing. Please check that your file has "Leidinggevende" and "Parttime (%)" columns.');
            setLoading(false);
            return;
          }
          
          // Transform data - Important fix for participation status
          const formattedData = rows.map(row => {
            // Extract the "Aanwezig" value (or equivalent)
            let presentValue = aanwezigIndex !== -1 ? row[aanwezigIndex] : undefined;
            
            // Ensure we have a consistent representation
            if (presentValue === undefined) {
              presentValue = 'nee'; // Default to "no" if the field doesn't exist
            }
            
            return {
              id: row[0],
              name: row[1],
              function: row[2],
              employmentType: row[3],
              employeeType: row[4],
              startDate: row[5] ? new Date(row[5]) : null,
              endDate: row[6] ? new Date(row[6]) : null,
              employer: row[9],
              manager: row[leidinggevendeIndex],
              partTimePercentage: row[parttimeIndex],
              present: presentValue
            };
          });
          
          // Calculate statistics per manager
          const managerStats = {};
          formattedData.forEach(employee => {
            if (employee.manager) {
              if (!managerStats[employee.manager]) {
                managerStats[employee.manager] = {
                  totalEmployees: 0,
                  presentEmployees: 0,
                  absentEmployees: 0,
                  totalPartTimePercentage: 0,
                  avgPartTimePercentage: 0
                };
              }
              
              managerStats[employee.manager].totalEmployees++;
              
              // Check if present (ja/yes) or not - Improved logic
              const presentValue = employee.present;
              let isPresent = false;
              
              // Handle different formats of 'present' value
              if (presentValue !== undefined && presentValue !== null) {
                if (typeof presentValue === 'boolean') {
                  isPresent = presentValue;
                } else if (typeof presentValue === 'number') {
                  isPresent = presentValue !== 0;
                } else if (typeof presentValue === 'string') {
                  const normalizedValue = presentValue.toLowerCase().trim();
                  isPresent = ['ja', 'yes', 'y', 'true', '1'].includes(normalizedValue);
                }
              }
              
              if (isPresent) {
                managerStats[employee.manager].presentEmployees++;
              } else {
                managerStats[employee.manager].absentEmployees++;
              }
              
              if (employee.partTimePercentage) {
                managerStats[employee.manager].totalPartTimePercentage += 
                  typeof employee.partTimePercentage === 'string' ? 
                  parseFloat(employee.partTimePercentage) : 
                  employee.partTimePercentage;
              }
            }
          });
          
          // Calculate averages
          Object.keys(managerStats).forEach(manager => {
            managerStats[manager].avgPartTimePercentage = 
              (managerStats[manager].totalPartTimePercentage / managerStats[manager].totalEmployees).toFixed(2);
          });
          
          setData(formattedData);
          setStats(managerStats);
          setActiveTab('structure');
          setLoading(false);
        } catch (err) {
          console.error('Error processing file:', err);
          setError('Failed to process the Excel file. Please check the format.');
          setLoading(false);
        }
      };
      
      reader.onerror = () => {
        setError('Error reading file.');
        setLoading(false);
      };
      
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error('Error handling file:', err);
      setError('Failed to process the file. Please try again.');
      setLoading(false);
    }
  }, []);

  // Handle file drop event
  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0]);
    }
  }, [processFile]);

  // Handle file input change
  const handleChange = useCallback((e) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0]);
    }
  }, [processFile]);

  // Function to create downloadable Excel file with highlighting
  const downloadExcelWithHighlighting = () => {
    if (!data.length) return;
    
    // Create a new workbook
    const wb = XLSX.utils.book_new();
    
    // Convert data to worksheet format
    const wsData = [
      ['Name', 'Function', 'PO', 'Part-time %', 'Present']
    ];
    
    data.forEach(employee => {
      wsData.push([
        employee.name,
        employee.function,
        employee.manager,
        employee.partTimePercentage,
        formatPresentDisplay(employee.present)
      ]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, "Employees");
    
    // Add a statistics sheet
    const statsData = [
      ['Product Owner', 'Total Employees', 'Present', 'Absent', 'Avg. Part-time %']
    ];
    
    Object.entries(stats).forEach(([manager, stat]) => {
      statsData.push([
        manager,
        stat.totalEmployees,
        stat.presentEmployees,
        stat.absentEmployees,
        stat.avgPartTimePercentage
      ]);
    });
    
    const statsWs = XLSX.utils.aoa_to_sheet(statsData);
    XLSX.utils.book_append_sheet(wb, statsWs, "Statistics");
    
    // Generate Excel file and trigger download
    XLSX.writeFile(wb, "employee_analysis.xlsx");
  };

  // Determine if a person is not participating - Improved logic
  const isPresentClassName = (present) => {
    // Explicitly check for various "non-present" values
    if (present === undefined || 
        present === null || 
        present === false || 
        present === 0 || 
        present === "" ||
        (typeof present === 'string' && 
          (['nee', 'no', 'n', 'false', '0'].includes(present.toLowerCase()) || 
           !['ja', 'yes', 'y', 'true', '1'].includes(present.toLowerCase()))
        )) {
      return 'bg-red-200';
    }
    return '';
  };

  // Utility function to format the present field display
  const formatPresentDisplay = (present) => {
    if (typeof present === 'string') {
      const normalized = present.toLowerCase().trim();
      if (['ja', 'yes', 'y', 'true', '1'].includes(normalized)) {
        return 'Yes';
      } else if (['nee', 'no', 'n', 'false', '0'].includes(normalized)) {
        return 'No';
      }
    }
    return present === true || present === 1 ? 'Yes' : 'No';
  };

  // Count non-participating employees
  const nonParticipatingCount = useMemo(() => {
    return data.filter(employee => isPresentClassName(employee.present) !== '').length;
  }, [data]);

  const renderUploadTab = () => (
    <div className="flex flex-col items-center justify-center border-2 border-dashed border-gray-300 rounded p-8 h-64"
         onDragEnter={handleDrag}
         onDragLeave={handleDrag}
         onDragOver={handleDrag}
         onDrop={handleDrop}
         style={{ backgroundColor: dragActive ? '#f0f9ff' : 'white' }}>
      <div className="mb-4 text-gray-500">
        <svg className="w-12 h-12 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
        </svg>
        <p className="text-center">Drag & drop your Excel file here or</p>
      </div>
      <label className="px-4 py-2 bg-blue-500 text-white rounded cursor-pointer hover:bg-blue-600">
        Browse Files
        <input type="file" onChange={handleChange} accept=".xlsx,.xls,.xlsb,.xlsm" className="hidden" />
      </label>
      <p className="mt-4 text-sm text-gray-500 text-center">
        Supports .xlsx, .xls, .xlsb, and .xlsm files
      </p>
    </div>
  );

  const renderFileStructure = () => (
    <div className="bg-white rounded shadow p-4 mb-4">
      <div className="flex justify-between items-center mb-4">
        <h2 className="text-lg font-semibold">File Structure</h2>
        <button 
          className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
          onClick={() => setActiveTab('employees')}
        >
          Continue to Analysis
        </button>
      </div>
      
      <div className="grid grid-cols-2 gap-4 mb-4">
        <div>
          <p><span className="font-medium">File name:</span> {fileStructure.fileName}</p>
          <p><span className="font-medium">File size:</span> {fileStructure.fileSize}</p>
          <p><span className="font-medium">Sheets:</span> {fileStructure.sheetNames.join(', ')}</p>
        </div>
        <div>
          <p><span className="font-medium">Active sheet:</span> {fileStructure.activeSheet}</p>
          <p><span className="font-medium">Rows:</span> {fileStructure.rowCount}</p>
          <p><span className="font-medium">Columns:</span> {fileStructure.columnCount}</p>
        </div>
      </div>
      
      <h3 className="font-medium mb-2">Column Headers and Types</h3>
      <div className="overflow-x-auto mb-4">
        <table className="min-w-full border">
          <thead>
            <tr className="bg-gray-100">
              <th className="py-2 px-3 border text-left">Column Name</th>
              <th className="py-2 px-3 border text-left">Data Type</th>
            </tr>
          </thead>
          <tbody>
            {fileStructure.headers.map((header, index) => (
              header && (
                <tr key={index} className={index % 2 === 0 ? 'bg-gray-50' : ''}>
                  <td className="py-2 px-3 border">{header}</td>
                  <td className="py-2 px-3 border">{fileStructure.columnTypes[header] || 'unknown'}</td>
                </tr>
              )
            ))}
          </tbody>
        </table>
      </div>
      
      <div className="border rounded mb-4">
        <button 
          className="w-full flex justify-between items-center p-3 bg-gray-100 hover:bg-gray-200"
          onClick={() => setShowSampleData(!showSampleData)}
        >
          <h3 className="font-medium">Sample Data (First 5 Rows)</h3>
          <svg
            className={`w-5 h-5 transition-transform ${showSampleData ? 'transform rotate-180' : ''}`}
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
            xmlns="http://www.w3.org/2000/svg"
          >
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7"></path>
          </svg>
        </button>
        
        {showSampleData && (
          <div className="overflow-x-auto p-3">
            <table className="min-w-full border">
              <thead>
                <tr className="bg-gray-100">
                  {fileStructure.headers.map((header, index) => (
                    header && <th key={index} className="py-2 px-3 border text-left">{header}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {fileStructure.sampleRows.map((row, rowIndex) => (
                  <tr key={rowIndex} className={rowIndex % 2 === 0 ? 'bg-gray-50' : ''}>
                    {fileStructure.headers.map((header, colIndex) => (
                      header && <td key={colIndex} className="py-2 px-3 border">{row[colIndex] !== undefined ? String(row[colIndex]) : ''}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );

  const renderEmployeeTable = () => (
    <div className="overflow-x-auto">
      <div className="mb-4 flex justify-between items-center">
        <h2 className="text-lg font-semibold">Employee List</h2>
        <div className="flex space-x-3">
          <span className="text-sm py-1 px-2 bg-red-200 rounded">
            Non-Participating: {nonParticipatingCount}
          </span>
          <button 
            className="px-3 py-1 bg-green-500 text-white rounded hover:bg-green-600 text-sm"
            onClick={downloadExcelWithHighlighting}
          >
            Download Excel
          </button>
        </div>
      </div>
      <table className="min-w-full bg-white border">
        <thead className="bg-gray-100">
          <tr>
            <th className="py-2 px-4 border-b text-left">Name</th>
            <th className="py-2 px-4 border-b text-left">Function</th>
            <th className="py-2 px-4 border-b text-left">Manager</th>
            <th className="py-2 px-4 border-b text-right">Part-time %</th>
            <th className="py-2 px-4 border-b text-center">Present</th>
          </tr>
        </thead>
        <tbody>
          {data.map((employee, index) => (
            <tr key={index} className={isPresentClassName(employee.present)}>
              <td className="py-2 px-4 border-b">{employee.name}</td>
              <td className="py-2 px-4 border-b">{employee.function}</td>
              <td className="py-2 px-4 border-b">{employee.manager}</td>
              <td className="py-2 px-4 border-b text-right">{employee.partTimePercentage}%</td>
              <td className="py-2 px-4 border-b text-center">{formatPresentDisplay(employee.present)}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <div className="mt-4 bg-gray-100 p-3 rounded text-sm">
        <p>
          <span className="font-medium">Note:</span> Employees who won't participate are highlighted in red.
        </p>
      </div>
    </div>
  );

  const renderStatisticsTab = () => (
    <div className="flex flex-col space-y-6 mt-10">
      <div className="mb-4 flex justify-between items-center">
        <h2 className="text-lg font-semibold">Data Statistics</h2>
        <button 
          className="px-3 py-1 bg-green-500 text-white rounded hover:bg-green-600 text-sm"
          onClick={downloadExcelWithHighlighting}
        >
          Download Excel
        </button>
      </div>
      
      <div className="overflow-x-auto">
        <table className="min-w-full bg-white border">
          <thead className="bg-gray-100">
            <tr>
              <th className="py-2 px-4 border-b text-left">Manager</th>
              <th className="py-2 px-4 border-b text-right">Total Employees</th>
              <th className="py-2 px-4 border-b text-right">Present</th>
              <th className="py-2 px-4 border-b text-right">Absent</th>
              <th className="py-2 px-4 border-b text-right">Avg. Part-time %</th>
            </tr>
          </thead>
          <tbody>
            {Object.entries(stats).map(([manager, stat], index) => (
              <tr key={index} className={index % 2 === 0 ? 'bg-gray-50' : ''}>
                <td className="py-2 px-4 border-b">{manager}</td>
                <td className="py-2 px-4 border-b text-right">{stat.totalEmployees}</td>
                <td className="py-2 px-4 border-b text-right">{stat.presentEmployees}</td>
                <td className="py-2 px-4 border-b text-right">{stat.absentEmployees}</td>
                <td className="py-2 px-4 border-b text-right">{stat.avgPartTimePercentage}%</td>
              </tr>
            ))}
            <tr className="bg-blue-50 font-medium">
              <td className="py-2 px-4 border-b">Total</td>
              <td className="py-2 px-4 border-b text-right">
                {Object.values(stats).reduce((sum, stat) => sum + stat.totalEmployees, 0)}
              </td>
              <td className="py-2 px-4 border-b text-right">
                {Object.values(stats).reduce((sum, stat) => sum + stat.presentEmployees, 0)}
              </td>
              <td className="py-2 px-4 border-b text-right">
                {Object.values(stats).reduce((sum, stat) => sum + stat.absentEmployees, 0)}
              </td>
              <td className="py-2 px-4 border-b text-right">
                {(Object.values(stats).reduce((sum, stat) => 
                  sum + (parseFloat(stat.avgPartTimePercentage) * stat.totalEmployees), 0) / 
                 Object.values(stats).reduce((sum, stat) => sum + stat.totalEmployees, 0)).toFixed(2)}%
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      
      <div className="grid grid-cols-1 md:grid-cols-2 grid-rows-2 gap-6">
        {/* Pie chart for manager distribution */}
        <div className="bg-white p-4 rounded shadow">
          <h3 className="text-lg font-medium mb-4">PMC Distribution (Pie Chart)</h3>
          <div className="relative h-64">
            <svg width="100%" height="100%" viewBox="0 0 400 400">
              {(() => {
                const center = { x: 200, y: 200 };
                const radius = 200;
                const managerData = Object.entries(stats);
                const totalEmployees = managerData.reduce((sum, [_, stat]) => sum + stat.totalEmployees, 0);
                
                // Colors for pie slices
                const colors = ['#4299e1', '#48bb78', '#ed8936', '#9f7aea', '#f56565', '#38b2ac'];
                
                let startAngle = 0;
                let output = [];
                let legend = [];
                
                managerData.forEach(([manager, stat], index) => {
                  const percentage = stat.totalEmployees / totalEmployees;
                  const endAngle = startAngle + percentage * 2 * Math.PI;
                  
                  // Calculate pie slice path
                  const startX = center.x + radius * Math.cos(startAngle);
                  const startY = center.y + radius * Math.sin(startAngle);
                  const endX = center.x + radius * Math.cos(endAngle);
                  const endY = center.y + radius * Math.sin(endAngle);
                  
                  // Determine if the arc should be drawn as a large arc
                  const largeArcFlag = percentage > 0.5 ? 1 : 0;
                  
                  // Create path for pie slice
                  const pathData = [
                    `M ${center.x},${center.y}`,
                    `L ${startX},${startY}`,
                    `A ${radius},${radius} 0 ${largeArcFlag} 1 ${endX},${endY}`,
                    'Z'
                  ].join(' ');
                  
                  output.push(
                    <path 
                      key={`slice-${index}`}
                      d={pathData} 
                      fill={colors[index % colors.length]} 
                      stroke="#fff" 
                      strokeWidth="1"
                    />
                  );
                  
                  // Add label at the center of the slice
                  const labelAngle = startAngle + (endAngle - startAngle) / 2;
                  const labelRadius = radius * 0.7;
                  const labelX = center.x + labelRadius * Math.cos(labelAngle);
                  const labelY = center.y + labelRadius * Math.sin(labelAngle);
                  
                  if (percentage > 0.05) { // Only show label if slice is large enough
                    output.push(
                      <text 
                        key={`label-${index}`}
                        x={labelX} 
                        y={labelY} 
                        textAnchor="middle" 
                        dominantBaseline="middle"
                        fill="#fff"
                        fontWeight="bold"
                        fontSize="20"
                      >
                        {`${Math.round(percentage * 100)}%`}
                      </text>
                    );
                  }
                  
                  // Add to legend
                  legend.push({ manager, color: colors[index % colors.length], count: stat.totalEmployees });
                  
                  startAngle = endAngle;
                });
                
                // Add legend below the chart
                output.push(
                  <foreignObject key="legend" x="-200" y="200" width="400" height="300">
                    <div 
                      xmlns="http://www.w3.org/1999/xhtml" 
                      className="text-xs"
                    >
                      {legend.map(({ manager, color, count }, i) => (
                        <div key={i} className="flex items-center">
                          <div className="w-3 h-3 mr-1" style={{ backgroundColor: color }}></div>
                          <span className="truncate text-lg">{manager} ({count})</span>
                        </div>
                      ))}
                    </div>
                  </foreignObject>
                );
                
                return output;
              })()}
            </svg>
          </div>
        </div>
        
        {/* Bar chart for manager distribution */}
        <div className="bg-white p-4 rounded shadow">
          <h3 className="text-lg font-medium mb-4">Distribution</h3>
          <div className="flex items-end h-64 space-x-4">
            {Object.entries(stats).map(([manager, stat], index) => {
              // Fixed height - 50px per employee
              const heightPerEmployee = 5;
              const totalHeight = Math.min(stat.totalEmployees * heightPerEmployee, 250);
              
              return (
                <div key={index} className="flex flex-col items-center flex-1">
                  <div className="w-full flex justify-center mb-2">
                    <div className="flex flex-col items-center">
                      <div className="text-sm font-medium">{stat.totalEmployees}</div>
                      <div 
                        className="bg-blue-500 w-full" 
                        style={{height: `${totalHeight}px`, minHeight: '10px'}}
                      ></div>
                    </div>
                  </div>
                  <div className="text-xs text-center truncate w-22">{manager}</div>
                </div>
              );
            })}
          </div>
        </div>
        
        <div className="bg-white p-4 rounded shadow">
          <h3 className="text-lg font-medium mb-4">Present vs Absent by PMC</h3>
          <div className="flex items-end h-64 space-x-4">
            {Object.entries(stats).map(([manager, stat], index) => {
              // Fixed height scale
              const heightPerEmployee = 5;
              const presentHeight = stat.presentEmployees * heightPerEmployee;
              const absentHeight = stat.absentEmployees * heightPerEmployee;
              
              return (
                <div key={index} className="flex flex-col items-center flex-1">
                  <div className="w-full flex justify-center mb-2">
                    <div className="flex flex-col items-center w-full">
                      <div className="text-xs mb-1">{stat.presentEmployees}</div>
                      <div className="bg-green-500 w-full" style={{height: `${presentHeight}px`, minHeight: stat.presentEmployees ? '10px' : '0px'}}></div>
                      <div className="bg-red-500 w-full" style={{height: `${absentHeight}px`, minHeight: stat.absentEmployees ? '10px' : '0px'}}></div>
                      <div className="text-xs mt-1">{stat.absentEmployees}</div>
                    </div>
                  </div>
                  <div className="text-xs text-center truncate w-20">{manager}</div>
                </div>
              );
            })}
          </div>
          <div className="flex items-center justify-center mt-4 space-x-4">
            <div className="flex items-center">
              <div className="w-4 h-4 bg-green-500 mr-2"></div>
              <span className="text-sm">Present</span>
            </div>
            <div className="flex items-center">
              <div className="w-4 h-4 bg-red-500 mr-2"></div>
              <span className="text-sm">Absent</span>
            </div>
          </div>
        </div>

        <div className="bg-white p-4 rounded shadow">
          <h3 className="text-lg font-medium mb-4">Parttime Percentage by PMC</h3>
          <div className="flex items-end h-64 space-x-4">
            {Object.entries(stats).map(([manager, stat], index) => {
              // Scale height based on percentage (0-100%)
              const heightFactor = 1.5; // 2.5px per percentage point, max height for 100% = 250px
              const height = parseFloat(stat.avgPartTimePercentage) * heightFactor;

              return (
                <div key={index} className="flex flex-col items-center flex-1">
                  <div className="w-full flex justify-center mb-2">
                    <div className="flex flex-col items-center">
                      <div className="text-xs mb-1">{stat.avgPartTimePercentage}%</div>
                      <div 
                        className="bg-purple-500 w-full" 
                        style={{height: `${height}px`, minHeight: '10px'}}
                      ></div>
                    </div>
                  </div>
                  <div className="text-xs text-center truncate w-20">{manager}</div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );

  // Create horizontal bar chart to visualize present vs. absent percentages
  const renderHorizontalBarChart = () => {
    const totalPresentEmployees = Object.values(stats).reduce(
      (sum, stat) => sum + stat.presentEmployees,
      0
    );
    const totalAbsentEmployees = Object.values(stats).reduce(
      (sum, stat) => sum + stat.absentEmployees,
      0
    );
    const totalEmployees = totalPresentEmployees + totalAbsentEmployees;
    const presentPercentage = (totalPresentEmployees / totalEmployees * 100).toFixed(1);
    const absentPercentage = (totalAbsentEmployees / totalEmployees * 100).toFixed(1);

    return (
      <div className="bg-white p-4 rounded shadow mt-6">
        <h3 className="text-lg font-medium mb-4">Overall Participation Rate</h3>
        <div className="relative pt-1">
          <div className="flex items-center justify-between mb-2">
            <div>
              <span className="text-xs font-semibold inline-block text-green-600">
                Present: {totalPresentEmployees} ({presentPercentage}%)
              </span>
            </div>
            <div>
              <span className="text-xs font-semibold inline-block text-red-600">
                Absent: {totalAbsentEmployees} ({absentPercentage}%)
              </span>
            </div>
          </div>
          <div className="flex h-6 mb-4 overflow-hidden rounded-lg bg-gray-200">
            <div
              style={{ width: `${presentPercentage}%` }}
              className="flex flex-col justify-center text-center text-white bg-green-500 shadow-none whitespace-nowrap"
            >
              {presentPercentage > 5 && `${presentPercentage}%`}
            </div>
            <div
              style={{ width: `${absentPercentage}%` }}
              className="flex flex-col justify-center text-center text-white bg-red-500 shadow-none whitespace-nowrap"
            >
              {absentPercentage > 5 && `${absentPercentage}%`}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderDeploymentInstructions = () => (
    <div className="bg-white p-4 rounded shadow mt-8">
      <h2 className="text-lg font-semibold mb-3">GitHub Pages Deployment Instructions</h2>
      <ol className="list-decimal pl-5 space-y-2">
        <li>Create a new GitHub repository</li>
        <li>Initialize your project with Create React App or Next.js</li>
        <li>
          Install required dependencies:
          <pre className="bg-gray-100 p-2 rounded mt-1 text-sm overflow-x-auto">
            npm install xlsx lodash gh-pages
          </pre>
        </li>
        <li>Copy this component into your project</li>
        <li>Set up GitHub Pages deployment in your repository settings</li>
        <li>
          Deploy your application with:
          <pre className="bg-gray-100 p-2 rounded mt-1 text-sm overflow-x-auto">
            npm run build
            npm run deploy
          </pre>
        </li>
      </ol>
      <p className="mt-4 text-sm">
        <strong>Note:</strong> This application processes files entirely in the browser - no server is needed and no data is sent anywhere.
      </p>
    </div>
  );

  return (
    <div className="max-w-6xl mx-auto p-4">
      <h1 className="text-2xl font-bold mb-6">Employee Data Analysis</h1>

      {loading ? (
        <div className="flex justify-center items-center h-64 bg-white rounded shadow">
          <div className="text-center">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-gray-900 mb-2"></div>
            <p className="text-lg">Processing your file...</p>
          </div>
        </div>
      ) : error ? (
        <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
          <p className="font-medium">Error:</p>
          <p>{error}</p>
          <button 
            className="mt-3 px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
            onClick={() => {
              setError(null);
              setActiveTab('upload');
              setFile(null);
              setFileStructure(null);
              setData([]);
              setStats({});
            }}
          >
            Try Again
          </button>
        </div>
      ) : fileStructure && data.length > 0 ? (
        <div>
          <div className="flex mb-4 border-b">
            <button 
              className={`py-2 px-4 font-medium ${activeTab === 'structure' ? 'border-b-2 border-blue-500 text-blue-500' : 'text-gray-500'}`}
              onClick={() => setActiveTab('structure')}
            >
              File Structure
            </button>
            <button 
              className={`py-2 px-4 font-medium ${activeTab === 'employees' ? 'border-b-2 border-blue-500 text-blue-500' : 'text-gray-500'}`}
              onClick={() => setActiveTab('employees')}
            >
              Employees ({data.length})
            </button>
            <button 
              className={`py-2 px-4 font-medium ${activeTab === 'statistics' ? 'border-b-2 border-blue-500 text-blue-500' : 'text-gray-500'}`}
              onClick={() => setActiveTab('statistics')}
            >
              Statistics
            </button>
            <button 
              className={`py-2 px-4 font-medium ${activeTab === 'deploy' ? 'border-b-2 border-blue-500 text-blue-500' : 'text-gray-500'}`}
              onClick={() => setActiveTab('deploy')}
            >
              Deployment
            </button>
          </div>

          {activeTab === 'structure' ? renderFileStructure() 
          : activeTab === 'employees' ? renderEmployeeTable() 
          : activeTab === 'statistics' ? (
            <>
              {renderHorizontalBarChart()}
              {renderStatisticsTab()}
            </>
          ) : renderDeploymentInstructions()}

          {activeTab !== 'deploy' && activeTab !== 'structure' && (
            <div className="mt-4 bg-gray-100 p-3 rounded text-sm">
              <p>
                <span className="font-medium">Note:</span> Employees who won't participate are highlighted in red.
                Part-time percentages have been fixed and properly displayed.
              </p>
            </div>
          )}
        </div>
      ) : fileStructure ? (
        <div>
          {renderFileStructure()}
        </div>
      ) : (
        renderUploadTab()
      )}
    </div>
  );
};

// Add GitHub Corners for a "Fork me on GitHub" button
const GitHubCorner = () => (
  <a href="https://github.com/YOUR_USERNAME/employee-excel-viewer" 
    className="github-corner" 
    aria-label="View source on GitHub"
    target="_blank"
    rel="noopener noreferrer"
    style={{
      position: 'absolute',
      top: 0,
      right: 0,
      border: 0,
      zIndex: 10
    }}>
    <svg width="80" height="80" viewBox="0 0 250 250" style={{
      fill: '#151513', 
      color: '#fff', 
      position: 'absolute', 
      top: 0, 
      border: 0, 
      right: 0
    }} aria-hidden="true">
      <path d="M0,0 L115,115 L130,115 L142,142 L250,250 L250,0 Z"></path>
      <path d="M128.3,109.0 C113.8,99.7 119.0,89.6 119.0,89.6 C122.0,82.7 120.5,78.6 120.5,78.6 C119.2,72.0 123.4,76.3 123.4,76.3 C127.3,80.9 125.5,87.3 125.5,87.3 C122.9,97.6 130.6,101.9 134.4,103.2" fill="currentColor" style={{transformOrigin: '130px 106px'}} className="octo-arm"></path>
      <path d="M115.0,115.0 C114.9,115.1 118.7,116.5 119.8,115.4 L133.7,101.6 C136.9,99.2 139.9,98.4 142.2,98.6 C133.8,88.0 127.5,74.4 143.8,58.0 C148.5,53.4 154.0,51.2 159.7,51.0 C160.3,49.4 163.2,43.6 171.4,40.1 C171.4,40.1 176.1,42.5 178.8,56.2 C183.1,58.6 187.2,61.8 190.9,65.4 C194.5,69.0 197.7,73.2 200.1,77.6 C213.8,80.2 216.3,84.9 216.3,84.9 C212.7,93.1 206.9,96.0 205.4,96.6 C205.1,102.4 203.0,107.8 198.3,112.5 C181.9,128.9 168.3,122.5 157.7,114.1 C157.9,116.9 156.7,120.9 152.7,124.9 L141.0,136.5 C139.8,137.7 141.6,141.9 141.8,141.8 Z" fill="currentColor" className="octo-body"></path>
    </svg>
  </a>
);

// Create our App component with GitHub corner
const App = () => {
  return (
    <div className="relative">
      <GitHubCorner />
      <EmployeeDataVisualization />
    </div>
  );
};

export default App;