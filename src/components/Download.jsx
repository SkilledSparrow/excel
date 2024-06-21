import React, { useState, useEffect } from 'react';
import { saveAs } from 'file-saver';
import ExcelJS from 'exceljs';
import data from '../assets/data.json';
import './Download.css'; // Import your CSS file

const Download = () => {
  const [selectedColumns, setSelectedColumns] = useState([]);
  const [dropdownVisible, setDropdownVisible] = useState(false);
  const [selectAll, setSelectAll] = useState(false);
  const [conditionalFormatting, setConditionalFormatting] = useState({
    Status: false,
    Score: false,
    'Additional feedback': false,
  });

  const headers = Object.keys(data[0]);

  useEffect(() => {
    if (selectAll) {
      setSelectedColumns(headers);
    } else if (selectedColumns.length === headers.length) {
      setSelectedColumns([]); // Clear all selected columns if deselecting "Select All"
    }
  }, [selectAll, headers, selectedColumns.length]);

  const handleColumnChange = (event) => {
    const { value, checked } = event.target;
    setSelectedColumns(prev =>
      checked ? [...prev, value] : prev.filter(col => col !== value)
    );
    if (!checked) {
      setSelectAll(false);
      if (value in conditionalFormatting) {
        setConditionalFormatting(prev => ({
          ...prev,
          [value]: false,
        }));
      }
    }
  };

  const handleConditionalFormatChange = (event) => {
    const { value, checked } = event.target;
    setConditionalFormatting(prev => ({
      ...prev,
      [value]: checked,
    }));
  };

  const applyConditionalFormatting = (cell, header, item) => {
    if (header === 'Status' && conditionalFormatting.Status) {
      const status = item[header].toLowerCase(); // Normalize status value

      if (status === 'open') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '99FF99' } // Green background
        };
      } else if (status === 'closed') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF9999' } // Red background
        };
      } else {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFF99' } // Yellow background
        };
      }
    } else if (header === 'Score' && conditionalFormatting.Score) {
      const score = parseFloat(item[header]);

      if (!isNaN(score)) {
        if (score > 8) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '99FF99' } // Green background
          };
        } else if (score > 6) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF99' } // Yellow background
          };
        } else {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF9999' } // Red background
          };
        }
      }
    } else if (header === 'Additional feedback' && conditionalFormatting['Additional feedback']) {
      // Example conditional formatting for 'Additional feedback'
      // Add specific conditional formatting logic here
      const feedback = item[header];
      if (feedback.includes('Yes')) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '99FF99' } // Green background
        };
      } else if (feedback.includes('No')) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF9999' } // Red background
        };
      }
    }
  };

  const handleDownload = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    const filteredHeaders = headers.filter(header => selectedColumns.includes(header));

    // Add title to cell B2
    const titleCell = worksheet.getCell('B2');
    titleCell.value = 'Data Export';
    titleCell.font = { size: 18, bold: true };

    filteredHeaders.forEach((header, index) => {
      const column = worksheet.getColumn(index + 2);
      column.width = header.length > 10 ? header.length * 1.3 : 20; // Adjust column width

      const cell = worksheet.getCell(4, index + 2);
      cell.value = header;
      cell.style.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9EAD3' }, // Light blue background
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // Add the data rows starting from row 5 and column B
    data.forEach((item, rowIndex) => {
      filteredHeaders.forEach((header, colIndex) => {
        const cell = worksheet.getCell(rowIndex + 5, colIndex + 2);
        cell.value = item[header];
        cell.alignment = { wrapText: true, vertical: 'top' };

        applyConditionalFormatting(cell, header, item);

        const wordsCount = String(item[header]).split(' ').length;
        if (wordsCount > 10) {
          cell.alignment = { wrapText: true };
          // Double the column width if more than 10 words
          const currentWidth = worksheet.getColumn(colIndex + 2).width;
          worksheet.getColumn(colIndex + 2).width = currentWidth * 1.2;
        }

        if (header === 'Attributes') {
          cell.value = String(item[header]).split(',').join(',\n');
        }

        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });

    worksheet.views = [{ showGridLines: false }];

    // Save the file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'data.xlsx');
  };

  return (
    <div className="download-container">
      <h2 className="download-header">Excel Data Export</h2>
      <div className="download-buttons">
        <button onClick={() => setDropdownVisible(!dropdownVisible)}>
          {dropdownVisible ? 'Hide Columns' : 'Select Columns'}
        </button>
        <button onClick={handleDownload} className="download-button">
          Download Excel
        </button>
      </div>
      {dropdownVisible && (
        <div className="column-selector">
          <div>
            <input
              type="checkbox"
              checked={selectAll}
              onChange={() => setSelectAll(!selectAll)}
            />
            <label>Select All</label>
          </div>
          {headers.map((column) => (
            <div key={column} className="checkbox-container">
              <input
                type="checkbox"
                value={column}
                checked={selectedColumns.includes(column)}
                onChange={handleColumnChange}
              />
              <label>{column}</label>
              {['Status', 'Score', 'Additional feedback'].includes(column) && (
                <>
                  <input
                    type="checkbox"
                    value={column}
                    checked={conditionalFormatting[column]}
                    onChange={handleConditionalFormatChange}
                  />
                  <label className="conditional-format-label">Conditional Formatting</label>
                </>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default Download;


