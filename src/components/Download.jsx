import React, { useState, useEffect } from 'react';
import { saveAs } from 'file-saver';
import ExcelJS from 'exceljs';
import data from '../assets/data.json';
import './Download.css'; // CSS file

const Download = () => {
  const [selectedColumns, setSelectedColumns] = useState([]);
  const [dropdownVisible, setDropdownVisible] = useState(false);
  const [conditionalFormattingVisible, setConditionalFormattingVisible] = useState(false);
  const [selectAll, setSelectAll] = useState(false);
  const [conditionalFormattingRules, setConditionalFormattingRules] = useState([]);

  const headers = Object.keys(data[0]);

  useEffect(() => {
    if (selectAll) {
      setSelectedColumns(headers);
    } else if (selectedColumns.length === headers.length) {
      setSelectedColumns([]); // Clear all selected columns if deselecting "Select All". Convert it into just deselecting that specific column
    }
  }, [selectAll, headers, selectedColumns.length]);

  const handleColumnChange = (event) => {
    const { value, checked } = event.target;
    setSelectedColumns(prev =>
      checked ? [...prev, value] : prev.filter(col => col !== value)
    );
    if (!checked) {
      setSelectAll(false);
    }
  };

  const handleAddConditionalFormattingRule = () => {
    setConditionalFormattingRules(prev => [...prev, { column: '', condition: '', color: '' }]);
  };

  const handleConditionalFormattingRuleChange = (index, field, value) => {
    const newRules = [...conditionalFormattingRules];
    newRules[index][field] = value;
    setConditionalFormattingRules(newRules);
  };

  const evaluateCondition = (value, condition) => {
    console.log(`Evaluating condition: ${condition} for value: ${value}`);
    const operators = {
      '>': (a, b) => a > b,
      '<': (a, b) => a < b,
      '>=': (a, b) => a >= b,
      '<=': (a, b) => a <= b,
      '==': (a, b) => a == b,
      '===': (a, b) => a === b,
      '!=': (a, b) => a != b,
      '!==': (a, b) => a !== b
    };

    const match = condition.match(/([><=!]+)\s*(\d+)/);
    if (match) {
      const [, operator, threshold] = match;
      const thresholdNumber = parseFloat(threshold);

      if (operators[operator]) {
        return operators[operator](value, thresholdNumber);
      }
    }

    return false;
  };

  const applyConditionalFormatting = (cell, header, item) => {
    conditionalFormattingRules.forEach(rule => {
      if (rule.column === header) {
        const condition = rule.condition;
        const color = rule.color;

        if (evaluateCondition(item[header], condition)) {
          console.log(`Applying color: ${color} to cell with value: ${item[header]} for header: ${header}`);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color }
          };
        }
      }
    });
  };

  const handleDownload = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Escalations');

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

        // if (header === 'Attributes') {
        //   cell.value = String(item[header]).split(',').join(',\n');
        // }

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
    saveAs(blob, 'escalations.xlsx');
  };

  return (
    <div className="download-container">
      <h2 className="download-header">Excel Data Export</h2>
      <div className="download-buttons">
        <button onClick={() => setDropdownVisible(!dropdownVisible)}>
          {dropdownVisible ? 'Hide Columns' : 'Select Columns'}
        </button>
        <button onClick={() => setConditionalFormattingVisible(!conditionalFormattingVisible)}>
          {conditionalFormattingVisible ? 'Hide Conditional Formatting' : 'Conditional Formatting'}
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
            </div>
          ))}
        </div>
      )}
      {conditionalFormattingVisible && (
        <div className="conditional-formatting">
          {conditionalFormattingRules.map((rule, index) => (
            <div key={index} className="conditional-formatting-rule">
              <select
                value={rule.column}
                onChange={(e) => handleConditionalFormattingRuleChange(index, 'column', e.target.value)}
              >
                <option value="">Select Column</option>
                {headers.map((header) => (
                  <option key={header} value={header}>{header}</option>
                ))}
              </select>
              <input
                type="text"
                placeholder="Condition (e.g., >8)"
                value={rule.condition}
                onChange={(e) => handleConditionalFormattingRuleChange(index, 'condition', e.target.value)}
              />
              <input
                type="text"
                placeholder="Color (e.g., FFFFFF)"
                value={rule.color}
                onChange={(e) => handleConditionalFormattingRuleChange(index, 'color', e.target.value)}
              />
            </div>
          ))}
          {conditionalFormattingRules.length < 10 && (
            <button onClick={handleAddConditionalFormattingRule}>Add Rule</button>
          )}
        </div>
      )}
    </div>
  );
};

export default Download;
