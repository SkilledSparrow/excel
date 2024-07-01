import React, { useState, useEffect, useMemo, useRef } from 'react';
import { saveAs } from 'file-saver';
import ExcelJS from 'exceljs';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';
import data from '../assets/data.json';
import './Download.css'; // CSS file

const Download = () => {
  const [rowData, setRowData] = useState([]);
  const [floatingFilterVisible, setFloatingFilterVisible] = useState(false);
  const [columnDefs, setColumnDefs] = useState([]);
  const gridApiRef = useRef(null);
  const [columnOrder, setColumnOrder] = useState([]);
  const [selectedColumns, setSelectedColumns] = useState([]);
  const [dropdownVisible, setDropdownVisible] = useState(false);

  useEffect(() => {
    setRowData(data);
    const headers = Object.keys(data[0]);
    const initialColumnDefs = headers.map((header, index) => ({
      headerName: header,
      field: header,
      editable: false,
      sortable: true,
      filter: true,
      floatingFilter: floatingFilterVisible,
      resizable: true,
      headerCheckboxSelection: index === 0,
    }));
    setColumnDefs(initialColumnDefs);
    setColumnOrder(headers);
  }, [floatingFilterVisible]);

  const handleDownloadExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Escalations');

    const columnsToDownload = selectedColumns.filter(column => columnOrder.includes(column));



    columnsToDownload.forEach((header, index) => {
      const column = worksheet.getColumn(index + 1);
      column.width = header.length > 10 ? header.length * 1.3 : 20;

      const cell = worksheet.getCell(1, index + 1);
      cell.value = header;
      cell.style.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9EAD3' },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    rowData.forEach((item, rowIndex) => {
      columnsToDownload.forEach((column, colIndex) => {
        const cell = worksheet.getCell(rowIndex + 2, colIndex + 1);
        cell.value = item[column];
        cell.alignment = { wrapText: true, vertical: 'top' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
    worksheet.views = [{ showGridLines: false }];
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    saveAs(blob, 'escalations.xlsx');
  };

  const handleDownloadCSV = () => {
    const columnsToDownload = selectedColumns.filter(column => columnOrder.includes(column));
    const csvContent = [
      columnsToDownload.join(','), // Add headers row

      ...rowData.map(item => columnsToDownload.map(column => item[column]).join(',')) // Add data rows
    ].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    saveAs(blob, 'escalations.csv');
  };

  const toggleFloatingFilter = () => {
    setFloatingFilterVisible(!floatingFilterVisible);
  };

  const onGridReady = params => {
    gridApiRef.current = params.api;
  };

  const onColumnMoved = event => {
    const newColumnOrder = event.columnApi.getAllDisplayedColumns().map(col => col.getColId());
    setColumnOrder(newColumnOrder);
  };
  const handleColumnSelectionChange = (event) => {

    const { value, checked } = event.target;

    setSelectedColumns(prevSelectedColumns =>

      checked ? [...prevSelectedColumns, value] : prevSelectedColumns.filter(column => column !== value)

    );

  };



  const toggleDropdown = () => {

    setDropdownVisible(!dropdownVisible);

  };



  useEffect(() => {

    const handleClickOutside = (event) => {

      if (dropdownVisible && !event.target.closest('.dropdown')) {

        setDropdownVisible(false);

      }

    };

    document.addEventListener('click', handleClickOutside);

    return () => {

      document.removeEventListener('click', handleClickOutside);

    };

  }, [dropdownVisible]);


  return (
    <div className="download-container">
      <h2 className="download-header">Excel Data Export</h2>
      <div className="download-buttons">
        <button onClick={handleDownloadExcel} className="download-button">
          Download Excel
        </button>
        <button onClick={handleDownloadCSV} className="download-button">
          Download CSV
        </button>
        <div className="dropdown">

<button onClick={toggleDropdown} className="dropdown-button">

  Select Columns

</button>

{dropdownVisible && (

  <div className="dropdown-content">

    {columnOrder.map(column => (

      <div key={column} className="dropdown-item">

        <input
          type="checkbox"
          id={column}
          value={column}
          checked={selectedColumns.includes(column)}
          onChange={handleColumnSelectionChange}
        />
        <label htmlFor={column}>{column}</label>
      </div>
    ))}

  </div>

)}

</div>
      </div>
      <div className="ag-theme-alpine" style={{ height: 400, width: '100%' }}>
        <AgGridReact
          columnDefs={columnDefs}
          rowData={rowData}
          deltaRowDataMode={true}
          getRowId={params => params.data.id} // Assuming each row data has a unique 'id' field
          // rowSelection="multiple"
          animateRows={false}
          enableCellTextSelection={false}
          suppressContextMenu={true}
          suppressCellSelection={true}
          onGridReady={onGridReady}
          onColumnMoved={onColumnMoved}
        />
        <button onClick={toggleFloatingFilter} className="toggle-filter-button">
          {floatingFilterVisible ? 'Hide Filters' : 'Show Filters'}
        </button>
      </div>
    </div>
  );
};

export default Download;