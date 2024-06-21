import React from 'react';

const DropdownMenu = ({ headers, onSelectAll, onColumnChange, selectedColumns, selectAll }) => {
  const handleSelectAll = () => {
    onSelectAll(!selectAll);
  };

  const handleColumnChange = (event) => {
    const { value, checked } = event.target;
    onColumnChange(value, checked);
  };

  return (
    <div>
      <div>
        <input
          type="checkbox"
          id="selectAll"
          checked={selectAll}
          onChange={handleSelectAll}
        />
        <label htmlFor="selectAll">Select All</label>
      </div>
      {headers.map((header) => (
        <div key={header}>
          <input
            type="checkbox"
            id={header}
            value={header}
            checked={selectedColumns.includes(header)}
            onChange={handleColumnChange}
          />
          <label htmlFor={header}>{header}</label>
        </div>
      ))}
    </div>
  );
};

export default DropdownMenu;
